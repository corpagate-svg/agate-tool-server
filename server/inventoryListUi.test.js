const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const test = require("node:test");
const assert = require("node:assert/strict");

const source = fs.readFileSync(path.join(__dirname, "server.js"), "utf8");

test("在庫一覧の表示列順が指定どおりで、売却個数合計をExcel列へ追加しない", () => {
  const expected = [
    "商品ID", "商品名", "リアル在庫", "在庫数(現物)", "出品国数", "US_出品ID",
    "仕入価格(円)", "US価格(USD)", "仕入日", "売却個数 合計",
    "US累計売却数", "UK累計売却数", "AU累計売却数", "備考",
  ];
  const declaration = /const INV_TABLE_HEADERS = \[([\s\S]*?)\];/.exec(source);
  assert.ok(declaration, "在庫一覧の表示列定義が見つかりません");
  const headers = [...declaration[1].matchAll(/"([^"]+)"|INV_SOLD_TOTAL/g)]
    .map((match) => match[1] || "売却個数 合計");
  assert.deepEqual(headers, expected);

  const inventoryHeaders = /const INV_HEADERS = \[([\s\S]*?)\];/.exec(source);
  assert.ok(inventoryHeaders);
  assert.doesNotMatch(inventoryHeaders[1], /売却個数 合計/);
});

test("売却個数合計はUS・UK・AUの累計を足し、空欄を0として扱う", () => {
  const functionSource = /function inventorySoldTotal\(row\) \{[\s\S]*?\n\}/.exec(source);
  assert.ok(functionSource, "売却個数合計関数が見つかりません");
  const context = {
    invHeaders: ["US累計売却数", "UK累計売却数", "AU累計売却数"],
  };
  vm.runInNewContext(`${functionSource[0]}; resultA = inventorySoldTotal([2, 3, 4]); resultB = inventorySoldTotal([5, null, ""]);`, context);
  assert.equal(context.resultA, 9);
  assert.equal(context.resultB, 5);
});

test("売却個数合計は表示専用かつソート可能で、各国累計列も維持する", () => {
  assert.match(source, /if \(idx === INV_SOLD_TOTAL\)/);
  assert.match(source, /inventorySoldTotal\(a\) - inventorySoldTotal\(b\)/);
  assert.match(source, /if \(h === INV_SOLD_TOTAL\)/);
  for (const header of ["US累計売却数", "UK累計売却数", "AU累計売却数"]) {
    assert.ok(source.includes(`"${header}"`));
  }
});

test("在庫一覧のチェックボックス列だけを少し狭くする", () => {
  assert.match(source, /#inv-table-scroll td\.checkbox-col, #inv-table-scroll th\.checkbox-col \{ width: 26px; min-width: 26px; padding-left: 6px; \}/);
});

test("一時的な注文parser診断表示が残っていない", () => {
  const parserSource = fs.readFileSync(path.join(__dirname, "orderParser.js"), "utf8");
  for (const marker of ["PARSER_BUILD_ID", "ne-diag", "neParseDiagCallCount", "[診断] Parser Build", "診断専用"]) {
    assert.equal(source.includes(marker) || parserSource.includes(marker), false, `${marker} が残っています`);
  }
});
