const test = require("node:test");
const assert = require("node:assert/strict");
const { parseOrderText } = require("./orderParser");

function productBlock({ title = "Test Product", itemId = "123456789012", quantity = 1, price = "$4.98", total = "$9.96" } = {}) {
  return [
    title,
    "独自のラベル（SKU）: TEST-SKU",
    `商品ID: ${itemId}`,
    "数量",
    String(quantity),
    "（在庫2点）",
    "商品価格",
    price,
    "商品合計",
    total,
  ].join("\n");
}

function orderText({ block = productBlock(), summary = [], subtotalPoints = 1 } = {}) {
  return [
    "商品",
    block,
    "注文",
    "25-15080-33365",
    "販売",
    "Sep 5, 2026",
    subtotalPoints === null ? "小計" : `小計（${subtotalPoints}点）`,
    ...summary,
    "お客様の収益",
    "注文の収益",
    "US $13.48",
  ].join("\n");
}

test("実US注文形式: 弱い$表記とUS $の強シグナルからUSと判定する", () => {
  const parsed = parseOrderText(orderText({
    block: productBlock({ quantity: 2, price: "$4.98", total: "$9.96" }),
    subtotalPoints: 2,
    summary: ["US $9.96", "注文の合計金額", "US $16.46"],
  }));
  assert.equal(parsed.site, "US");
  assert.equal(parsed.parseStatus, "OK");
});

test("既存UK注文のGB £判定を維持する", () => {
  const parsed = parseOrderText(orderText({
    block: productBlock({ price: "GB £6.51", total: "GB £6.51" }),
    summary: ["GB £6.51", "注文の合計金額", "GB £13.57"],
  }));
  assert.equal(parsed.site, "UK");
  assert.equal(parsed.parseStatus, "OK");
});

test("Dandadan UK単品注文は商品ブロック数fallbackでOKになる", () => {
  const parsed = parseOrderText(orderText({
    block: productBlock({ title: "Dandadan Jiji Big Acrylic Stand", itemId: "137173536640", price: "GB £6.51", total: "GB £6.51" }),
    subtotalPoints: null,
    summary: ["GB £6.51", "注文の合計金額", "GB £13.57"],
  }));
  assert.equal(parsed.site, "UK");
  assert.equal(parsed.subtotalCheckMethod, "商品ブロック数");
  assert.equal(parsed.parseStatus, "OK");
});

test("UK数量4注文は小計点数で照合してOKになる", () => {
  const parsed = parseOrderText(orderText({
    block: productBlock({ quantity: 4, price: "GB £6.51", total: "GB £26.04" }),
    subtotalPoints: 4,
    summary: ["GB £26.04", "注文の合計金額", "GB £30.84"],
  }));
  assert.equal(parsed.site, "UK");
  assert.equal(parsed.quantityTotal, 4);
  assert.equal(parsed.subtotalCheckMethod, "小計点数");
  assert.equal(parsed.parseStatus, "OK");
});

test("AU $の強シグナルだけならAUと判定する", () => {
  const parsed = parseOrderText(orderText({
    block: productBlock({ price: "AU $15.00", total: "AU $15.00" }),
    summary: ["AU $15.00", "注文の合計金額", "AU $20.00"],
  }));
  assert.equal(parsed.site, "AU");
  assert.equal(parsed.parseStatus, "OK");
});

test("商品価格が$単体でも支払い側AU $からAUと判定する", () => {
  const parsed = parseOrderText(orderText({
    block: productBlock({ price: "$10.00", total: "$10.00" }),
    summary: ["AU $10.00", "注文の合計金額", "AU $15.00"],
  }));
  assert.equal(parsed.site, "AU");
  assert.equal(parsed.parseStatus, "OK");
});

test("US $とGB £が金額ラベル直後に混在すれば要確認", () => {
  const parsed = parseOrderText(orderText({
    block: productBlock({ price: "US $10.00", total: "GB £10.00" }),
    summary: ["US $10.00"],
  }));
  assert.equal(parsed.site, "要確認");
  assert.equal(parsed.parseStatus, "要確認");
});

test("$単体しかない場合は要確認", () => {
  const parsed = parseOrderText(orderText({
    block: productBlock({ price: "$10.00", total: "$10.00" }),
    summary: ["$10.00", "注文の合計金額", "$15.00"],
  }));
  assert.equal(parsed.site, "要確認");
  assert.equal(parsed.parseStatus, "要確認");
});

test("US $とAU $が金額ラベル直後に混在すれば要確認", () => {
  const parsed = parseOrderText(orderText({
    block: productBlock({ price: "US $10.00", total: "AU $10.00" }),
    summary: ["US $10.00"],
  }));
  assert.equal(parsed.site, "要確認");
  assert.equal(parsed.parseStatus, "要確認");
});
