// 在庫管理表(在庫数(現物)・US/UK/AU出品ID・累計売却数など)の再構築エンジン。
// server.jsのhandleRebuildInventory(CSVアップロード)から移設したロジックで、動作は変更していません。
// eBay Trading API経由の再構築(ebay/inventorySync.js)からも共通で利用します。
const ExcelJS = require("exceljs");
const { withInventoryLock, atomicWriteWorkbook } = require("./inventoryLock");
const { copyProtectedSheets } = require("./inventoryProtectedSheets");

const INV_HEADER_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F4E78" } };
const INV_HEADER_FONT = { bold: true, color: { argb: "FFFFFFFF" } };
const INV_FLAG_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF2CC" } };
const INV_REMOVED_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF2DCDB" } };

function normText(s) {
  return String(s || "").trim().replace(/\s+/g, " ");
}

const SHIPPING_NOTE_SUFFIXES = [
  "【Extra Items Ship FREE】",
  "【FlaExtra Items Ship FREE】",
  "【Flat S/H】",
  "(Flat rate after first)",
];
function stripShippingNote(title) {
  let t = String(title || "").trim();
  let changed = true;
  while (changed) {
    changed = false;
    for (const suf of SHIPPING_NOTE_SUFFIXES) {
      if (t.endsWith(suf)) {
        t = t.slice(0, -suf.length).trim();
        changed = true;
      }
    }
  }
  return t;
}
function nameKey(title, variation) {
  return normText(stripShippingNote(title)) + " || " + normText(variation);
}

function numOrNullCsv(v) {
  const n = parseFloat(v);
  return Number.isFinite(n) ? n : null;
}

function intOrNullCsv(v) {
  const n = parseInt(v, 10);
  return Number.isFinite(n) ? n : null;
}

const MONTH_ABBR_EN = { Jan: 1, Feb: 2, Mar: 3, Apr: 4, May: 5, Jun: 6, Jul: 7, Aug: 8, Sep: 9, Oct: 10, Nov: 11, Dec: 12 };
function parseEbayStartDate(s) {
  const m = /^([A-Za-z]{3})-(\d{1,2})-(\d{2})/.exec(String(s || ""));
  if (!m) return null;
  const mm = MONTH_ABBR_EN[m[1]];
  if (!mm) return null;
  const yyyy = 2000 + parseInt(m[3], 10);
  return `${yyyy}-${String(mm).padStart(2, "0")}-${String(m[2]).padStart(2, "0")}`;
}

// records: [{ Title, "Listing site"(US/UK/AU), "Item number", "Variation details", "Available quantity", "Current price", "Sold quantity", "Custom label (SKU)", "Start date"(MMM-DD-YY), Currency }]
// 最初の読み込みから最後の書き込みまでの間に他の更新が割り込むと、その更新が黙って
// 消えてしまう(lost update)ため、この関数全体を在庫管理表.xlsxの共有ロックで直列化する。
function rebuildInventoryFromRecords(args) {
  return withInventoryLock(() => rebuildInventoryFromRecordsLocked(args));
}

async function rebuildInventoryFromRecordsLocked({ records, sourceNote, removedNoteLabel, INV_HEADERS, INVENTORY_PATH, loadInventoryWorkbook }) {
  const oldWb = await loadInventoryWorkbook();
  const oldWs = oldWb.getWorksheet("在庫管理表") || oldWb.worksheets[0];
  const saved = new Map();
  let maxId = 0;
  oldWs.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const pid = row.getCell(1).value;
    if (!pid) return;
    const idNum = parseInt(String(pid).slice(1), 10);
    if (Number.isFinite(idNum)) maxId = Math.max(maxId, idNum);
    const fullRow = [];
    for (let c = 1; c <= INV_HEADERS.length; c++) fullRow.push(row.getCell(c).value);
    const key = nameKey(fullRow[1], fullRow[2]);
    saved.set(key, {
      商品ID: pid,
      fullRow,
      wasRemoved: String(fullRow[18] || "").includes("見当たりません"),
    });
  });

  const groups = new Map();
  records.forEach((r) => {
    const key = nameKey(r["Title"], r["Variation details"]);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(r);
  });

  const anomalous = [];
  const mainKeys = [];
  for (const [key, items] of groups) {
    if (items.length > 3) anomalous.push(...items);
    else mainKeys.push(key);
  }

  const today = new Date().toISOString().slice(0, 10);
  const rowsOut = [];
  const newProducts = [];
  const noUsSkipped = [];
  let nextId = maxId + 1;
  const newKeySet = new Set();

  for (const key of mainKeys) {
    const items = groups.get(key);
    const qtys = new Set(items.map((it) => it["Available quantity"]));
    const bySite = {};
    items.forEach((it) => {
      const site = it["Listing site"];
      if (["US", "UK", "AU"].includes(site) && !bySite[site]) bySite[site] = it;
    });

    // USに出品していないものは商品リストに載せない(USを基準にUK/AUへ転送している運用のため)
    if (!bySite.US) {
      noUsSkipped.push(...items);
      continue;
    }
    newKeySet.add(key);
    const anchor = bySite.US;
    const siteCount = Object.keys(bySite).length;

    let pid;
    let purchasePrice = null, purchaseDate = null, purchaseFrom = null, note = "";
    let realStock = null, realStockConfirmedAt = null, stocktakeQty = null, stocktakeAt = null;
    let imageUrl = null, jaName = null, stocktakeChecked = null;
    if (saved.has(key)) {
      const s = saved.get(key);
      pid = s.商品ID;
      purchasePrice = s.fullRow[15]; purchaseDate = s.fullRow[16]; purchaseFrom = s.fullRow[17];
      note = s.wasRemoved ? "" : (s.fullRow[18] || "");
      realStock = s.fullRow[21] ?? null;
      realStockConfirmedAt = s.fullRow[22] ?? null;
      stocktakeQty = s.fullRow[23] ?? null;
      stocktakeAt = s.fullRow[24] ?? null;
      imageUrl = s.fullRow[25] ?? null;
      jaName = s.fullRow[26] ?? null;
      stocktakeChecked = s.fullRow[28] ?? null;
    } else {
      pid = "P" + String(nextId).padStart(4, "0");
      nextId++;
      newProducts.push([pid, anchor["Title"]]);
    }
    if (!purchaseDate) purchaseDate = parseEbayStartDate(anchor["Start date"]);
    const usQty = intOrNullCsv(anchor["Available quantity"]);
    // リアル在庫は「当社が実際に保有する数量」の独自管理値。CSV再取込では上書きしない。
    // ただし未設定(導入前・移行直後)の場合のみ、初期値としてeBayのUS在庫数をコピーする。
    if (realStock === null || realStock === undefined) realStock = usQty;

    rowsOut.push({
      row: [
        pid, normText(stripShippingNote(anchor["Title"])), normText(anchor["Variation details"]) || null, usQty,
        siteCount, qtys.size > 1 ? "要確認" : "",
        numOrNullCsv(bySite.US["Item number"]), numOrNullCsv(bySite.US["Current price"]), intOrNullCsv(bySite.US["Sold quantity"]),
        bySite.UK ? numOrNullCsv(bySite.UK["Item number"]) : null, bySite.UK ? numOrNullCsv(bySite.UK["Current price"]) : null, bySite.UK ? intOrNullCsv(bySite.UK["Sold quantity"]) : null,
        bySite.AU ? numOrNullCsv(bySite.AU["Item number"]) : null, bySite.AU ? numOrNullCsv(bySite.AU["Current price"]) : null, bySite.AU ? intOrNullCsv(bySite.AU["Sold quantity"]) : null,
        purchasePrice, purchaseDate, purchaseFrom, note,
        bySite.UK ? intOrNullCsv(bySite.UK["Available quantity"]) : null, bySite.AU ? intOrNullCsv(bySite.AU["Available quantity"]) : null,
        realStock, realStockConfirmedAt, stocktakeQty, stocktakeAt,
        imageUrl, jaName, anchor["Custom label (SKU)"] || null, stocktakeChecked,
      ],
      status: "current",
      flag: siteCount < 3 || qtys.size > 1,
    });
  }

  let removedNewCount = 0;
  let removedStillCount = 0;
  for (const [key, s] of saved) {
    if (newKeySet.has(key)) continue;
    let note = s.fullRow[18] || "";
    if (s.wasRemoved) {
      removedStillCount++;
    } else {
      note = (note ? note + " / " : "") + `${today}時点の${removedNoteLabel}に見当たりません(削除/売り切れの可能性)`;
      removedNewCount++;
    }
    const carried = s.fullRow.slice(0, 15);
    rowsOut.push({
      row: carried.concat([s.fullRow[15], s.fullRow[16], s.fullRow[17], note]).concat(s.fullRow.slice(19, 29)),
      status: "removed",
    });
  }

  rowsOut.sort((a, b) => {
    if ((a.status === "removed") !== (b.status === "removed")) return a.status === "removed" ? 1 : -1;
    return String(a.row[0]).localeCompare(String(b.row[0]));
  });

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("在庫管理表");
  ws.addRow(INV_HEADERS);
  for (let c = 1; c <= INV_HEADERS.length; c++) {
    const cell = ws.getRow(1).getCell(c);
    cell.fill = INV_HEADER_FILL;
    cell.font = INV_HEADER_FONT;
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  }
  ws.getRow(1).height = 30;
  ws.views = [{ state: "frozen", ySplit: 1 }];

  rowsOut.forEach((r) => {
    const excelRow = ws.addRow(r.row);
    if (r.status === "removed" || r.flag) {
      const fill = r.status === "removed" ? INV_REMOVED_FILL : INV_FLAG_FILL;
      for (let c = 1; c <= INV_HEADERS.length; c++) excelRow.getCell(c).fill = fill;
    }
  });
  const widths = [10, 40, 30, 12, 10, 12, 14, 12, 12, 14, 12, 12, 14, 12, 12, 14, 12, 16, 30, 12, 12, 12, 16, 12, 18, 40, 30, 16, 12];
  widths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });

  const ws2 = wb.addWorksheet("要確認(重複出品)");
  ws2.addRow(["名寄せキー", "出品ID", "商品名", "バリエーション詳細", "SKU", "在庫数", "出品国", "現在価格", "通貨"]);
  for (let c = 1; c <= 9; c++) {
    ws2.getRow(1).getCell(c).fill = INV_HEADER_FILL;
    ws2.getRow(1).getCell(c).font = INV_HEADER_FONT;
  }
  ws2.views = [{ state: "frozen", ySplit: 1 }];
  anomalous.forEach((it) => {
    ws2.addRow([
      (it["Title"] || "") + " || " + (it["Variation details"] || ""),
      numOrNullCsv(it["Item number"]), it["Title"], it["Variation details"], it["Custom label (SKU)"],
      intOrNullCsv(it["Available quantity"]), it["Listing site"], numOrNullCsv(it["Current price"]), it["Currency"],
    ]);
  });
  [50, 14, 40, 30, 20, 10, 10, 12, 10].forEach((w, i) => { ws2.getColumn(i + 1).width = w; });

  const ws3 = wb.addWorksheet("この表について");
  ws3.getCell("A1").value = "在庫管理表について";
  ws3.getCell("A1").font = { bold: true, size: 13 };
  const notes = [
    sourceNote,
    "・商品IDは前回から引き継いでいます(売上管理表の商品IDとの対応を保つため)。新しく出品された商品には新しいIDを振っています。",
    "・黄色でハイライトした行は「3カ国揃っていない」または「在庫数が国ごとに違う」商品です。",
    "・赤色でハイライトした行は、今回の取得データに見当たらなかった商品です。削除されたか、売り切れて自動終了した可能性があります。",
    "・「要確認(重複出品)」シートには、同じ名前の出品が3件を超えて存在した商品を出品ID単位でそのまま残しています。",
    "・USに出品されていない商品(UK/AUのみ)は、転送漏れ・名称不一致とみなし商品リストには含めていません。",
    "・仕入価格・仕入日・仕入先・備考は、前回入力済みだった内容をそのまま引き継いでいます。",
    "・「リアル在庫」「リアル在庫確認日」「棚卸入力数量」「棚卸入力日時」は当社独自管理の値のため、このCSV再取込では一切上書きしません(前回の値をそのまま引き継ぎます)。",
    "・「UK在庫数」「AU在庫数」は、eBay自己申告のCSV上の在庫数をそのまま反映したものです(リアル在庫とは別物です)。",
    "・「画像URL」「日本語商品名」も当社独自管理の値のため、このCSV再取込では一切上書きしません(前回の値をそのまま引き継ぎます。画像URLはeBay同期を実行した場合のみ更新されます)。",
    "・「SKU」はCSV上のUS出品のCustom label(SKU)をそのまま反映したものです(検索用の補助情報で、画面には表示していません)。",
    "・「棚卸チェック」は棚卸で実際に数えたかどうかを管理する当社独自のフラグです(このCSV再取込では一切上書きしません。解除は棚卸画面の「棚卸チェックをすべて解除」ボタンでのみ行います)。",
  ];
  notes.forEach((n, i) => {
    const cell = ws3.getCell(`A${i + 3}`);
    cell.value = n;
    cell.font = { italic: true, color: { argb: "FF808080" }, size: 9 };
  });
  ws3.getColumn(1).width = 110;

  copyProtectedSheets(oldWb, wb);
  await atomicWriteWorkbook(wb, INVENTORY_PATH);

  return {
    total: rowsOut.length,
    new: newProducts.length,
    newProducts,
    removedNew: removedNewCount,
    removedStill: removedStillCount,
    anomalous: anomalous.length,
    noUsSkipped: noUsSkipped.length,
  };
}

module.exports = {
  rebuildInventoryFromRecords,
  // 以下はUS起点方式(ebay/inventorySync.js)からも再利用するための共有ヘルパー・定数
  normText,
  stripShippingNote,
  nameKey,
  INV_HEADER_FILL,
  INV_HEADER_FONT,
  INV_FLAG_FILL,
  INV_REMOVED_FILL,
};
