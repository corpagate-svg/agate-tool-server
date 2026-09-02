const {
  ORDER_LINES_SHEET,
  ORDER_LINE_HEADERS,
  ORDER_LINE_STATUS,
  MAPPING_STATUS,
  validateProtectedSheets,
  readItemMappings,
  addItemMapping,
} = require("./inventoryProtectedSheets");

const SITE_ITEM_COLUMNS = Object.freeze({ US: "US_出品ID", UK: "UK_出品ID", AU: "AU_出品ID" });

function httpError(statusCode, message) {
  const error = new Error(message);
  error.statusCode = statusCode;
  return error;
}

function headerMap(sheet) {
  const map = new Map();
  sheet.getRow(1).eachCell({ includeEmpty: true }, (cell, column) => map.set(String(cell.value || "").trim(), column));
  return map;
}

function isBlank(value) {
  return value === null || value === undefined || value === "";
}

function findUniqueProduct(sheet, INV_HEADERS, pid) {
  const pidColumn = INV_HEADERS.indexOf("商品ID") + 1;
  const matches = [];
  for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber++) {
    const row = sheet.getRow(rowNumber);
    if (String(row.getCell(pidColumn).value || "").trim() === pid) matches.push(row);
  }
  if (!matches.length) throw httpError(404, `商品ID「${pid}」が在庫管理表にありません`);
  if (matches.length > 1) throw httpError(409, `商品ID「${pid}」が在庫管理表に複数あります`);
  return matches[0];
}

function itemIdProductIds(sheet, INV_HEADERS, site, itemId) {
  const pidColumn = INV_HEADERS.indexOf("商品ID") + 1;
  const itemColumn = INV_HEADERS.indexOf(SITE_ITEM_COLUMNS[site]) + 1;
  const matches = [];
  for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber++) {
    const row = sheet.getRow(rowNumber);
    if (String(row.getCell(itemColumn).value || "").trim() !== itemId) continue;
    const pid = String(row.getCell(pidColumn).value || "").trim();
    if (pid) matches.push(pid);
  }
  return matches;
}

function resolveOrderLineInWorkbook({ workbook, lineKey, pid, INV_HEADERS, confirmedBy = "Agate Trade Hub手動確認", now = new Date().toISOString() }) {
  validateProtectedSheets(workbook);
  const selectedPid = String(pid || "").trim();
  if (!/^P\d+$/.test(selectedPid)) throw httpError(400, "有効なPxxxxを指定してください");

  const orderSheet = workbook.getWorksheet(ORDER_LINES_SHEET);
  const orderHeaders = headerMap(orderSheet);
  let targetRow = null;
  for (let rowNumber = 2; rowNumber <= orderSheet.rowCount; rowNumber++) {
    const row = orderSheet.getRow(rowNumber);
    if (String(row.getCell(orderHeaders.get("明細キー")).value || "") === String(lineKey || "")) {
      targetRow = row;
      break;
    }
  }
  if (!targetRow) throw httpError(404, "対象の注文明細が見つかりません");
  const value = (name) => targetRow.getCell(orderHeaders.get(name)).value;
  const status = String(value("適用状態") || "").trim();
  const orderQuantity = Number(value("注文数量"));
  const site = String(value("販売サイト") || "").trim();
  const itemIdValue = value("eBay Item ID");
  const itemId = typeof itemIdValue === "string" ? itemIdValue.trim() : "";
  const appliedPid = String(value("適用商品ID") || "").trim();
  const appliedQuantity = Number(value("適用数量"));

  if (status === ORDER_LINE_STATUS.APPLIED) {
    if (appliedPid === selectedPid && appliedQuantity === orderQuantity) {
      return { writeNeeded: false, alreadyApplied: true, lineKey, pid: selectedPid, quantity: orderQuantity };
    }
    throw httpError(409, "この明細は別の商品または数量ですでに適用済みです");
  }
  if (status !== ORDER_LINE_STATUS.UNAPPLIED) {
    throw httpError(409, `適用状態「${status}」の明細は通常の未紐付け解決では適用できません`);
  }
  if (!isBlank(value("適用商品ID")) || !isBlank(value("適用数量")) || !isBlank(value("適用日時")) || !isBlank(value("解除日時"))) {
    throw httpError(409, "未適用明細に既存の適用情報があります");
  }
  if (!isBlank(value("Agate商品ID"))) throw httpError(409, "未適用明細に既存のAgate商品IDがあります");
  if (!Number.isInteger(orderQuantity) || orderQuantity < 1) throw httpError(409, "注文明細の注文数量が不正です");
  if (!SITE_ITEM_COLUMNS[site]) throw httpError(409, "注文明細の販売サイトが自動適用対象ではありません");
  if (!itemId) throw httpError(409, "注文明細のeBay Item IDが不正です");

  const inventorySheet = workbook.getWorksheet("在庫管理表") || workbook.worksheets[0];
  const productRow = findUniqueProduct(inventorySheet, INV_HEADERS, selectedPid);
  const itemMatches = itemIdProductIds(inventorySheet, INV_HEADERS, site, itemId);
  if (itemMatches.some((matchedPid) => matchedPid !== selectedPid)) {
    throw httpError(409, `在庫管理表では ${site} + ${itemId} が別の商品IDに紐付いています`);
  }

  const mappings = readItemMappings(workbook).filter((mapping) => String(mapping["販売サイト"] || "").trim() === site
    && String(mapping["eBay Item ID"] || "").trim() === itemId);
  const activeMappings = mappings.filter((mapping) => mapping["状態"] === MAPPING_STATUS.ACTIVE);
  if (activeMappings.length > 1) throw httpError(409, "同じItem IDの有効な確定対応が複数あります");
  if (activeMappings.length === 1 && String(activeMappings[0]["Agate商品ID"] || "").trim() !== selectedPid) {
    throw httpError(409, "このItem IDには別Pxxxxの有効な確定対応があります");
  }
  if (!activeMappings.length && mappings.length) {
    throw httpError(409, "このItem IDには過去の対応履歴があります。version変更処理で確認してください");
  }

  const realStockColumn = INV_HEADERS.indexOf("リアル在庫") + 1;
  const productNameColumn = INV_HEADERS.indexOf("商品名") + 1;
  const before = Number(productRow.getCell(realStockColumn).value) || 0;
  const after = before - orderQuantity;
  productRow.getCell(realStockColumn).value = after;
  productRow.commit();

  targetRow.getCell(orderHeaders.get("Agate商品ID")).value = selectedPid;
  targetRow.getCell(orderHeaders.get("適用商品ID")).value = selectedPid;
  targetRow.getCell(orderHeaders.get("適用数量")).value = orderQuantity;
  targetRow.getCell(orderHeaders.get("適用状態")).value = ORDER_LINE_STATUS.APPLIED;
  targetRow.getCell(orderHeaders.get("紐付け方法")).value = "人間によるItem ID確認";
  targetRow.getCell(orderHeaders.get("更新日時")).value = now;
  targetRow.getCell(orderHeaders.get("適用日時")).value = now;
  targetRow.commit();

  let mappingAdded = false;
  if (!activeMappings.length) {
    addItemMapping(workbook, {
      "販売サイト": site, "eBay Item ID": itemId, "対応版": 1, "Agate商品ID": selectedPid,
      "状態": MAPPING_STATUS.ACTIVE, "確認日時": now, "確認者": confirmedBy,
      "確認理由": `未紐付け注文明細 ${lineKey} を人間が確認`, "有効開始日時": now,
      "有効終了日時": "", "変更元対応キー": "",
    });
    mappingAdded = true;
  }
  validateProtectedSheets(workbook);
  return {
    writeNeeded: true, alreadyApplied: false, lineKey, orderNo: String(value("注文番号") || ""),
    site, ebayItemId: itemId, quantity: orderQuantity, pid: selectedPid,
    productName: String(productRow.getCell(productNameColumn).value || ""), before, after, mappingAdded,
  };
}

async function resolveOrderLineTransaction({
  withInventoryLock, loadInventoryWorkbookLocked, atomicWriteWorkbook, INVENTORY_PATH,
  lineKey, pid, INV_HEADERS, confirmedBy, now,
}) {
  return withInventoryLock(async () => {
    const workbook = await loadInventoryWorkbookLocked();
    const result = resolveOrderLineInWorkbook({ workbook, lineKey, pid, INV_HEADERS, confirmedBy, now });
    if (result.writeNeeded) await atomicWriteWorkbook(workbook, INVENTORY_PATH);
    return result;
  });
}

module.exports = { SITE_ITEM_COLUMNS, resolveOrderLineInWorkbook, resolveOrderLineTransaction };
