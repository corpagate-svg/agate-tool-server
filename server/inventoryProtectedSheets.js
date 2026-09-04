const ORDER_LINES_SHEET = "注文明細";
const ITEM_MAPPINGS_SHEET = "確定ItemID対応";
const REPLENISHMENT_CANDIDATES_SHEET = "補充候補";
const PROTECTED_SHEET_NAMES = Object.freeze([ORDER_LINES_SHEET, ITEM_MAPPINGS_SHEET, REPLENISHMENT_CANDIDATES_SHEET]);

const ORDER_LINE_HEADERS = Object.freeze([
  "明細キー", "注文番号", "明細連番", "販売サイト", "eBay Item ID", "商品タイトル", "SKU", "注文数量",
  "Agate商品ID", "適用商品ID", "適用数量", "適用状態", "紐付け方法",
  "作成日時", "更新日時", "適用日時", "解除日時",
]);

const ITEM_MAPPING_HEADERS = Object.freeze([
  "対応キー", "販売サイト", "eBay Item ID", "対応版", "Agate商品ID", "状態",
  "確認日時", "確認者", "確認理由", "有効開始日時", "有効終了日時", "変更元対応キー",
]);

const REPLENISHMENT_CANDIDATE_HEADERS = Object.freeze([
  "補充候補ID", "商品ID", "US出品ID", "同期前US在庫", "同期後US在庫", "補充候補数量",
  "検知日時", "検知元", "状態", "処理日時", "処理者", "適用数量",
  "適用前リアル在庫", "適用後リアル在庫", "処理理由", "在庫履歴記録状態", "在庫履歴イベントID",
]);

const ORDER_LINE_STATUS = Object.freeze({
  UNAPPLIED: "未適用",
  APPLIED: "適用済み",
  REVIEW: "要確認",
  CONFLICT: "矛盾",
  REVERSED: "解除済み",
  CANCELLED: "取消済み",
});

const MAPPING_STATUS = Object.freeze({ ACTIVE: "有効", INACTIVE: "無効", CONFLICT: "矛盾" });
const REPLENISHMENT_STATUS = Object.freeze({
  PENDING: "未処理",
  APPROVED: "承認済み",
  REJECTED: "却下",
  INVALIDATED: "手動変更により失効",
  REVIEW: "商品不明／要確認",
});

const HEADER_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FF4F6B82" } };
const HEADER_FONT = { bold: true, color: { argb: "FFFFFFFF" } };

function normalizeKeyPart(value) {
  return encodeURIComponent(String(value === undefined || value === null ? "" : value).trim());
}

function createOrderLineKey({ orderNo, site, ebayItemId, lineNumber }) {
  const sequence = Number(lineNumber);
  if (!orderNo || !site || !ebayItemId || !Number.isInteger(sequence) || sequence < 1) {
    throw new Error("明細キーの生成に必要な値が不足しています");
  }
  return `OL:${normalizeKeyPart(orderNo)}:${normalizeKeyPart(site)}:${normalizeKeyPart(ebayItemId)}:${String(sequence).padStart(4, "0")}`;
}

function createMappingKey({ site, ebayItemId, version }) {
  const mappingVersion = Number(version);
  if (!site || !ebayItemId || !Number.isInteger(mappingVersion) || mappingVersion < 1) {
    throw new Error("対応キーの生成に必要な値が不足しています");
  }
  return `MAP:${normalizeKeyPart(site)}:${normalizeKeyPart(ebayItemId)}:v${String(mappingVersion).padStart(4, "0")}`;
}

function styleManagedSheet(ws, headers) {
  ws.addRow(headers);
  const header = ws.getRow(1);
  header.height = 28;
  for (let col = 1; col <= headers.length; col++) {
    header.getCell(col).fill = HEADER_FILL;
    header.getCell(col).font = HEADER_FONT;
    header.getCell(col).alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  }
  ws.views = [{ state: "frozen", ySplit: 1 }];
  headers.forEach((headerName, index) => {
    const widths = {
      "明細キー": 48, "対応キー": 42, "注文番号": 22, "明細連番": 10, "販売サイト": 12,
      "eBay Item ID": 18, "商品タイトル": 55, "SKU": 22, "注文数量": 12, "Agate商品ID": 16,
      "適用商品ID": 16, "適用数量": 12, "適用状態": 14, "紐付け方法": 16,
      "対応版": 10, "状態": 12, "確認者": 18, "確認理由": 40, "変更元対応キー": 42,
    };
    ws.getColumn(index + 1).width = widths[headerName] || 22;
  });
}

function ensureProtectedSheets(workbook) {
  let orderLines = workbook.getWorksheet(ORDER_LINES_SHEET);
  if (!orderLines) {
    orderLines = workbook.addWorksheet(ORDER_LINES_SHEET);
    styleManagedSheet(orderLines, ORDER_LINE_HEADERS);
  }
  let mappings = workbook.getWorksheet(ITEM_MAPPINGS_SHEET);
  if (!mappings) {
    mappings = workbook.addWorksheet(ITEM_MAPPINGS_SHEET);
    styleManagedSheet(mappings, ITEM_MAPPING_HEADERS);
  }
  return { orderLines, mappings };
}

// 補充候補は候補が発生した時だけ遅延作成する。注文処理など、無関係な管理シート
// 初期化では作成しない。
function ensureReplenishmentCandidatesSheet(workbook) {
  let candidates = workbook.getWorksheet(REPLENISHMENT_CANDIDATES_SHEET);
  if (!candidates) {
    candidates = workbook.addWorksheet(REPLENISHMENT_CANDIDATES_SHEET);
    styleManagedSheet(candidates, REPLENISHMENT_CANDIDATE_HEADERS);
  }
  return candidates;
}

function headerIndex(ws) {
  const map = new Map();
  ws.getRow(1).eachCell({ includeEmpty: true }, (cell, col) => map.set(String(cell.value || "").trim(), col));
  return map;
}

function validateHeaders(ws, expectedHeaders) {
  if (!ws) throw new Error("管理シートがありません");
  const actual = [];
  for (let col = 1; col <= expectedHeaders.length; col++) actual.push(String(ws.getRow(1).getCell(col).value || "").trim());
  if (actual.length !== expectedHeaders.length || actual.some((value, index) => value !== expectedHeaders[index])) {
    throw new Error(`${ws.name} のヘッダー構造が不正です`);
  }
}

function isBlankRow(row, columnCount) {
  for (let col = 1; col <= columnCount; col++) {
    const value = row.getCell(col).value;
    if (value !== null && value !== undefined && value !== "") return false;
  }
  return true;
}

function isUnset(value) {
  return value === null || value === undefined || value === "";
}

function validateOrderLinesSheet(ws) {
  validateHeaders(ws, ORDER_LINE_HEADERS);
  const index = headerIndex(ws);
  const keys = new Set();
  const allowedStatuses = new Set(Object.values(ORDER_LINE_STATUS));
  for (let rowNumber = 2; rowNumber <= ws.rowCount; rowNumber++) {
    const row = ws.getRow(rowNumber);
    if (isBlankRow(row, ORDER_LINE_HEADERS.length)) continue;
    const value = (name) => row.getCell(index.get(name)).value;
    const key = String(value("明細キー") || "").trim();
    const orderNo = String(value("注文番号") || "").trim();
    const site = String(value("販売サイト") || "").trim();
    const itemId = value("eBay Item ID");
    const title = String(value("商品タイトル") || "").trim();
    const quantity = Number(value("注文数量"));
    const sequence = Number(value("明細連番"));
    const status = String(value("適用状態") || "").trim();
    const createdAt = value("作成日時");
    const updatedAt = value("更新日時");
    if (!key || !orderNo || !site || typeof itemId !== "string" || !itemId.trim() || !title || !createdAt || !updatedAt) {
      throw new Error(`${ORDER_LINES_SHEET} ${rowNumber}行目の必須フィールドが不足しています`);
    }
    if (!Number.isInteger(quantity) || quantity < 1 || !Number.isInteger(sequence) || sequence < 1) {
      throw new Error(`${ORDER_LINES_SHEET} ${rowNumber}行目の数量または明細連番が不正です`);
    }
    const expectedKey = createOrderLineKey({ orderNo, site, ebayItemId: itemId, lineNumber: sequence });
    if (key !== expectedKey) throw new Error(`${ORDER_LINES_SHEET} ${rowNumber}行目の明細キーが行内容と一致しません`);
    if (!allowedStatuses.has(status)) throw new Error(`${ORDER_LINES_SHEET} ${rowNumber}行目の適用状態が不正です`);
    const appliedPidValue = value("適用商品ID");
    const appliedQuantityValue = value("適用数量");
    const appliedAt = value("適用日時");
    const reversedAt = value("解除日時");
    const appliedPid = String(appliedPidValue || "").trim();
    const appliedQuantity = Number(appliedQuantityValue);
    const unappliedStatus = status === ORDER_LINE_STATUS.UNAPPLIED
      || status === ORDER_LINE_STATUS.REVIEW
      || status === ORDER_LINE_STATUS.CONFLICT
      || status === ORDER_LINE_STATUS.CANCELLED;
    if (unappliedStatus && (!isUnset(appliedPidValue) || !isUnset(appliedQuantityValue) || !isUnset(appliedAt) || !isUnset(reversedAt))) {
      throw new Error(`${ORDER_LINES_SHEET} ${rowNumber}行目の未適用状態に適用情報があります`);
    }
    if (status === ORDER_LINE_STATUS.APPLIED || status === ORDER_LINE_STATUS.REVERSED) {
      if (!/^P\d+$/.test(appliedPid) || !Number.isInteger(appliedQuantity) || appliedQuantity < 1 || appliedQuantity > quantity || isUnset(appliedAt)) {
        throw new Error(`${ORDER_LINES_SHEET} ${rowNumber}行目の適用記録が不足しています`);
      }
    }
    if (status === ORDER_LINE_STATUS.APPLIED && !isUnset(reversedAt)) {
      throw new Error(`${ORDER_LINES_SHEET} ${rowNumber}行目の適用済みに解除日時があります`);
    }
    if (status === ORDER_LINE_STATUS.REVERSED && isUnset(reversedAt)) {
      throw new Error(`${ORDER_LINES_SHEET} ${rowNumber}行目の解除日時がありません`);
    }
    if (keys.has(key)) throw new Error(`${ORDER_LINES_SHEET} に重複した明細キーがあります: ${key}`);
    keys.add(key);
  }
  return { count: keys.size };
}

function validateItemMappingsSheet(ws) {
  validateHeaders(ws, ITEM_MAPPING_HEADERS);
  const index = headerIndex(ws);
  const keys = new Set();
  const activeMappings = new Set();
  const allowedStatuses = new Set(Object.values(MAPPING_STATUS));
  for (let rowNumber = 2; rowNumber <= ws.rowCount; rowNumber++) {
    const row = ws.getRow(rowNumber);
    if (isBlankRow(row, ITEM_MAPPING_HEADERS.length)) continue;
    const value = (name) => row.getCell(index.get(name)).value;
    const key = String(value("対応キー") || "").trim();
    const site = String(value("販売サイト") || "").trim();
    const itemId = value("eBay Item ID");
    const pid = String(value("Agate商品ID") || "").trim();
    const status = String(value("状態") || "").trim();
    const version = Number(value("対応版"));
    if (!key || !site || typeof itemId !== "string" || !itemId.trim() || !/^P\d+$/.test(pid)
      || !value("確認日時") || !String(value("確認者") || "").trim() || !String(value("確認理由") || "").trim() || !value("有効開始日時")) {
      throw new Error(`${ITEM_MAPPINGS_SHEET} ${rowNumber}行目の必須フィールドが不足または不正です`);
    }
    if (!Number.isInteger(version) || version < 1) throw new Error(`${ITEM_MAPPINGS_SHEET} ${rowNumber}行目の対応版が不正です`);
    if (!new Set(["US", "UK", "AU"]).has(site)) throw new Error(`${ITEM_MAPPINGS_SHEET} ${rowNumber}行目の販売サイトが不正です`);
    const expectedKey = createMappingKey({ site, ebayItemId: itemId, version });
    if (key !== expectedKey) throw new Error(`${ITEM_MAPPINGS_SHEET} ${rowNumber}行目の対応キーが行内容と一致しません`);
    if (!allowedStatuses.has(status)) throw new Error(`${ITEM_MAPPINGS_SHEET} ${rowNumber}行目の状態が不正です`);
    if (status === MAPPING_STATUS.INACTIVE && !value("有効終了日時")) {
      throw new Error(`${ITEM_MAPPINGS_SHEET} ${rowNumber}行目の有効終了日時がありません`);
    }
    if (keys.has(key)) throw new Error(`${ITEM_MAPPINGS_SHEET} に重複した対応キーがあります: ${key}`);
    keys.add(key);
    if (status === MAPPING_STATUS.ACTIVE) {
      const activeKey = `${site}\u0000${itemId}`;
      if (activeMappings.has(activeKey)) throw new Error(`${site} + ${itemId} に有効な確定対応が複数あります`);
      activeMappings.add(activeKey);
    }
  }
  return { count: keys.size, activeCount: activeMappings.size };
}

function validateReplenishmentCandidatesSheet(ws) {
  validateHeaders(ws, REPLENISHMENT_CANDIDATE_HEADERS);
  const index = headerIndex(ws);
  const ids = new Set();
  const allowedStatuses = new Set(Object.values(REPLENISHMENT_STATUS));
  const uuidPattern = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;
  for (let rowNumber = 2; rowNumber <= ws.rowCount; rowNumber++) {
    const row = ws.getRow(rowNumber);
    if (isBlankRow(row, REPLENISHMENT_CANDIDATE_HEADERS.length)) continue;
    const value = (name) => row.getCell(index.get(name)).value;
    const id = String(value("補充候補ID") || "").trim();
    const pid = String(value("商品ID") || "").trim();
    const itemId = value("US出品ID");
    const oldQty = Number(value("同期前US在庫"));
    const newQty = Number(value("同期後US在庫"));
    const candidateQty = Number(value("補充候補数量"));
    const status = String(value("状態") || "").trim();
    if (!uuidPattern.test(id) || !/^P\d+$/.test(pid) || typeof itemId !== "string" || !itemId.trim()
      || !value("検知日時") || !String(value("検知元") || "").trim()) {
      throw new Error(`${REPLENISHMENT_CANDIDATES_SHEET} ${rowNumber}行目の必須フィールドが不足または不正です`);
    }
    if (!Number.isSafeInteger(oldQty) || oldQty < 0 || !Number.isSafeInteger(newQty) || newQty <= oldQty
      || !Number.isSafeInteger(candidateQty) || candidateQty !== newQty - oldQty) {
      throw new Error(`${REPLENISHMENT_CANDIDATES_SHEET} ${rowNumber}行目の在庫数量が不正です`);
    }
    if (!allowedStatuses.has(status)) throw new Error(`${REPLENISHMENT_CANDIDATES_SHEET} ${rowNumber}行目の状態が不正です`);
    if (status === REPLENISHMENT_STATUS.PENDING) {
      for (const name of ["処理日時", "処理者", "適用数量", "適用前リアル在庫", "適用後リアル在庫", "処理理由", "在庫履歴記録状態", "在庫履歴イベントID"]) {
        if (!isUnset(value(name))) throw new Error(`${REPLENISHMENT_CANDIDATES_SHEET} ${rowNumber}行目の未処理候補に処理情報があります`);
      }
    } else if (status === REPLENISHMENT_STATUS.APPROVED) {
      const appliedQty = Number(value("適用数量"));
      const beforeReal = Number(value("適用前リアル在庫"));
      const afterReal = Number(value("適用後リアル在庫"));
      const historyState = String(value("在庫履歴記録状態") || "").trim();
      const historyEventId = String(value("在庫履歴イベントID") || "").trim();
      if (isUnset(value("処理日時")) || !String(value("処理者") || "").trim() || !String(value("処理理由") || "").trim()
        || !Number.isSafeInteger(appliedQty) || appliedQty !== candidateQty
        || !Number.isSafeInteger(beforeReal) || beforeReal < 0
        || !Number.isSafeInteger(afterReal) || afterReal !== beforeReal + appliedQty
        || !new Set(["記録待ち", "記録済み", "記録失敗"]).has(historyState) || !uuidPattern.test(historyEventId)) {
        throw new Error(`${REPLENISHMENT_CANDIDATES_SHEET} ${rowNumber}行目の承認記録が不正です`);
      }
    } else {
      if (isUnset(value("処理日時")) || !String(value("処理者") || "").trim() || !String(value("処理理由") || "").trim()) {
        throw new Error(`${REPLENISHMENT_CANDIDATES_SHEET} ${rowNumber}行目の処理記録が不足しています`);
      }
      for (const name of ["適用数量", "適用前リアル在庫", "適用後リアル在庫", "在庫履歴記録状態", "在庫履歴イベントID"]) {
        if (!isUnset(value(name))) throw new Error(`${REPLENISHMENT_CANDIDATES_SHEET} ${rowNumber}行目の非承認候補に適用情報があります`);
      }
    }
    if (ids.has(id)) throw new Error(`${REPLENISHMENT_CANDIDATES_SHEET} に重複した補充候補IDがあります: ${id}`);
    ids.add(id);
  }
  return { count: ids.size };
}

function validateProtectedSheets(workbook, { allowMissing = false } = {}) {
  const orderLines = workbook.getWorksheet(ORDER_LINES_SHEET);
  const mappings = workbook.getWorksheet(ITEM_MAPPINGS_SHEET);
  const candidates = workbook.getWorksheet(REPLENISHMENT_CANDIDATES_SHEET);
  if (allowMissing && !orderLines && !mappings) {
    if (candidates) validateReplenishmentCandidatesSheet(candidates);
    return { orderLines: 0, mappings: 0 };
  }
  if (!orderLines || !mappings) throw new Error("サーバー管理シートの一部が不足しています");
  const result = {
    orderLines: validateOrderLinesSheet(orderLines).count,
    mappings: validateItemMappingsSheet(mappings).count,
  };
  if (candidates) validateReplenishmentCandidatesSheet(candidates);
  return result;
}

function cloneCellValue(value) {
  if (value instanceof Date) return new Date(value.getTime());
  if (value && typeof value === "object") return { ...value };
  return value;
}

function copyWorksheet(source, destinationWorkbook, targetName = source.name) {
  const existing = destinationWorkbook.getWorksheet(targetName);
  if (existing) destinationWorkbook.removeWorksheet(existing.id);
  const target = destinationWorkbook.addWorksheet(targetName, { properties: { ...source.properties }, views: source.views.map((view) => ({ ...view })) });
  source.eachRow({ includeEmpty: true }, (sourceRow, rowNumber) => {
    const targetRow = target.getRow(rowNumber);
    targetRow.height = sourceRow.height;
    sourceRow.eachCell({ includeEmpty: true }, (sourceCell, colNumber) => {
      const targetCell = targetRow.getCell(colNumber);
      targetCell.value = cloneCellValue(sourceCell.value);
      if (sourceCell.style) targetCell.style = JSON.parse(JSON.stringify(sourceCell.style));
      if (sourceCell.numFmt) targetCell.numFmt = sourceCell.numFmt;
    });
    targetRow.commit();
  });
  source.columns.forEach((column, index) => {
    if (column.width) target.getColumn(index + 1).width = column.width;
    if (column.hidden) target.getColumn(index + 1).hidden = true;
  });
  target.state = source.state;
  return target;
}

function copyProtectedSheets(sourceWorkbook, destinationWorkbook) {
  const sourceHasAny = PROTECTED_SHEET_NAMES.some((name) => sourceWorkbook.getWorksheet(name));
  if (sourceHasAny) validateProtectedSheets(sourceWorkbook);
  for (const name of PROTECTED_SHEET_NAMES) {
    const source = sourceWorkbook.getWorksheet(name);
    if (source) copyWorksheet(source, destinationWorkbook, name);
  }
  ensureProtectedSheets(destinationWorkbook);
  validateProtectedSheets(destinationWorkbook);
  return destinationWorkbook;
}

function validateInventorySheet(workbook, requiredHeaders) {
  const ws = workbook.getWorksheet("在庫管理表");
  if (!ws) throw new Error("在庫管理表 シートが見つかりません");
  const headers = headerIndex(ws);
  for (const header of requiredHeaders) {
    if (!headers.has(header)) throw new Error(`在庫管理表に必須列「${header}」がありません`);
  }
  const pidColumn = headers.get("商品ID");
  const pids = new Set();
  for (let rowNumber = 2; rowNumber <= ws.rowCount; rowNumber++) {
    const pid = String(ws.getRow(rowNumber).getCell(pidColumn).value || "").trim();
    if (!pid) continue;
    if (pids.has(pid)) throw new Error(`在庫管理表に重複した商品IDがあります: ${pid}`);
    pids.add(pid);
  }
  if (!pids.size) throw new Error("在庫管理表の商品が0件のため取り込みを中止しました");
  return { productCount: pids.size };
}

function buildProtectedImportWorkbook({ uploadedWorkbook, serverWorkbook, requiredInventoryHeaders }) {
  const validation = validateInventorySheet(uploadedWorkbook, requiredInventoryHeaders);
  // アップロードWorkbook自体を土台にすることで、通常シートの画像・設定・書式を
  // 不要に作り直さない。アップロード側の管理シートだけを捨て、サーバー側で置換する。
  const merged = uploadedWorkbook;
  for (const name of PROTECTED_SHEET_NAMES) {
    const uploadedProtected = merged.getWorksheet(name);
    if (uploadedProtected) merged.removeWorksheet(uploadedProtected.id);
  }
  copyProtectedSheets(serverWorkbook, merged);
  validateInventorySheet(merged, requiredInventoryHeaders);
  validateProtectedSheets(merged);
  return { workbook: merged, ...validation };
}

function rowValues(headers, record) {
  return headers.map((header) => {
    const value = record[header];
    if (header === "eBay Item ID") return value === null || value === undefined ? "" : String(value);
    return value === undefined ? "" : value;
  });
}

function addOrderLine(workbook, record) {
  const { orderLines } = ensureProtectedSheets(workbook);
  const normalized = { ...record, "eBay Item ID": String(record["eBay Item ID"] || "") };
  if (!normalized["明細キー"]) {
    normalized["明細キー"] = createOrderLineKey({
      orderNo: normalized["注文番号"], site: normalized["販売サイト"], ebayItemId: normalized["eBay Item ID"], lineNumber: normalized["明細連番"],
    });
  }
  orderLines.addRow(rowValues(ORDER_LINE_HEADERS, normalized));
  validateOrderLinesSheet(orderLines);
  return normalized["明細キー"];
}

function addItemMapping(workbook, record) {
  const { mappings } = ensureProtectedSheets(workbook);
  const normalized = { ...record, "eBay Item ID": String(record["eBay Item ID"] || "") };
  if (!normalized["対応キー"]) {
    normalized["対応キー"] = createMappingKey({ site: normalized["販売サイト"], ebayItemId: normalized["eBay Item ID"], version: normalized["対応版"] });
  }
  mappings.addRow(rowValues(ITEM_MAPPING_HEADERS, normalized));
  validateItemMappingsSheet(mappings);
  return normalized["対応キー"];
}

function addReplenishmentCandidate(workbook, record) {
  const crypto = require("crypto");
  const candidates = ensureReplenishmentCandidatesSheet(workbook);
  const normalized = {
    ...record,
    "補充候補ID": record["補充候補ID"] || crypto.randomUUID(),
    "US出品ID": String(record["US出品ID"] || ""),
    "状態": record["状態"] || REPLENISHMENT_STATUS.PENDING,
  };
  candidates.addRow(rowValues(REPLENISHMENT_CANDIDATE_HEADERS, normalized));
  validateReplenishmentCandidatesSheet(candidates);
  return normalized["補充候補ID"];
}

function readManagedRows(workbook, sheetName, headers) {
  const ws = workbook.getWorksheet(sheetName);
  if (!ws) return [];
  validateHeaders(ws, headers);
  const rows = [];
  for (let rowNumber = 2; rowNumber <= ws.rowCount; rowNumber++) {
    const row = ws.getRow(rowNumber);
    if (isBlankRow(row, headers.length)) continue;
    const record = {};
    headers.forEach((header, index) => { record[header] = row.getCell(index + 1).value; });
    rows.push(record);
  }
  return rows;
}

module.exports = {
  ORDER_LINES_SHEET, ITEM_MAPPINGS_SHEET, REPLENISHMENT_CANDIDATES_SHEET, PROTECTED_SHEET_NAMES,
  ORDER_LINE_HEADERS, ITEM_MAPPING_HEADERS, REPLENISHMENT_CANDIDATE_HEADERS,
  ORDER_LINE_STATUS, MAPPING_STATUS, REPLENISHMENT_STATUS,
  createOrderLineKey, createMappingKey, ensureProtectedSheets, ensureReplenishmentCandidatesSheet,
  validateOrderLinesSheet, validateItemMappingsSheet, validateReplenishmentCandidatesSheet, validateProtectedSheets,
  copyWorksheet, copyProtectedSheets, validateInventorySheet, buildProtectedImportWorkbook,
  addOrderLine, addItemMapping, addReplenishmentCandidate,
  readOrderLines: (workbook) => readManagedRows(workbook, ORDER_LINES_SHEET, ORDER_LINE_HEADERS),
  readItemMappings: (workbook) => readManagedRows(workbook, ITEM_MAPPINGS_SHEET, ITEM_MAPPING_HEADERS),
  readReplenishmentCandidates: (workbook) => readManagedRows(workbook, REPLENISHMENT_CANDIDATES_SHEET, REPLENISHMENT_CANDIDATE_HEADERS),
};
