const {
  ORDER_LINES_SHEET,
  ORDER_LINE_HEADERS,
  ORDER_LINE_STATUS,
  MAPPING_STATUS,
  validateProtectedSheets,
  readItemMappings,
  addOrderLine,
} = require("./inventoryProtectedSheets");
const { SITE_ITEM_COLUMNS, isParserSafe } = require("./orderRegistration");

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

function orderRows(workbook, orderNo) {
  const sheet = workbook.getWorksheet(ORDER_LINES_SHEET);
  if (!sheet) return { sheet: null, headers: new Map(), rows: [] };
  const headers = headerMap(sheet);
  const rows = [];
  for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber++) {
    const row = sheet.getRow(rowNumber);
    if (String(row.getCell(headers.get("注文番号")).value || "") === String(orderNo)) rows.push({ row, rowNumber });
  }
  return { sheet, headers, rows };
}

function cellValue(entry, headers, name) {
  return entry.row.getCell(headers.get(name)).value;
}

function activeEntries(entries, headers) {
  return entries.filter((entry) => {
    const status = String(cellValue(entry, headers, "適用状態") || "");
    return status !== ORDER_LINE_STATUS.REVERSED && status !== ORDER_LINE_STATUS.CANCELLED;
  });
}

function desiredMatchesCurrent(entries, headers, parsed) {
  if (entries.length !== parsed.items.length) return false;
  return entries.every((entry, index) => {
    const item = parsed.items[index];
    return String(cellValue(entry, headers, "販売サイト") || "") === String(parsed.site)
      && String(cellValue(entry, headers, "eBay Item ID") || "") === String(item.ebayItemId)
      && String(cellValue(entry, headers, "商品タイトル") || "") === String(item.title)
      && String(cellValue(entry, headers, "SKU") || "") === String(item.sku || "")
      && Number(cellValue(entry, headers, "注文数量")) === Number(item.quantity);
  });
}

function buildInventoryIndex(workbook, INV_HEADERS, site) {
  const sheet = workbook.getWorksheet("在庫管理表") || workbook.worksheets[0];
  const pidColumn = INV_HEADERS.indexOf("商品ID") + 1;
  const itemColumn = INV_HEADERS.indexOf(SITE_ITEM_COLUMNS[site]) + 1;
  const stockColumn = INV_HEADERS.indexOf("リアル在庫") + 1;
  if (!sheet || !pidColumn || !itemColumn || !stockColumn) throw new Error("在庫管理表の必須列が不足しています");
  const byPid = new Map();
  const byItem = new Map();
  for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber++) {
    const row = sheet.getRow(rowNumber);
    const pid = String(row.getCell(pidColumn).value || "").trim();
    if (!pid) continue;
    if (!byPid.has(pid)) byPid.set(pid, []);
    byPid.get(pid).push(row);
    const itemId = String(row.getCell(itemColumn).value || "").trim();
    if (!itemId) continue;
    if (!byItem.has(itemId)) byItem.set(itemId, []);
    byItem.get(itemId).push(pid);
  }
  return { sheet, byPid, byItem, stockColumn };
}

function uniqueProductRow(inventory, pid) {
  const rows = inventory.byPid.get(pid) || [];
  if (!rows.length) throw httpError(409, `適用履歴の商品ID「${pid}」が現在の在庫管理表にありません`);
  if (rows.length > 1) throw httpError(409, `商品ID「${pid}」が在庫管理表に複数あります`);
  return rows[0];
}

function activeMappingIndex(workbook) {
  const index = new Map();
  for (const mapping of readItemMappings(workbook)) {
    if (String(mapping["状態"] || "") !== MAPPING_STATUS.ACTIVE) continue;
    index.set(`${String(mapping["販売サイト"] || "")}\u0000${String(mapping["eBay Item ID"] || "")}`, String(mapping["Agate商品ID"] || ""));
  }
  return index;
}

function prepareNewLines({ workbook, parsed, INV_HEADERS }) {
  if (!isParserSafe(parsed)) throw httpError(400, "編集後の注文明細を安全に解析できません。登録前レビューで内容を確認してください");
  const inventory = buildInventoryIndex(workbook, INV_HEADERS, parsed.site);
  const mappings = activeMappingIndex(workbook);
  const prepared = parsed.items.map((item) => {
    const itemId = String(item.ebayItemId);
    const mappingPid = mappings.get(`${parsed.site}\u0000${itemId}`) || "";
    const inventoryPids = inventory.byItem.get(itemId) || [];
    const inventoryPid = inventoryPids.length === 1 ? inventoryPids[0] : "";
    const inventoryConflict = inventoryPids.length > 1;
    let pid = "";
    let status = ORDER_LINE_STATUS.UNAPPLIED;
    let method = "";
    if (inventoryConflict || (mappingPid && inventoryPid && mappingPid !== inventoryPid)) {
      status = ORDER_LINE_STATUS.CONFLICT;
    } else if (mappingPid && (inventory.byPid.get(mappingPid) || []).length !== 1) {
      status = ORDER_LINE_STATUS.CONFLICT;
    } else {
      pid = mappingPid || inventoryPid;
      if (pid) {
        if ((inventory.byPid.get(pid) || []).length !== 1) throw httpError(409, `商品ID「${pid}」が在庫管理表に存在しないか重複しています`);
        status = ORDER_LINE_STATUS.APPLIED;
        method = mappingPid ? "確定ItemID対応" : "在庫管理表Item ID完全一致";
      }
    }
    return { item, pid, status, method };
  });
  return { inventory, prepared };
}

function applyStockDeltas(inventory, deltas) {
  const changes = [];
  for (const [pid, delta] of deltas) {
    const row = uniqueProductRow(inventory, pid);
    const before = Number(row.getCell(inventory.stockColumn).value) || 0;
    const after = before + delta;
    row.getCell(inventory.stockColumn).value = after;
    row.commit();
    changes.push({ pid, delta, before, after });
  }
  return changes;
}

function collectAppliedRestoreDeltas(entries, headers, deltas) {
  for (const entry of entries) {
    const status = String(cellValue(entry, headers, "適用状態") || "");
    if (status !== ORDER_LINE_STATUS.APPLIED) continue;
    const pid = String(cellValue(entry, headers, "適用商品ID") || "").trim();
    const quantity = Number(cellValue(entry, headers, "適用数量"));
    if (!/^P\d+$/.test(pid) || !Number.isInteger(quantity) || quantity < 1) {
      throw httpError(409, "過去の適用記録が不正なため在庫を復元できません");
    }
    deltas.set(pid, (deltas.get(pid) || 0) + quantity);
  }
}

function markAppliedEntriesReversed(entries, headers, now) {
  for (const entry of entries) {
    const status = String(cellValue(entry, headers, "適用状態") || "");
    if (status !== ORDER_LINE_STATUS.APPLIED) continue;
    entry.row.getCell(headers.get("適用状態")).value = ORDER_LINE_STATUS.REVERSED;
    entry.row.getCell(headers.get("更新日時")).value = now;
    entry.row.getCell(headers.get("解除日時")).value = now;
    entry.row.commit();
  }
}

function cancelNeverAppliedEntries(entries, headers, now) {
  for (const entry of entries) {
    const status = String(cellValue(entry, headers, "適用状態") || "");
    if (status === ORDER_LINE_STATUS.APPLIED || status === ORDER_LINE_STATUS.REVERSED || status === ORDER_LINE_STATUS.CANCELLED) continue;
    entry.row.getCell(headers.get("適用状態")).value = ORDER_LINE_STATUS.CANCELLED;
    entry.row.getCell(headers.get("更新日時")).value = now;
    entry.row.commit();
  }
}

function parsedFromSingleActiveLine(entry, headers, orderNo, quantity) {
  const item = {
    title: String(cellValue(entry, headers, "商品タイトル") || ""),
    ebayItemId: String(cellValue(entry, headers, "eBay Item ID") || ""),
    quantity,
    sku: String(cellValue(entry, headers, "SKU") || ""),
    parseStatus: "OK",
  };
  const site = String(cellValue(entry, headers, "販売サイト") || "");
  return { orderNo: String(orderNo), site, items: [item], itemCount: 1, quantityTotal: quantity, subtotalQuantity: quantity, parseStatus: "OK" };
}

function editManagedOrderInWorkbook({ workbook, orderNo, parsed = null, quantityOverride, INV_HEADERS, now = new Date().toISOString() }) {
  validateProtectedSheets(workbook, { allowMissing: true });
  const current = orderRows(workbook, orderNo);
  if (!current.rows.length) return { managed: false, writeNeeded: false };
  validateProtectedSheets(workbook);
  const active = activeEntries(current.rows, current.headers);
  if (!active.length) throw httpError(409, "この注文はすでに削除・解除済みです");
  if (!parsed && quantityOverride === undefined) {
    return { managed: true, writeNeeded: false, inventoryUnchanged: true, total: active.length };
  }
  if (!parsed) {
    const quantity = Number(quantityOverride);
    if (active.length !== 1) throw httpError(409, "複数明細注文の数量は注文詳細を再解析して編集してください");
    if (!Number.isInteger(quantity) || quantity < 1) throw httpError(400, "数量は1以上の整数で指定してください");
    parsed = parsedFromSingleActiveLine(active[0], current.headers, orderNo, quantity);
  }
  if (String(parsed.orderNo || "") !== String(orderNo)) throw httpError(400, "編集後明細の注文番号が一致しません");
  if (desiredMatchesCurrent(active, current.headers, parsed)) {
    return { managed: true, writeNeeded: false, retry: true, total: active.length };
  }

  const { inventory, prepared } = prepareNewLines({ workbook, parsed, INV_HEADERS });
  const deltas = new Map();
  collectAppliedRestoreDeltas(active, current.headers, deltas);
  for (const entry of prepared) {
    if (entry.status === ORDER_LINE_STATUS.APPLIED) deltas.set(entry.pid, (deltas.get(entry.pid) || 0) - Number(entry.item.quantity));
  }
  // 復元先・新適用先が全て一意に存在することを、行変更より前に確認する。
  for (const pid of deltas.keys()) uniqueProductRow(inventory, pid);
  markAppliedEntriesReversed(active, current.headers, now);
  const stockChanges = applyStockDeltas(inventory, deltas);
  cancelNeverAppliedEntries(active, current.headers, now);

  const maxSequence = current.rows.reduce((max, entry) => Math.max(max, Number(cellValue(entry, current.headers, "明細連番")) || 0), 0);
  prepared.forEach((entry, index) => {
    addOrderLine(workbook, {
      "注文番号": orderNo, "明細連番": maxSequence + index + 1, "販売サイト": parsed.site,
      "eBay Item ID": String(entry.item.ebayItemId), "商品タイトル": entry.item.title, "SKU": entry.item.sku || "",
      "注文数量": entry.item.quantity, "Agate商品ID": entry.pid, "適用商品ID": entry.pid,
      "適用数量": entry.status === ORDER_LINE_STATUS.APPLIED ? entry.item.quantity : "", "適用状態": entry.status,
      "紐付け方法": entry.method, "作成日時": now, "更新日時": now,
      "適用日時": entry.status === ORDER_LINE_STATUS.APPLIED ? now : "", "解除日時": "",
    });
  });
  validateProtectedSheets(workbook);
  return {
    managed: true, writeNeeded: true, retry: false, total: prepared.length,
    applied: prepared.filter((entry) => entry.status === ORDER_LINE_STATUS.APPLIED).length,
    unapplied: prepared.filter((entry) => entry.status !== ORDER_LINE_STATUS.APPLIED).length,
    conflict: prepared.filter((entry) => entry.status === ORDER_LINE_STATUS.CONFLICT).length,
    stockChanges,
  };
}

function deleteManagedOrdersInWorkbook({ workbook, orderNumbers, INV_HEADERS, now = new Date().toISOString() }) {
  validateProtectedSheets(workbook, { allowMissing: true });
  const targets = Array.from(orderNumbers, String);
  const found = targets.map((orderNo) => ({ orderNo, ...orderRows(workbook, orderNo) })).filter((entry) => entry.rows.length);
  if (!found.length) return { managedOrderNumbers: [], writeNeeded: false, restored: [], alreadyDeleted: [] };
  validateProtectedSheets(workbook);
  const inventory = buildInventoryIndex(workbook, INV_HEADERS, "US"); // 復元はItem ID列を使わず、PIDと在庫列だけを使う。
  const deltas = new Map();
  const alreadyDeleted = [];
  let changed = false;
  for (const order of found) {
    const active = activeEntries(order.rows, order.headers);
    if (!active.length) { alreadyDeleted.push(order.orderNo); continue; }
    changed = true;
    collectAppliedRestoreDeltas(active, order.headers, deltas);
  }
  for (const pid of deltas.keys()) uniqueProductRow(inventory, pid);
  for (const order of found) {
    const active = activeEntries(order.rows, order.headers);
    markAppliedEntriesReversed(active, order.headers, now);
    cancelNeverAppliedEntries(active, order.headers, now);
  }
  const restored = applyStockDeltas(inventory, deltas);
  validateProtectedSheets(workbook);
  return {
    managedOrderNumbers: found.map((entry) => entry.orderNo),
    writeNeeded: changed,
    restored,
    alreadyDeleted,
  };
}

async function editManagedOrderTransaction({ withInventoryLock, loadInventoryWorkbookLocked, atomicWriteWorkbook, INVENTORY_PATH, ...args }) {
  return withInventoryLock(async () => {
    const workbook = await loadInventoryWorkbookLocked();
    const result = editManagedOrderInWorkbook({ workbook, ...args });
    if (result.writeNeeded) await atomicWriteWorkbook(workbook, INVENTORY_PATH);
    return result;
  });
}

async function deleteManagedOrdersTransaction({ withInventoryLock, loadInventoryWorkbookLocked, atomicWriteWorkbook, INVENTORY_PATH, ...args }) {
  return withInventoryLock(async () => {
    const workbook = await loadInventoryWorkbookLocked();
    const result = deleteManagedOrdersInWorkbook({ workbook, ...args });
    if (result.writeNeeded) await atomicWriteWorkbook(workbook, INVENTORY_PATH);
    return result;
  });
}

module.exports = {
  editManagedOrderInWorkbook,
  deleteManagedOrdersInWorkbook,
  editManagedOrderTransaction,
  deleteManagedOrdersTransaction,
};
