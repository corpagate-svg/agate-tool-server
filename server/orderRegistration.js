const {
  ORDER_LINE_STATUS,
  MAPPING_STATUS,
  createOrderLineKey,
  ensureProtectedSheets,
  validateProtectedSheets,
  addOrderLine,
  readOrderLines,
  readItemMappings,
} = require("./inventoryProtectedSheets");

const SITE_ITEM_COLUMNS = Object.freeze({ US: "US_出品ID", UK: "UK_出品ID", AU: "AU_出品ID" });

// 「小計（○点）」が無い現在のeBay表示形式では、小計自体は金額表記になりsubtotalQuantityが
// 取得できない。その場合は、Item ID解析とは独立な商品ブロック数(商品価格/商品合計ラベルの
// 出現数)の突合が取れているときだけ安全とみなす。parseOrderText側のparseStatus判定を
// そのまま信頼せず、ここでも同じ安全条件を独立に再検証する(既存の多重防御方針を踏襲)。
function isSubtotalSafe(parsed) {
  if (parsed.subtotalQuantity !== null) return parsed.quantityTotal === parsed.subtotalQuantity;
  return Boolean(parsed.itemCount > 0
    && parsed.productPriceLabelCount === parsed.itemCount
    && parsed.productTotalLabelCount === parsed.itemCount
    && parsed.productPriceLabelCount === parsed.productTotalLabelCount);
}

function isParserSafe(parsed) {
  return Boolean(parsed && parsed.parseStatus === "OK"
    && SITE_ITEM_COLUMNS[parsed.site]
    && parsed.itemCount > 0
    && isSubtotalSafe(parsed)
    && parsed.items.every((item) => item.parseStatus === "OK" && typeof item.ebayItemId === "string"
      && item.ebayItemId && Number.isInteger(item.quantity) && item.quantity > 0));
}

function inventoryRowsByPidAndItem(ws, INV_HEADERS, parsedSite) {
  const pidCol = INV_HEADERS.indexOf("商品ID") + 1;
  const itemCol = INV_HEADERS.indexOf(SITE_ITEM_COLUMNS[parsedSite]) + 1;
  const realCol = INV_HEADERS.indexOf("リアル在庫") + 1;
  if (!pidCol || !itemCol || !realCol) throw new Error("在庫管理表の必須列が不足しています");
  const byPid = new Map();
  const byItem = new Map();
  for (let rowNumber = 2; rowNumber <= ws.rowCount; rowNumber++) {
    const row = ws.getRow(rowNumber);
    const pid = String(row.getCell(pidCol).value || "").trim();
    if (!pid) continue;
    byPid.set(pid, row);
    const itemId = String(row.getCell(itemCol).value || "").trim();
    if (!itemId) continue;
    if (!byItem.has(itemId)) byItem.set(itemId, []);
    byItem.get(itemId).push(pid);
  }
  return { byPid, byItem, realCol };
}

function activeMappingIndex(mappings) {
  const index = new Map();
  for (const mapping of mappings) {
    if (String(mapping["状態"] || "").trim() !== MAPPING_STATUS.ACTIVE) continue;
    const key = `${String(mapping["販売サイト"] || "").trim()}\u0000${String(mapping["eBay Item ID"] || "").trim()}`;
    index.set(key, String(mapping["Agate商品ID"] || "").trim());
  }
  return index;
}

function parsedLineIdentity(parsed, item, index) {
  return {
    key: createOrderLineKey({ orderNo: parsed.orderNo, site: parsed.site, ebayItemId: item.ebayItemId, lineNumber: index + 1 }),
    orderNo: String(parsed.orderNo),
    lineNumber: index + 1,
    site: String(parsed.site),
    itemId: String(item.ebayItemId),
    quantity: Number(item.quantity),
  };
}

function existingLinesMatch(parsed, existing) {
  if (existing.length !== parsed.items.length) return false;
  const byKey = new Map(existing.map((line) => [String(line["明細キー"] || ""), line]));
  return parsed.items.every((item, index) => {
    const expected = parsedLineIdentity(parsed, item, index);
    const line = byKey.get(expected.key);
    return line
      && String(line["注文番号"] || "") === expected.orderNo
      && Number(line["明細連番"]) === expected.lineNumber
      && String(line["販売サイト"] || "") === expected.site
      && String(line["eBay Item ID"] || "") === expected.itemId
      && Number(line["注文数量"]) === expected.quantity;
  });
}

function conflictError(message) {
  const error = new Error(message);
  error.statusCode = 409;
  return error;
}

function processOrderInventoryWorkbook({ workbook, parsed, salesOrderExists, INV_HEADERS, now = new Date().toISOString() }) {
  validateProtectedSheets(workbook, { allowMissing: true });
  const hadProtectedSheets = Boolean(workbook.getWorksheet("注文明細"));
  if (!hadProtectedSheets) ensureProtectedSheets(workbook);
  validateProtectedSheets(workbook);

  const allLines = readOrderLines(workbook);
  const existing = allLines.filter((line) => String(line["注文番号"] || "") === String(parsed.orderNo || ""));
  if (existing.length) {
    const hasDeletedHistory = existing.some((line) => line["適用状態"] === ORDER_LINE_STATUS.REVERSED
      || line["適用状態"] === ORDER_LINE_STATUS.CANCELLED);
    if (hasDeletedHistory) {
      throw conflictError("この注文番号には削除済みの注文履歴があります。通常の再取り込みはできません。");
    }
    if (!existingLinesMatch(parsed, existing)) throw conflictError("同じ注文番号の保存済み注文明細と今回の内容が一致しません");
    return {
      writeNeeded: false, retry: !salesOrderExists, alreadyRegistered: salesOrderExists,
      total: existing.length,
      applied: existing.filter((line) => line["適用状態"] === ORDER_LINE_STATUS.APPLIED).length,
      unapplied: existing.filter((line) => line["適用状態"] !== ORDER_LINE_STATUS.APPLIED).length,
      conflict: existing.filter((line) => line["適用状態"] === ORDER_LINE_STATUS.CONFLICT).length,
    };
  }
  if (salesOrderExists) {
    return { writeNeeded: false, legacyOrder: true, total: 0, applied: 0, unapplied: 0, conflict: 0 };
  }

  const safe = isParserSafe(parsed);
  if (!parsed.items.length) {
    return { writeNeeded: hadProtectedSheets ? false : true, parserSafe: false, total: 0, applied: 0, unapplied: 0, conflict: 0 };
  }
  const canPersistAllItems = parsed.items.every((item) => typeof item.ebayItemId === "string" && item.ebayItemId
    && item.title && Number.isInteger(item.quantity) && item.quantity > 0);
  // 必須値を失った明細を一部だけ保存すると、後のretryで「注文全体の同一性」を保証できない。
  // 解析結果はUIに表示できるが、管理シートへの部分保存・在庫反映はどちらも行わない。
  if (!canPersistAllItems) {
    return { writeNeeded: hadProtectedSheets ? false : true, parserSafe: false, total: parsed.items.length, applied: 0, unapplied: parsed.items.length, conflict: 0 };
  }

  const ws = workbook.getWorksheet("在庫管理表") || workbook.worksheets[0];
  const inventory = safe ? inventoryRowsByPidAndItem(ws, INV_HEADERS, parsed.site) : null;
  const mappings = safe ? activeMappingIndex(readItemMappings(workbook)) : new Map();
  const prepared = [];
  const deltas = new Map();

  parsed.items.forEach((item, index) => {
    let pid = "";
    let status = safe ? ORDER_LINE_STATUS.UNAPPLIED : ORDER_LINE_STATUS.REVIEW;
    let method = "";
    if (safe) {
      const mappingPid = mappings.get(`${parsed.site}\u0000${item.ebayItemId}`) || "";
      const inventoryPids = inventory.byItem.get(String(item.ebayItemId)) || [];
      const inventoryPid = inventoryPids.length === 1 ? inventoryPids[0] : "";
      const inventoryConflict = inventoryPids.length > 1;
      if (inventoryConflict || (mappingPid && inventoryPid && mappingPid !== inventoryPid)
        || (mappingPid && !inventory.byPid.has(mappingPid))) {
        status = ORDER_LINE_STATUS.CONFLICT;
      } else {
        pid = mappingPid || inventoryPid;
        if (pid) {
          status = ORDER_LINE_STATUS.APPLIED;
          method = mappingPid ? "確定ItemID対応" : "在庫管理表Item ID完全一致";
          deltas.set(pid, (deltas.get(pid) || 0) + item.quantity);
        }
      }
    }
    prepared.push({ item, index, pid, status, method });
  });

  for (const [pid, quantity] of deltas) {
    const row = inventory.byPid.get(pid);
    if (!row) throw new Error(`適用対象商品ID「${pid}」が在庫管理表にありません`);
    const before = Number(row.getCell(inventory.realCol).value) || 0;
    row.getCell(inventory.realCol).value = before - quantity; // 既存仕様どおりマイナス在庫を許可し、0へ丸めない。
    row.commit();
  }

  for (const entry of prepared) {
    const { item, index, pid, status, method } = entry;
    if (status === ORDER_LINE_STATUS.APPLIED && !pid) throw new Error("適用商品IDと在庫減算対象が一致しません");
    addOrderLine(workbook, {
      "注文番号": parsed.orderNo, "明細連番": index + 1, "販売サイト": parsed.site,
      "eBay Item ID": String(item.ebayItemId), "商品タイトル": item.title, "SKU": item.sku || "",
      "注文数量": item.quantity, "Agate商品ID": pid, "適用商品ID": pid,
      "適用数量": status === ORDER_LINE_STATUS.APPLIED ? item.quantity : "", "適用状態": status,
      "紐付け方法": method, "作成日時": now, "更新日時": now,
      "適用日時": status === ORDER_LINE_STATUS.APPLIED ? now : "", "解除日時": "",
    });
  }
  validateProtectedSheets(workbook);
  return {
    writeNeeded: true, parserSafe: safe, total: prepared.length,
    applied: prepared.filter((entry) => entry.status === ORDER_LINE_STATUS.APPLIED).length,
    unapplied: prepared.filter((entry) => entry.status !== ORDER_LINE_STATUS.APPLIED).length,
    conflict: prepared.filter((entry) => entry.status === ORDER_LINE_STATUS.CONFLICT).length,
  };
}

async function processOrderInventoryTransaction({
  withInventoryLock, loadInventoryWorkbookLocked, atomicWriteWorkbook, INVENTORY_PATH,
  parsed, salesOrderExists, INV_HEADERS, now,
}) {
  return withInventoryLock(async () => {
    const workbook = await loadInventoryWorkbookLocked();
    const result = processOrderInventoryWorkbook({ workbook, parsed, salesOrderExists, INV_HEADERS, now });
    if (!result.legacyOrder && result.writeNeeded) await atomicWriteWorkbook(workbook, INVENTORY_PATH);
    return result;
  });
}

module.exports = {
  SITE_ITEM_COLUMNS, isParserSafe, existingLinesMatch,
  processOrderInventoryWorkbook, processOrderInventoryTransaction,
};
