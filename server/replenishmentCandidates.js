const crypto = require("crypto");
const { withInventoryLock, atomicWriteWorkbook } = require("./inventoryLock");
const { appendHistoryOnce } = require("./inventoryHistory");
const {
  REPLENISHMENT_CANDIDATES_SHEET,
  REPLENISHMENT_CANDIDATE_HEADERS,
  REPLENISHMENT_STATUS,
  readReplenishmentCandidates,
  validateProtectedSheets,
  validateReplenishmentCandidatesSheet,
} = require("./inventoryProtectedSheets");

const UUID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;
const HISTORY_PENDING = "記録待ち";
const HISTORY_RECORDED = "記録済み";
const HISTORY_FAILED = "記録失敗";

function columnMap(ws) {
  const result = new Map();
  ws.getRow(1).eachCell({ includeEmpty: true }, (cell, column) => result.set(String(cell.value || "").trim(), column));
  return result;
}

function findCandidateRow(ws, candidateId) {
  const columns = columnMap(ws);
  const idColumn = columns.get("補充候補ID");
  for (let rowNumber = 2; rowNumber <= ws.rowCount; rowNumber++) {
    const row = ws.getRow(rowNumber);
    if (String(row.getCell(idColumn).value || "").trim() === candidateId) return { row, columns };
  }
  return null;
}

function cellValue(found, name) {
  return found.row.getCell(found.columns.get(name)).value;
}

function setCell(found, name, value) {
  found.row.getCell(found.columns.get(name)).value = value;
}

function error(statusCode, message) {
  const value = new Error(message);
  value.statusCode = statusCode;
  return value;
}

function assertCandidateId(candidateId) {
  if (typeof candidateId !== "string" || !UUID_PATTERN.test(candidateId)) throw error(400, "補充候補IDが不正です");
}

function inventoryRowsByPid(workbook, INV_COL) {
  const ws = workbook.getWorksheet("在庫管理表") || workbook.worksheets[0];
  const rows = new Map();
  for (let rowNumber = 2; rowNumber <= ws.rowCount; rowNumber++) {
    const row = ws.getRow(rowNumber);
    const pid = String(row.getCell(INV_COL.商品ID).value || "").trim();
    if (pid) rows.set(pid, row);
  }
  return rows;
}

function requiredInventoryColumn(INV_HEADERS, headerName) {
  if (!Array.isArray(INV_HEADERS)) throw error(500, "在庫管理表の列定義が不正です");
  const index = INV_HEADERS.indexOf(headerName);
  if (index < 0) throw error(500, `在庫管理表に必須列「${headerName}」がありません`);
  return index + 1;
}

async function listPendingCandidates({ loadInventoryWorkbook, INV_HEADERS, INV_COL }) {
  const workbook = await loadInventoryWorkbook();
  validateProtectedSheets(workbook, { allowMissing: true });
  const candidateSheet = workbook.getWorksheet(REPLENISHMENT_CANDIDATES_SHEET);
  if (!candidateSheet) return [];
  const currentUsStockColumn = requiredInventoryColumn(INV_HEADERS, "在庫数(現物)");
  const products = inventoryRowsByPid(workbook, INV_COL);
  return readReplenishmentCandidates(workbook)
    .filter((candidate) => candidate["状態"] === REPLENISHMENT_STATUS.PENDING)
    .map((candidate) => {
      const product = products.get(String(candidate["商品ID"] || "").trim());
      return {
        candidateId: candidate["補充候補ID"], productId: candidate["商品ID"],
        productName: product ? product.getCell(INV_COL.商品名).value : "",
        usItemId: String(candidate["US出品ID"] || ""), beforeUsStock: candidate["同期前US在庫"],
        afterUsStock: candidate["同期後US在庫"], candidateQuantity: candidate["補充候補数量"],
        currentUsStock: product ? product.getCell(currentUsStockColumn).value : null,
        currentRealStock: product ? product.getCell(INV_COL.リアル在庫).value : null,
        detectedAt: candidate["検知日時"], status: candidate["状態"],
      };
    });
}

function markNonApplied(found, status, { at, actor, reason }) {
  setCell(found, "状態", status);
  setCell(found, "処理日時", at);
  setCell(found, "処理者", actor);
  setCell(found, "適用数量", "");
  setCell(found, "適用前リアル在庫", "");
  setCell(found, "適用後リアル在庫", "");
  setCell(found, "処理理由", reason);
  setCell(found, "在庫履歴記録状態", "");
  setCell(found, "在庫履歴イベントID", "");
  found.row.commit();
}

function invalidatePendingCandidates(workbook, productId, { at = new Date().toISOString(), actor = "画面操作", reason = "リアル在庫手動変更により失効" } = {}) {
  const ws = workbook.getWorksheet(REPLENISHMENT_CANDIDATES_SHEET);
  if (!ws) return 0;
  validateReplenishmentCandidatesSheet(ws);
  const columns = columnMap(ws);
  let invalidated = 0;
  for (let rowNumber = 2; rowNumber <= ws.rowCount; rowNumber++) {
    const row = ws.getRow(rowNumber);
    const found = { row, columns };
    if (String(cellValue(found, "商品ID") || "").trim() !== String(productId)) continue;
    if (String(cellValue(found, "状態") || "").trim() !== REPLENISHMENT_STATUS.PENDING) continue;
    markNonApplied(found, REPLENISHMENT_STATUS.INVALIDATED, { at, actor, reason });
    invalidated++;
  }
  validateReplenishmentCandidatesSheet(ws);
  return invalidated;
}

async function updateHistoryState({ loadInventoryWorkbookLocked, INVENTORY_PATH, candidateId, eventId, state }) {
  return withInventoryLock(async () => {
    const workbook = await loadInventoryWorkbookLocked();
    const ws = workbook.getWorksheet(REPLENISHMENT_CANDIDATES_SHEET);
    if (!ws) return;
    const found = findCandidateRow(ws, candidateId);
    if (!found || cellValue(found, "状態") !== REPLENISHMENT_STATUS.APPROVED || cellValue(found, "在庫履歴イベントID") !== eventId) return;
    setCell(found, "在庫履歴記録状態", state);
    found.row.commit();
    validateProtectedSheets(workbook, { allowMissing: true });
    await atomicWriteWorkbook(workbook, INVENTORY_PATH);
  });
}

async function updateBulkHistoryState({ loadInventoryWorkbookLocked, INVENTORY_PATH, eventId, state }) {
  return withInventoryLock(async () => {
    const workbook = await loadInventoryWorkbookLocked();
    const ws = workbook.getWorksheet(REPLENISHMENT_CANDIDATES_SHEET);
    if (!ws) return;
    const columns = columnMap(ws);
    let changed = false;
    for (let rowNumber = 2; rowNumber <= ws.rowCount; rowNumber++) {
      const found = { row: ws.getRow(rowNumber), columns };
      if (cellValue(found, "状態") !== REPLENISHMENT_STATUS.APPROVED
        || cellValue(found, "処理理由") !== "補充候補一括承認"
        || cellValue(found, "在庫履歴イベントID") !== eventId) continue;
      setCell(found, "在庫履歴記録状態", state);
      found.row.commit();
      changed = true;
    }
    if (!changed) return;
    validateProtectedSheets(workbook, { allowMissing: true });
    await atomicWriteWorkbook(workbook, INVENTORY_PATH);
  });
}

async function recordApprovalHistory(deps, approval) {
  try {
    await (deps.appendHistoryOnce || appendHistoryOnce)(deps.HISTORY_PATH, {
      pid: approval.productId, name: approval.productName, before: approval.before,
      after: approval.after, reason: "補充候補承認", at: approval.processedAt,
    }, approval.eventId);
    await updateHistoryState({ ...deps, candidateId: approval.candidateId, eventId: approval.eventId, state: HISTORY_RECORDED });
    return { ...approval, historyStatus: HISTORY_RECORDED };
  } catch (historyError) {
    try {
      await updateHistoryState({ ...deps, candidateId: approval.candidateId, eventId: approval.eventId, state: HISTORY_FAILED });
    } catch (_) { /* source of truthの承認済み状態は既に保存済みなので維持する */ }
    return { ...approval, historyStatus: HISTORY_FAILED, historyWarning: "在庫変更履歴の記録に失敗しました" };
  }
}

async function approveCandidate(deps) {
  assertCandidateId(deps.candidateId);
  const usItemIdColumn = requiredInventoryColumn(deps.INV_HEADERS, "US_出品ID");
  const outcome = await withInventoryLock(async () => {
    const workbook = await deps.loadInventoryWorkbookLocked();
    validateProtectedSheets(workbook, { allowMissing: true });
    const candidateSheet = workbook.getWorksheet(REPLENISHMENT_CANDIDATES_SHEET);
    if (!candidateSheet) throw error(404, "補充候補が見つかりません");
    const found = findCandidateRow(candidateSheet, deps.candidateId);
    if (!found) throw error(404, "補充候補が見つかりません");
    const status = String(cellValue(found, "状態") || "").trim();
    if (status === REPLENISHMENT_STATUS.APPROVED) {
      const productId = String(cellValue(found, "商品ID") || "");
      const product = inventoryRowsByPid(workbook, deps.INV_COL).get(productId);
      return {
        alreadyProcessed: true, candidateId: deps.candidateId,
        eventId: String(cellValue(found, "在庫履歴イベントID") || ""),
        historyStatus: String(cellValue(found, "在庫履歴記録状態") || ""),
        processingReason: String(cellValue(found, "処理理由") || ""),
        productId, productName: product ? product.getCell(deps.INV_COL.商品名).value : "",
        before: Number(cellValue(found, "適用前リアル在庫")), after: Number(cellValue(found, "適用後リアル在庫")),
        processedAt: cellValue(found, "処理日時"),
      };
    }
    if (status !== REPLENISHMENT_STATUS.PENDING) throw error(409, "この補充候補は既に処理済みです");
    const productId = String(cellValue(found, "商品ID") || "").trim();
    const usItemId = String(cellValue(found, "US出品ID") || "").trim();
    const products = inventoryRowsByPid(workbook, deps.INV_COL);
    const product = products.get(productId);
    const currentUsItemId = product ? String(product.getCell(usItemIdColumn).value || "").trim() : "";
    if (!product || currentUsItemId !== usItemId) {
      const reason = !product ? "商品が在庫管理表に存在しないため要確認" : "US出品IDが現在の商品情報と一致しないため要確認";
      markNonApplied(found, REPLENISHMENT_STATUS.REVIEW, { at: deps.processedAt, actor: deps.actor, reason });
      validateProtectedSheets(workbook, { allowMissing: true });
      await atomicWriteWorkbook(workbook, deps.INVENTORY_PATH);
      throw error(409, reason);
    }
    const quantity = Number(cellValue(found, "補充候補数量"));
    const beforeValue = product.getCell(deps.INV_COL.リアル在庫).value;
    const before = Number(beforeValue);
    if (!Number.isSafeInteger(quantity) || quantity <= 0) throw error(409, "補充候補数量が不正です");
    if (!Number.isSafeInteger(before) || before < 0) throw error(409, "現在のリアル在庫が不正です");
    const after = before + quantity;
    if (!Number.isSafeInteger(after)) throw error(409, "加算後のリアル在庫が安全な整数範囲を超えます");
    const eventId = crypto.randomUUID();
    product.getCell(deps.INV_COL.リアル在庫).value = after;
    product.commit();
    setCell(found, "状態", REPLENISHMENT_STATUS.APPROVED);
    setCell(found, "処理日時", deps.processedAt);
    setCell(found, "処理者", deps.actor);
    setCell(found, "適用数量", quantity);
    setCell(found, "適用前リアル在庫", before);
    setCell(found, "適用後リアル在庫", after);
    setCell(found, "処理理由", "補充候補承認");
    setCell(found, "在庫履歴記録状態", HISTORY_PENDING);
    setCell(found, "在庫履歴イベントID", eventId);
    found.row.commit();
    validateProtectedSheets(workbook, { allowMissing: true });
    await atomicWriteWorkbook(workbook, deps.INVENTORY_PATH);
    return { candidateId: deps.candidateId, eventId, productId, productName: product.getCell(deps.INV_COL.商品名).value, quantity, before, after, processedAt: deps.processedAt };
  });
  // 一括承認のイベントは商品単位の集約履歴である。個別APIから候補1件分の履歴を
  // 再試行すると集約内容が欠けるため、一括APIだけが履歴再試行を担当する。
  if (outcome.alreadyProcessed && outcome.processingReason === "補充候補一括承認") return outcome;
  if (outcome.alreadyProcessed && outcome.historyStatus === HISTORY_RECORDED) return outcome;
  return recordApprovalHistory(deps, outcome);
}

function collectPendingBulkPlan(workbook, deps) {
  const ws = workbook.getWorksheet(REPLENISHMENT_CANDIDATES_SHEET);
  if (!ws) return { candidates: [], products: [] };
  const usItemIdColumn = requiredInventoryColumn(deps.INV_HEADERS, "US_出品ID");
  const productRows = inventoryRowsByPid(workbook, deps.INV_COL);
  const seenProducts = new Set();
  const inventorySheet = workbook.getWorksheet("在庫管理表") || workbook.worksheets[0];
  for (let rowNumber = 2; rowNumber <= inventorySheet.rowCount; rowNumber++) {
    const pid = String(inventorySheet.getRow(rowNumber).getCell(deps.INV_COL.商品ID).value || "").trim();
    if (!pid) continue;
    if (seenProducts.has(pid)) throw error(409, `商品ID「${pid}」が重複しているため一括承認できません`);
    seenProducts.add(pid);
  }
  const columns = columnMap(ws);
  const candidates = [];
  const groups = new Map();
  for (let rowNumber = 2; rowNumber <= ws.rowCount; rowNumber++) {
    const found = { row: ws.getRow(rowNumber), columns };
    if (String(cellValue(found, "状態") || "").trim() !== REPLENISHMENT_STATUS.PENDING) continue;
    const candidateId = String(cellValue(found, "補充候補ID") || "").trim();
    assertCandidateId(candidateId);
    const productId = String(cellValue(found, "商品ID") || "").trim();
    const usItemId = String(cellValue(found, "US出品ID") || "").trim();
    const quantity = Number(cellValue(found, "補充候補数量"));
    const product = productRows.get(productId);
    if (!product) throw error(409, `商品ID「${productId}」が存在しないため一括承認を中止しました`);
    const currentUsItemId = String(product.getCell(usItemIdColumn).value || "").trim();
    if (currentUsItemId !== usItemId) throw error(409, `商品ID「${productId}」のUS出品IDが一致しないため一括承認を中止しました`);
    if (!Number.isSafeInteger(quantity) || quantity <= 0) throw error(409, `商品ID「${productId}」の補充候補数量が不正です`);
    let group = groups.get(productId);
    if (!group) {
      const before = Number(product.getCell(deps.INV_COL.リアル在庫).value);
      if (!Number.isSafeInteger(before) || before < 0) throw error(409, `商品ID「${productId}」の現在リアル在庫が不正です`);
      group = { productId, product, productName: product.getCell(deps.INV_COL.商品名).value, before, total: 0, candidates: [] };
      groups.set(productId, group);
    }
    const nextTotal = group.total + quantity;
    if (!Number.isSafeInteger(nextTotal)) throw error(409, `商品ID「${productId}」の候補合計が安全な整数範囲を超えます`);
    group.total = nextTotal;
    group.candidates.push({ found, candidateId, quantity });
    candidates.push({ found, candidateId, productId, quantity });
  }
  for (const group of groups.values()) {
    group.after = group.before + group.total;
    if (!Number.isSafeInteger(group.after)) throw error(409, `商品ID「${group.productId}」の加算後リアル在庫が安全な整数範囲を超えます`);
  }
  return { candidates, products: [...groups.values()] };
}

function collectOutstandingBulkHistories(workbook, INV_COL) {
  const ws = workbook.getWorksheet(REPLENISHMENT_CANDIDATES_SHEET);
  if (!ws) return [];
  const productRows = inventoryRowsByPid(workbook, INV_COL);
  const columns = columnMap(ws);
  const events = new Map();
  for (let rowNumber = 2; rowNumber <= ws.rowCount; rowNumber++) {
    const found = { row: ws.getRow(rowNumber), columns };
    if (cellValue(found, "状態") !== REPLENISHMENT_STATUS.APPROVED
      || cellValue(found, "処理理由") !== "補充候補一括承認"
      || cellValue(found, "在庫履歴記録状態") === HISTORY_RECORDED) continue;
    const eventId = String(cellValue(found, "在庫履歴イベントID") || "");
    if (!UUID_PATTERN.test(eventId)) continue;
    const productId = String(cellValue(found, "商品ID") || "");
    const before = Number(cellValue(found, "適用前リアル在庫"));
    const after = Number(cellValue(found, "適用後リアル在庫"));
    let event = events.get(eventId);
    if (!event) {
      const product = productRows.get(productId);
      event = { eventId, productId, productName: product ? product.getCell(INV_COL.商品名).value : "", before, after, processedAt: cellValue(found, "処理日時") };
      events.set(eventId, event);
    } else {
      event.before = Math.min(event.before, before);
      event.after = Math.max(event.after, after);
    }
  }
  return [...events.values()];
}

async function recordBulkHistories(deps, histories) {
  const results = [];
  for (const history of histories) {
    try {
      await (deps.appendHistoryOnce || appendHistoryOnce)(deps.HISTORY_PATH, {
        pid: history.productId, name: history.productName, before: history.before, after: history.after,
        reason: "補充候補一括承認", at: history.processedAt,
      }, history.eventId);
      await updateBulkHistoryState({ ...deps, eventId: history.eventId, state: HISTORY_RECORDED });
      results.push({ eventId: history.eventId, status: HISTORY_RECORDED });
    } catch (_) {
      try { await updateBulkHistoryState({ ...deps, eventId: history.eventId, state: HISTORY_FAILED }); } catch (_) {}
      results.push({ eventId: history.eventId, status: HISTORY_FAILED });
    }
  }
  return results;
}

async function approveAllCandidates(deps) {
  const outcome = await withInventoryLock(async () => {
    const workbook = await deps.loadInventoryWorkbookLocked();
    validateProtectedSheets(workbook, { allowMissing: true });
    const plan = collectPendingBulkPlan(workbook, deps);
    const processedAt = deps.processedAt;
    for (const group of plan.products) {
      const eventId = crypto.randomUUID();
      group.eventId = eventId;
      group.product.getCell(deps.INV_COL.リアル在庫).value = group.after;
      group.product.commit();
      let running = group.before;
      for (const candidate of group.candidates) {
        const candidateAfter = running + candidate.quantity;
        setCell(candidate.found, "状態", REPLENISHMENT_STATUS.APPROVED);
        setCell(candidate.found, "処理日時", processedAt);
        setCell(candidate.found, "処理者", deps.actor);
        setCell(candidate.found, "適用数量", candidate.quantity);
        setCell(candidate.found, "適用前リアル在庫", running);
        setCell(candidate.found, "適用後リアル在庫", candidateAfter);
        setCell(candidate.found, "処理理由", "補充候補一括承認");
        setCell(candidate.found, "在庫履歴記録状態", HISTORY_PENDING);
        setCell(candidate.found, "在庫履歴イベントID", eventId);
        candidate.found.row.commit();
        running = candidateAfter;
      }
    }
    if (plan.candidates.length) {
      validateProtectedSheets(workbook, { allowMissing: true });
      await atomicWriteWorkbook(workbook, deps.INVENTORY_PATH);
    }
    const outstanding = collectOutstandingBulkHistories(workbook, deps.INV_COL);
    return {
      approvedCandidateCount: plan.candidates.length,
      productCount: plan.products.length,
      totalQuantity: plan.products.reduce((sum, group) => sum + group.total, 0),
      products: plan.products.map((group) => ({ productId: group.productId, before: group.before, after: group.after, quantity: group.total })),
      outstanding,
    };
  });
  const historyResults = await recordBulkHistories(deps, outcome.outstanding);
  const { outstanding, ...summary } = outcome;
  return { ...summary, historyResults, historyWarningCount: historyResults.filter((result) => result.status === HISTORY_FAILED).length };
}

async function rejectCandidate(deps) {
  assertCandidateId(deps.candidateId);
  return withInventoryLock(async () => {
    const workbook = await deps.loadInventoryWorkbookLocked();
    validateProtectedSheets(workbook, { allowMissing: true });
    const ws = workbook.getWorksheet(REPLENISHMENT_CANDIDATES_SHEET);
    if (!ws) throw error(404, "補充候補が見つかりません");
    const found = findCandidateRow(ws, deps.candidateId);
    if (!found) throw error(404, "補充候補が見つかりません");
    const status = String(cellValue(found, "状態") || "").trim();
    if (status === REPLENISHMENT_STATUS.REJECTED) return { alreadyProcessed: true, candidateId: deps.candidateId };
    if (status !== REPLENISHMENT_STATUS.PENDING) throw error(409, "この補充候補は既に処理済みです");
    markNonApplied(found, REPLENISHMENT_STATUS.REJECTED, { at: deps.processedAt, actor: deps.actor, reason: "補充候補を手動で却下" });
    validateProtectedSheets(workbook, { allowMissing: true });
    await atomicWriteWorkbook(workbook, deps.INVENTORY_PATH);
    return { alreadyProcessed: false, candidateId: deps.candidateId };
  });
}

module.exports = {
  HISTORY_PENDING, HISTORY_RECORDED, HISTORY_FAILED,
  listPendingCandidates, approveCandidate, rejectCandidate, invalidatePendingCandidates,
  approveAllCandidates, requiredInventoryColumn,
};
