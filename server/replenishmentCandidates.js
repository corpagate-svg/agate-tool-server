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
  if (outcome.alreadyProcessed && outcome.historyStatus === HISTORY_RECORDED) return outcome;
  return recordApprovalHistory(deps, outcome);
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
  requiredInventoryColumn,
};
