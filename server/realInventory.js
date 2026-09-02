// 「リアル在庫」(当社独自管理の実在庫)まわりの共通ロジック。
// eBayのCSV再取込・API同期とは完全に独立して動作する。
//
// 在庫管理表.xlsxの列位置は server.js の INV_HEADERS/INV_COL を必ず参照し、
// このファイル単体ではハードコードしない(呼び出し側から dependency injection する)。
const fs = require("fs");
const ExcelJS = require("exceljs");
const { appendHistory } = require("./inventoryHistory");
const { withInventoryLock, atomicWriteWorkbook } = require("./inventoryLock");

// 棚卸確定時、この基準を超える差異は「異常値の可能性あり」として警告する。
// 絶対値5個以上、または現在のリアル在庫の50%以上の変動のどちらか。
const ABNORMAL_ABS_DIFF = 5;
const ABNORMAL_RATIO = 0.5;

function isAbnormalDiff(before, diff) {
  if (diff === 0) return false;
  const absDiff = Math.abs(diff);
  if (absDiff >= ABNORMAL_ABS_DIFF) return true;
  const base = Math.abs(Number(before) || 0);
  if (base > 0 && absDiff / base >= ABNORMAL_RATIO) return true;
  return false;
}

function findInventoryRow(ws, INV_COL, pid) {
  for (let r = 2; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    if (String(row.getCell(INV_COL.商品ID).value || "") === String(pid)) return row;
  }
  return null;
}

function todayStr() {
  return new Date().toISOString().slice(0, 10);
}

// 注文の登録・編集・削除・キャンセルからの自動増減。確認日は変更しない。
// 該当商品IDが見つからない場合はエラーにせず、warningのみを返す(注文処理は継続させる)。
async function adjustRealStock({ loadInventoryWorkbook, INVENTORY_PATH, HISTORY_PATH, INV_COL, pid, delta, reason, orderNo }) {
  if (!pid || !delta) return { ok: true, warning: null };
  const result = await withInventoryLock(async () => {
    const wb = await loadInventoryWorkbook();
    const ws = wb.getWorksheet("在庫管理表") || wb.worksheets[0];
    const row = findInventoryRow(ws, INV_COL, pid);
    if (!row) {
      return { ok: false, warning: `商品ID「${pid}」が在庫管理表に見つからないため、リアル在庫は変更されていません(${reason})` };
    }
    const before = Number(row.getCell(INV_COL.リアル在庫).value) || 0;
    const after = before + delta;
    row.getCell(INV_COL.リアル在庫).value = after;
    row.commit();
    await atomicWriteWorkbook(wb, INVENTORY_PATH);
    return { ok: true, warning: null, before, after, name: row.getCell(INV_COL.商品名).value };
  });
  if (result.ok) {
    await appendHistory(HISTORY_PATH, {
      pid, name: result.name, before: result.before, after: result.after, reason, orderNo,
    });
  }
  return result;
}

// 棚卸入力数量の検証。JSON number型で0以上の整数のみ許可し、空文字/null/undefinedは
// 「クリア」として扱う。それ以外(数字の文字列・空白文字列・真偽値・配列・NaN・Infinity・
// 負数・小数)は不正として拒否する(既存値を消さないため)。
// Number()による強制変換より前にtypeofで弾くことで、"0"や"10"、true/false、[]、[5]のような
// 「Number()に通せば数値になってしまう」値を確実に拒否する。
function validateStocktakeQty(raw) {
  if (raw === "" || raw === null || raw === undefined) return { ok: true, value: null };
  if (typeof raw !== "number") return { ok: false, error: "棚卸入力数量は数値で指定してください" };
  if (!Number.isFinite(raw)) return { ok: false, error: "棚卸入力数量は数値で指定してください" };
  if (!Number.isInteger(raw)) return { ok: false, error: "棚卸入力数量は整数で指定してください" };
  if (raw < 0) return { ok: false, error: "棚卸入力数量は0以上の整数で指定してください" };
  return { ok: true, value: raw };
}

// リアル在庫を絶対値で更新する純粋な変更処理(読み込み・書き込みは呼び出し側のトランザクションで行う)。
// 確認日は自動的に今日にする(明示指定があればそちらを優先)。
function applyRealStockToRow(row, INV_COL, value, confirmedAt) {
  const before = Number(row.getCell(INV_COL.リアル在庫).value) || 0;
  row.getCell(INV_COL.リアル在庫).value = value;
  row.getCell(INV_COL.リアル在庫確認日).value = confirmedAt || todayStr();
  return { before, after: value };
}

// 棚卸入力数量を反映する純粋な変更処理(読み込み・書き込みは呼び出し側のトランザクションで行う)。
// value は validateStocktakeQty() で検証済みの値(null=クリア、それ以外は0以上の整数)。
// 数量が正常に(0を含めて)保存された場合のみ「棚卸チェック」を自動的に立てる。
// このチェックは「実際にその商品を数えたか」を管理する独立フラグのため、
// 未入力に戻しても自動では解除しない(解除は棚卸チェックの一括解除操作のみで行う)。
function applyStagedQtyToRow(row, INV_COL, value) {
  const cleared = value === null;
  row.getCell(INV_COL.棚卸入力数量).value = cleared ? null : value;
  row.getCell(INV_COL.棚卸入力日時).value = cleared ? null : new Date().toISOString();
  if (!cleared) row.getCell(INV_COL.棚卸チェック).value = true;
}

// ある商品が「今回の棚卸で既に一括確定済みか」を判定する。
// confirmStocktake()は確定時に棚卸入力数量だけをnullにし棚卸入力日時は保持する(棚卸チェックも不変)ため、
// 「棚卸入力数量が空 かつ 棚卸入力日時が入っている」の組み合わせは一括確定を経由した場合にしか起こり得ない。
// (数量を空欄に戻して保存した場合はapplyStagedQtyToRowが数量・日時を同時にnullにするため、この組み合わせにはならない)
function isStocktakeConfirmed(row, INV_COL) {
  const staged = row.getCell(INV_COL.棚卸入力数量).value;
  const stagedAt = row.getCell(INV_COL.棚卸入力日時).value;
  const stagedEmpty = staged === null || staged === undefined || staged === "";
  const hasStagedAt = stagedAt !== null && stagedAt !== undefined && stagedAt !== "";
  return stagedEmpty && hasStagedAt;
}

// 1商品だけ「今回の棚卸をしていない状態」へ戻す(一括確定前の商品のみが対象)。
// 棚卸入力数量・棚卸チェック・棚卸入力日時の3列だけをクリアし、リアル在庫・リアル在庫確認日・
// 仕入価格・商品情報など他の列は一切変更しない(リアル在庫の自動復元・逆算は行わない)。
// 既に一括確定済み(isStocktakeConfirmed)の場合は、呼び出し側がHTTP 409を返せるようエラーで拒否する。
async function undoStocktakeEntry({ loadInventoryWorkbook, INVENTORY_PATH, INV_COL, pid }) {
  return withInventoryLock(async () => {
    const wb = await loadInventoryWorkbook();
    const ws = wb.getWorksheet("在庫管理表") || wb.worksheets[0];
    const row = findInventoryRow(ws, INV_COL, pid);
    if (!row) return { ok: false, status: 404, error: "該当する商品IDが見つかりません" };
    if (isStocktakeConfirmed(row, INV_COL)) {
      return { ok: false, status: 409, error: "この商品は既に一括確定済みのため、未確認に戻せません" };
    }
    row.getCell(INV_COL.棚卸入力数量).value = null;
    row.getCell(INV_COL.棚卸入力日時).value = null;
    row.getCell(INV_COL.棚卸チェック).value = null;
    row.commit();
    await atomicWriteWorkbook(wb, INVENTORY_PATH);
    return { ok: true };
  });
}

// 棚卸チェックの一括解除(手動運用のみ、次回棚卸サイクル開始時の明示的リセット)。
// 棚卸チェックと棚卸入力日時をあわせてクリアする(両方とも「今回の棚卸で確認済みか」を表す
// 対の情報のため)。リアル在庫・棚卸入力数量・日本語商品名など他の列は一切変更しない。
async function resetStocktakeChecks({ loadInventoryWorkbook, INVENTORY_PATH, INV_COL }) {
  return withInventoryLock(async () => {
    const wb = await loadInventoryWorkbook();
    const ws = wb.getWorksheet("在庫管理表") || wb.worksheets[0];
    let reset = 0;
    for (let r = 2; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      if (!row.getCell(INV_COL.棚卸チェック).value) continue;
      row.getCell(INV_COL.棚卸チェック).value = null;
      row.getCell(INV_COL.棚卸入力日時).value = null;
      row.commit();
      reset++;
    }
    await atomicWriteWorkbook(wb, INVENTORY_PATH);
    return { reset };
  });
}

function rowValues(row, INV_HEADERS) {
  const vals = {};
  INV_HEADERS.forEach((h, i) => { vals[h] = row.getCell(i + 1).value; });
  return vals;
}

// 「相違」ページ用: 2系統の比較のどちらかにズレがある商品のみを返す。
//   (1) リアル在庫 ↔ US在庫 (実地棚卸とeBayの申告値がズレていないか)
//   (2) US在庫 ↔ UK/AU在庫 (eBaymagの国間同期が壊れていないか。USを基準とする)
// 出品されていない国はその国の比較を除外する。UK/AUの判定はリアル在庫とではなく、
// 必ずUS在庫を基準に行う(リアル在庫が一致していてもUS↔UK/AUがズレていれば相違として扱う)。
function computeDiscrepancies({ ws, INV_HEADERS, INV_COL }) {
  const results = [];
  for (let r = 2; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const v = rowValues(row, INV_HEADERS);
    const real = v["リアル在庫"];
    if (real === null || real === undefined || real === "") continue;
    const realNum = Number(real) || 0;

    const hasSite = (idKey) => v[idKey] !== null && v[idKey] !== undefined && v[idKey] !== "";
    const hasUS = hasSite("US_出品ID");
    const usNum = hasUS ? (Number(v["在庫数(現物)"]) || 0) : null;

    const countries = [];
    let anyMismatch = false;

    if (hasUS) {
      const mismatch = realNum !== usNum;
      if (mismatch) anyMismatch = true;
      countries.push({ site: "US", value: usNum, mismatch });
    }
    for (const [site, idKey, valKey] of [["UK", "UK_出品ID", "UK在庫数"], ["AU", "AU_出品ID", "AU在庫数"]]) {
      if (!hasSite(idKey)) continue;
      const val = Number(v[valKey]) || 0;
      // US↔UK/AUの比較。US自体が未出品(通常あり得ない)の場合は判定不能として不一致扱いにしない。
      const mismatch = hasUS ? usNum !== val : false;
      if (mismatch) anyMismatch = true;
      countries.push({ site, value: val, mismatch });
    }

    if (!anyMismatch) continue;
    results.push({
      商品ID: v["商品ID"],
      商品名: v["商品名"],
      リアル在庫: realNum,
      リアル在庫確認日: v["リアル在庫確認日"] || null,
      unconfirmed: !v["リアル在庫確認日"],
      // US eBay商品ページへのリンク表示専用。相違判定には使わない。
      US_出品ID: hasUS ? v["US_出品ID"] : null,
      countries,
    });
  }
  return results;
}

// 棚卸一括確定の事前プレビュー: 棚卸入力数量が入っている行を対象に、差異件数・異常値件数を計算する。
function previewStocktake({ ws, INV_HEADERS }) {
  const rows = [];
  for (let r = 2; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const v = rowValues(row, INV_HEADERS);
    const staged = v["棚卸入力数量"];
    if (staged === null || staged === undefined || staged === "") continue;
    const before = Number(v["リアル在庫"]) || 0;
    const after = Number(staged) || 0;
    const diff = after - before;
    rows.push({
      商品ID: v["商品ID"],
      商品名: v["商品名"],
      リアル在庫: before,
      棚卸入力数量: after,
      差異: diff,
      異常値: isAbnormalDiff(before, diff),
      棚卸入力日時: v["棚卸入力日時"] || null,
    });
  }
  const increased = rows.filter((r) => r.差異 > 0).length;
  const decreased = rows.filter((r) => r.差異 < 0).length;
  const abnormal = rows.filter((r) => r.異常値).length;
  return {
    targetCount: rows.length,
    diffCount: rows.filter((r) => r.差異 !== 0).length,
    increased,
    decreased,
    abnormalCount: abnormal,
    rows,
  };
}

// 棚卸一括確定(段階2): 棚卸入力数量をリアル在庫へ反映し、確認日を更新、履歴に記録し、
// 棚卸入力数量(staging値)のみクリアする。棚卸チェック・棚卸入力日時は「今回の棚卸で確認済み」の
// 履歴として一括確定後も残す(次回棚卸開始時の「棚卸チェックをすべて解除」でのみクリアする)。
// targetPids未指定の場合は「棚卸入力数量が入っている全行」が対象。
async function confirmStocktake({ loadInventoryWorkbook, INVENTORY_PATH, HISTORY_PATH, INV_HEADERS, INV_COL, targetPids, confirmedAt }) {
  const applied = await withInventoryLock(async () => {
    const wb = await loadInventoryWorkbook();
    const ws = wb.getWorksheet("在庫管理表") || wb.worksheets[0];
    const targetSet = targetPids && targetPids.length ? new Set(targetPids.map(String)) : null;
    const rows = [];
    const date = confirmedAt || todayStr();
    for (let r = 2; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      const pid = String(row.getCell(INV_COL.商品ID).value || "");
      const staged = row.getCell(INV_COL.棚卸入力数量).value;
      if (staged === null || staged === undefined || staged === "") continue;
      if (targetSet && !targetSet.has(pid)) continue;
      const before = Number(row.getCell(INV_COL.リアル在庫).value) || 0;
      const after = Number(staged) || 0;
      row.getCell(INV_COL.リアル在庫).value = after;
      row.getCell(INV_COL.リアル在庫確認日).value = date;
      row.getCell(INV_COL.棚卸入力数量).value = null;
      row.commit();
      rows.push({ pid, name: row.getCell(INV_COL.商品名).value, before, after });
    }
    await atomicWriteWorkbook(wb, INVENTORY_PATH);
    return rows;
  });
  for (const a of applied) {
    await appendHistory(HISTORY_PATH, { pid: a.pid, name: a.name, before: a.before, after: a.after, reason: "棚卸確定" });
  }
  return { confirmed: applied.length, applied };
}

// 決算日から見て「実地確認が古い/未確認」の商品一覧(未確認チェック)。
function closingChecklist({ ws, INV_HEADERS, asOf }) {
  const asOfMs = Date.parse(asOf);
  const rows = [];
  for (let r = 2; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const v = rowValues(row, INV_HEADERS);
    const confirmedAt = v["リアル在庫確認日"];
    const confirmedMs = confirmedAt ? Date.parse(String(confirmedAt).slice(0, 10)) : null;
    const stale = !confirmedMs || confirmedMs < asOfMs;
    if (!stale) continue;
    rows.push({
      商品ID: v["商品ID"],
      商品名: v["商品名"],
      リアル在庫: v["リアル在庫"],
      リアル在庫確認日: confirmedAt || null,
    });
  }
  return rows;
}

// 棚卸資産一覧(決算エクスポート用)の行データを組み立てる。
function buildClosingRows({ ws, INV_HEADERS }) {
  const rows = [];
  let totalValue = 0;
  for (let r = 2; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const v = rowValues(row, INV_HEADERS);
    const qty = Number(v["リアル在庫"]) || 0;
    const price = Number(v["仕入価格(円)"]) || 0;
    const value = qty * price;
    totalValue += value;
    rows.push({
      商品ID: v["商品ID"],
      商品名: v["商品名"],
      リアル在庫: qty,
      仕入価格: price,
      評価額: value,
      リアル在庫確認日: v["リアル在庫確認日"] || null,
    });
  }
  return { rows, totalValue };
}

async function exportClosingXlsx({ ws, INV_HEADERS, asOf }) {
  const { rows, totalValue } = buildClosingRows({ ws, INV_HEADERS });
  const wb = new ExcelJS.Workbook();
  const out = wb.addWorksheet("棚卸資産");
  out.addRow(["基準日", asOf]);
  out.addRow([]);
  out.addRow(["商品ID", "商品名", "リアル在庫", "仕入価格(円)", "評価額(円)", "リアル在庫確認日"]);
  out.getRow(3).font = { bold: true };
  rows.forEach((r) => {
    out.addRow([r.商品ID, r.商品名, r.リアル在庫, r.仕入価格, r.評価額, r.リアル在庫確認日]);
  });
  out.addRow([]);
  out.addRow(["", "", "", "合計評価額", totalValue]);
  [10, 45, 12, 14, 14, 16].forEach((w, i) => { out.getColumn(i + 1).width = w; });
  return { workbook: wb, totalValue, rowCount: rows.length };
}

module.exports = {
  findInventoryRow,
  adjustRealStock,
  validateStocktakeQty,
  applyRealStockToRow,
  applyStagedQtyToRow,
  isStocktakeConfirmed,
  undoStocktakeEntry,
  resetStocktakeChecks,
  computeDiscrepancies,
  previewStocktake,
  confirmStocktake,
  closingChecklist,
  buildClosingRows,
  exportClosingXlsx,
  isAbnormalDiff,
  ABNORMAL_ABS_DIFF,
  ABNORMAL_RATIO,
};
