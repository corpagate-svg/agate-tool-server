// 「リアル在庫」(当社独自管理の実在庫)まわりの共通ロジック。
// eBayのCSV再取込・API同期とは完全に独立して動作する。
//
// 在庫管理表.xlsxの列位置は server.js の INV_HEADERS/INV_COL を必ず参照し、
// このファイル単体ではハードコードしない(呼び出し側から dependency injection する)。
const fs = require("fs");
const ExcelJS = require("exceljs");
const { appendHistory } = require("./inventoryHistory");

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
  await wb.xlsx.writeFile(INVENTORY_PATH);
  await appendHistory(HISTORY_PATH, {
    pid, name: row.getCell(INV_COL.商品名).value, before, after, reason, orderNo,
  });
  return { ok: true, warning: null, before, after };
}

// 手動編集による絶対値でのリアル在庫更新。確認日は自動的に今日にする(明示指定があればそちらを優先)。
async function setRealStock({ loadInventoryWorkbook, INVENTORY_PATH, HISTORY_PATH, INV_COL, pid, value, confirmedAt }) {
  const wb = await loadInventoryWorkbook();
  const ws = wb.getWorksheet("在庫管理表") || wb.worksheets[0];
  const row = findInventoryRow(ws, INV_COL, pid);
  if (!row) return { ok: false, error: "該当する商品IDが見つかりません" };
  const before = Number(row.getCell(INV_COL.リアル在庫).value) || 0;
  row.getCell(INV_COL.リアル在庫).value = value;
  row.getCell(INV_COL.リアル在庫確認日).value = confirmedAt || todayStr();
  row.commit();
  await wb.xlsx.writeFile(INVENTORY_PATH);
  await appendHistory(HISTORY_PATH, { pid, name: row.getCell(INV_COL.商品名).value, before, after: value, reason: "手動変更" });
  return { ok: true, before, after: value };
}

// 棚卸入力数量の一時保存(段階1)。リアル在庫そのものは変更しない。
async function stageStocktakeQty({ loadInventoryWorkbook, INVENTORY_PATH, INV_COL, pid, qty }) {
  const wb = await loadInventoryWorkbook();
  const ws = wb.getWorksheet("在庫管理表") || wb.worksheets[0];
  const row = findInventoryRow(ws, INV_COL, pid);
  if (!row) return { ok: false, error: "該当する商品IDが見つかりません" };
  row.getCell(INV_COL.棚卸入力数量).value = qty === null || qty === "" ? null : qty;
  row.getCell(INV_COL.棚卸入力日時).value = qty === null || qty === "" ? null : new Date().toISOString();
  row.commit();
  await wb.xlsx.writeFile(INVENTORY_PATH);
  return { ok: true };
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

// 棚卸一括確定(段階2): 棚卸入力数量をリアル在庫へ反映し、確認日を更新、履歴に記録、staging列をクリアする。
// targetPids未指定の場合は「棚卸入力数量が入っている全行」が対象。
async function confirmStocktake({ loadInventoryWorkbook, INVENTORY_PATH, HISTORY_PATH, INV_HEADERS, INV_COL, targetPids, confirmedAt }) {
  const wb = await loadInventoryWorkbook();
  const ws = wb.getWorksheet("在庫管理表") || wb.worksheets[0];
  const targetSet = targetPids && targetPids.length ? new Set(targetPids.map(String)) : null;
  const applied = [];
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
    row.getCell(INV_COL.棚卸入力日時).value = null;
    row.commit();
    applied.push({ pid, name: row.getCell(INV_COL.商品名).value, before, after });
  }
  await wb.xlsx.writeFile(INVENTORY_PATH);
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
  setRealStock,
  stageStocktakeQty,
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
