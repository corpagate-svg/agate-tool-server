// 月次為替レート(USD/JPY)を独立したExcelファイル(data/為替レート管理.xlsx)で管理する。
// 売上管理表.xlsxには一切列を追加しない(既存の列構成・月別シート走査ロジックへの影響を避けるため)。
// ネットワーク取得(exchangeRateFetcher.js)とは責務を分離し、このファイルはExcelの読み書きのみを扱う。
const fs = require("fs");
const ExcelJS = require("exceljs");
const { atomicWriteWorkbook } = require("./inventoryLock");

const SHEET_NAME = "為替レート";
const HEADERS = [
  "対象年月", "通貨ペア", "レート", "状態", "対象開始日", "対象終了日",
  "データ件数", "系列コード", "取得元", "取得日時", "確定日時", "最終更新日時", "入力方式",
];
const COL = {
  対象年月: 1, 通貨ペア: 2, レート: 3, 状態: 4, 対象開始日: 5, 対象終了日: 6,
  データ件数: 7, 系列コード: 8, 取得元: 9, 取得日時: 10, 確定日時: 11, 最終更新日時: 12, 入力方式: 13,
};
const STATUS = { PROVISIONAL: "暫定", CONFIRMED: "確定" };
const CURRENCY_PAIR = "USD/JPY";

// 在庫管理表.xlsxとは無関係の独立ファイルのため、専用のキューで直列化する
// (inventoryLock.jsのwithInventoryLockとは別キュー。atomicWriteWorkbookのみ共用する)。
let tail = Promise.resolve();
function withExchangeRateLock(fn) {
  const next = tail.then(fn, fn);
  tail = next.then(() => undefined, () => undefined);
  return next;
}

// 正式ファイルが存在しない場合のみ、ヘッダー行だけの新規ワークブックを安全に作成する。
async function loadExchangeRateWorkbook(RATE_PATH) {
  const wb = new ExcelJS.Workbook();
  if (fs.existsSync(RATE_PATH)) {
    await wb.xlsx.readFile(RATE_PATH);
    if (!wb.getWorksheet(SHEET_NAME)) {
      const ws = wb.addWorksheet(SHEET_NAME);
      ws.addRow(HEADERS);
      ws.getRow(1).font = { bold: true };
    }
    return wb;
  }
  const ws = wb.addWorksheet(SHEET_NAME);
  ws.addRow(HEADERS);
  ws.getRow(1).font = { bold: true };
  ws.views = [{ state: "frozen", ySplit: 1 }];
  ws.getColumn(1).width = 10;
  ws.getColumn(9).width = 12;
  ws.getColumn(10).width = 22;
  ws.getColumn(11).width = 22;
  ws.getColumn(12).width = 22;
  return wb;
}

function findRateRow(ws, yearMonth) {
  for (let r = 2; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    if (String(row.getCell(COL.対象年月).value || "") === yearMonth) return row;
  }
  return null;
}

function rowToRecord(row) {
  const record = {};
  Object.entries(COL).forEach(([key, idx]) => { record[key] = row.getCell(idx).value; });
  return record;
}

// 1件分の月次レートを取得する(なければnull)。読み取り専用、ロック不要。
async function getRate({ RATE_PATH, yearMonth }) {
  if (!fs.existsSync(RATE_PATH)) return null;
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(RATE_PATH);
  const ws = wb.getWorksheet(SHEET_NAME);
  if (!ws) return null;
  const row = findRateRow(ws, yearMonth);
  return row ? rowToRecord(row) : null;
}

function isConfirmedRow(row) {
  return row && String(row.getCell(COL.状態).value || "") === STATUS.CONFIRMED;
}

function writeRowValues(row, { yearMonth, rate, status, startDate, endDate, count, seriesCode, source, fetchedAt, confirmedAt, inputMethod, now }) {
  row.getCell(COL.対象年月).value = yearMonth;
  row.getCell(COL.通貨ペア).value = CURRENCY_PAIR;
  row.getCell(COL.レート).value = rate;
  row.getCell(COL.状態).value = status;
  row.getCell(COL.対象開始日).value = startDate || "";
  row.getCell(COL.対象終了日).value = endDate || "";
  row.getCell(COL.データ件数).value = count;
  row.getCell(COL.系列コード).value = seriesCode || "";
  row.getCell(COL.取得元).value = source || "";
  if (fetchedAt !== undefined) row.getCell(COL.取得日時).value = fetchedAt;
  if (confirmedAt !== undefined) row.getCell(COL.確定日時).value = confirmedAt;
  row.getCell(COL.最終更新日時).value = now;
  row.getCell(COL.入力方式).value = inputMethod || "自動";
}

// 暫定レートを新規保存/更新する。既に確定済みの月は絶対に上書きしない。
async function saveProvisionalRate({ RATE_PATH, yearMonth, rate, count, startDate, endDate, seriesCode, source, inputMethod, fetchedAt }) {
  return withExchangeRateLock(async () => {
    const wb = await loadExchangeRateWorkbook(RATE_PATH);
    const ws = wb.getWorksheet(SHEET_NAME);
    const row = findRateRow(ws, yearMonth);
    if (isConfirmedRow(row)) {
      return { ok: false, error: `${yearMonth}は既に確定済みのため、暫定レートで上書きできません`, status: 409 };
    }
    const now = new Date().toISOString();
    const target = row || ws.addRow([]);
    writeRowValues(target, {
      yearMonth, rate, status: STATUS.PROVISIONAL, startDate, endDate, count, seriesCode, source,
      fetchedAt: fetchedAt || now, confirmedAt: row ? row.getCell(COL.確定日時).value : "",
      inputMethod, now,
    });
    target.commit();
    await atomicWriteWorkbook(wb, RATE_PATH);
    return { ok: true };
  });
}

// 確定レートを保存する。既に確定済みの月への再確定は拒否する(再確定は別の明示操作が必要)。
async function confirmRate({ RATE_PATH, yearMonth, rate, count, startDate, endDate, seriesCode, source, inputMethod }) {
  return withExchangeRateLock(async () => {
    const wb = await loadExchangeRateWorkbook(RATE_PATH);
    const ws = wb.getWorksheet(SHEET_NAME);
    const row = findRateRow(ws, yearMonth);
    if (isConfirmedRow(row)) {
      return { ok: false, error: `${yearMonth}は既に確定済みです(再確定するには別の明示操作が必要です)`, status: 409 };
    }
    const now = new Date().toISOString();
    const target = row || ws.addRow([]);
    writeRowValues(target, {
      yearMonth, rate, status: STATUS.CONFIRMED, startDate, endDate, count, seriesCode, source,
      fetchedAt: row ? row.getCell(COL.取得日時).value : now, confirmedAt: now,
      inputMethod, now,
    });
    target.commit();
    await atomicWriteWorkbook(wb, RATE_PATH);
    return { ok: true };
  });
}

module.exports = {
  SHEET_NAME, HEADERS, COL, STATUS, CURRENCY_PAIR,
  loadExchangeRateWorkbook, getRate, saveProvisionalRate, confirmRate,
};
