// リアル在庫の変更履歴を追記式で記録する。
// 在庫管理表.xlsxはCSV取込・eBay同期のたびに丸ごと作り直されるため、
// 履歴は必ず独立したファイル(data/在庫変更履歴.xlsx)に保持する。
const fs = require("fs");
const ExcelJS = require("exceljs");

const HISTORY_HEADERS = ["日時", "商品ID", "商品名", "変更前数量", "変更後数量", "増減", "理由", "関連注文番号"];

async function loadHistoryWorkbook(HISTORY_PATH) {
  const wb = new ExcelJS.Workbook();
  if (fs.existsSync(HISTORY_PATH)) {
    await wb.xlsx.readFile(HISTORY_PATH);
    return wb;
  }
  const ws = wb.addWorksheet("在庫変更履歴");
  ws.addRow(HISTORY_HEADERS);
  ws.getRow(1).font = { bold: true };
  ws.views = [{ state: "frozen", ySplit: 1 }];
  ws.getColumn(1).width = 22;
  ws.getColumn(2).width = 10;
  ws.getColumn(3).width = 40;
  ws.getColumn(7).width = 20;
  ws.getColumn(8).width = 16;
  return wb;
}

// entry: { pid, name, before, after, reason, orderNo }
async function appendHistory(HISTORY_PATH, entry) {
  const wb = await loadHistoryWorkbook(HISTORY_PATH);
  const ws = wb.getWorksheet("在庫変更履歴");
  const before = entry.before === undefined || entry.before === null ? "" : entry.before;
  const after = entry.after === undefined || entry.after === null ? "" : entry.after;
  const diff = (typeof before === "number" && typeof after === "number") ? after - before : "";
  ws.addRow([
    entry.at || new Date().toISOString(),
    entry.pid || "",
    entry.name || "",
    before,
    after,
    diff,
    entry.reason || "",
    entry.orderNo || "",
  ]);
  const tmpPath = HISTORY_PATH + ".tmp";
  await wb.xlsx.writeFile(tmpPath);
  fs.renameSync(tmpPath, HISTORY_PATH);
}

async function listHistory(HISTORY_PATH, { pid } = {}) {
  if (!fs.existsSync(HISTORY_PATH)) return [];
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(HISTORY_PATH);
  const ws = wb.getWorksheet("在庫変更履歴");
  if (!ws) return [];
  const rows = [];
  ws.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const vals = HISTORY_HEADERS.map((_, i) => row.getCell(i + 1).value);
    if (pid && String(vals[1] || "") !== String(pid)) return;
    rows.push(vals);
  });
  return rows;
}

module.exports = { appendHistory, listHistory, HISTORY_HEADERS };
