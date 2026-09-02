const fs = require("fs");
const ExcelJS = require("exceljs");
const { withSalesLock } = require("./orderLock");
const { atomicWriteWorkbook } = require("./inventoryLock");

function createSalesWorkbookLoaders({ LEDGER_PATH, HEADERS, writeWorkbook = atomicWriteWorkbook }) {
  // 呼び出し元がsales lockを保持している場合だけ使用する低レベルloader。
  async function loadSalesWorkbookLocked() {
    const workbook = new ExcelJS.Workbook();
    if (fs.existsSync(LEDGER_PATH)) {
      await workbook.xlsx.readFile(LEDGER_PATH);
      return workbook;
    }
    const sheet = workbook.addWorksheet("記録");
    sheet.addRow(HEADERS);
    sheet.getRow(1).font = { bold: true };
    await writeWorkbook(workbook, LEDGER_PATH);
    return workbook;
  }

  // 既存ファイルは読み取りのみ。初回生成時だけ共通sales lock内で存在を再確認する。
  async function loadSalesWorkbook() {
    if (fs.existsSync(LEDGER_PATH)) {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(LEDGER_PATH);
      return workbook;
    }
    return withSalesLock(loadSalesWorkbookLocked);
  }

  return { loadSalesWorkbook, loadSalesWorkbookLocked };
}

module.exports = { createSalesWorkbookLoaders };
