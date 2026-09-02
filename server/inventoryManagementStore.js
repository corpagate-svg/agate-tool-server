const fs = require("fs");
const ExcelJS = require("exceljs");
const { withInventoryLock, atomicWriteWorkbook } = require("./inventoryLock");
const {
  ensureProtectedSheets,
  validateProtectedSheets,
  addOrderLine,
  addItemMapping,
  readOrderLines,
  readItemMappings,
} = require("./inventoryProtectedSheets");

function createInventoryWorkbookLoaders({ INVENTORY_PATH, INV_HEADERS, writeWorkbook = atomicWriteWorkbook }) {
  // 呼び出し元がinventory lockを保持している場合だけ使用する低レベルloader。
  async function loadInventoryWorkbookLocked() {
    const workbook = new ExcelJS.Workbook();
    if (fs.existsSync(INVENTORY_PATH)) {
      await workbook.xlsx.readFile(INVENTORY_PATH);
      return workbook;
    }
    const inventorySheet = workbook.addWorksheet("在庫管理表");
    inventorySheet.addRow(INV_HEADERS);
    inventorySheet.getRow(1).font = { bold: true };
    ensureProtectedSheets(workbook);
    await writeWorkbook(workbook, INVENTORY_PATH);
    return workbook;
  }

  // 既存ファイルの読込は非変更処理。初回生成だけ共有lock内で存在確認からやり直す。
  async function loadInventoryWorkbook() {
    if (fs.existsSync(INVENTORY_PATH)) {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(INVENTORY_PATH);
      return workbook;
    }
    return withInventoryLock(loadInventoryWorkbookLocked);
  }

  return { loadInventoryWorkbook, loadInventoryWorkbookLocked };
}

async function initializeInventoryManagementSheets({ loadInventoryWorkbookLocked, INVENTORY_PATH }) {
  return withInventoryLock(async () => {
    const workbook = await loadInventoryWorkbookLocked();
    ensureProtectedSheets(workbook);
    const result = validateProtectedSheets(workbook);
    await atomicWriteWorkbook(workbook, INVENTORY_PATH);
    return result;
  });
}

async function saveOrderLine({ loadInventoryWorkbookLocked, INVENTORY_PATH, record }) {
  return withInventoryLock(async () => {
    const workbook = await loadInventoryWorkbookLocked();
    ensureProtectedSheets(workbook);
    validateProtectedSheets(workbook);
    const key = addOrderLine(workbook, record);
    validateProtectedSheets(workbook);
    await atomicWriteWorkbook(workbook, INVENTORY_PATH);
    return key;
  });
}

async function saveItemMapping({ loadInventoryWorkbookLocked, INVENTORY_PATH, record }) {
  return withInventoryLock(async () => {
    const workbook = await loadInventoryWorkbookLocked();
    ensureProtectedSheets(workbook);
    validateProtectedSheets(workbook);
    const key = addItemMapping(workbook, record);
    validateProtectedSheets(workbook);
    await atomicWriteWorkbook(workbook, INVENTORY_PATH);
    return key;
  });
}

async function loadInventoryManagementData({ loadInventoryWorkbook }) {
  const workbook = await loadInventoryWorkbook();
  // 両方ない旧Workbookだけは空データとして許可する。一方だけ存在する場合や、
  // 行内容・キー・状態に矛盾がある場合は、壊れた管理データを返さずここで停止する。
  validateProtectedSheets(workbook, { allowMissing: true });
  return {
    orderLines: readOrderLines(workbook),
    itemMappings: readItemMappings(workbook),
  };
}

module.exports = {
  createInventoryWorkbookLoaders,
  initializeInventoryManagementSheets,
  saveOrderLine,
  saveItemMapping,
  loadInventoryManagementData,
};
