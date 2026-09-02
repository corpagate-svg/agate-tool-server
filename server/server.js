const http = require("http");
const crypto = require("crypto");
const path = require("path");
const fs = require("fs");
const ExcelJS = require("exceljs");
const { sendJson, sendHtml, contentDispositionAttachment } = require("./httpUtil");
const { handleEbayConnect, handleEbayCallback, handleEbayInventory, createInventoryRebuildHandler } = require("./ebay/routes");
const { rebuildInventoryFromRecords } = require("./inventoryRebuild");
const realInv = require("./realInventory");
const { listHistory, appendHistory } = require("./inventoryHistory");
const { withInventoryLock, atomicWriteWorkbook, atomicWriteBuffer } = require("./inventoryLock");
const { resolveManualInventoryLink, canEditOrderProductId } = require("./orderInventorySafety");
const { parseOrderText } = require("./orderParser");
const { processOrderInventoryTransaction } = require("./orderRegistration");
const { withSalesLock } = require("./orderLock");
const {
  validateInventorySheet,
  validateProtectedSheets,
  buildProtectedImportWorkbook,
  readOrderLines,
  ORDER_LINE_STATUS,
} = require("./inventoryProtectedSheets");
const { createInventoryWorkbookLoaders } = require("./inventoryManagementStore");
const { createSalesWorkbookLoaders } = require("./salesWorkbookStore");
const { resolveOrderLineTransaction } = require("./manualOrderResolution");
const { editManagedOrderTransaction, deleteManagedOrdersTransaction } = require("./managedOrderMutation");

const PORT = process.env.PORT || 3000;
const API_TOKEN = process.env.API_TOKEN || "";
// 通常は従来どおり data/ を使う。HTTP統合テストでは実Excelへ触れないよう隔離先を明示できる。
const DATA_DIR = process.env.AGATE_DATA_DIR ? path.resolve(process.env.AGATE_DATA_DIR) : path.join(__dirname, "..", "data");
const LEDGER_PATH = path.join(DATA_DIR, "売上管理表.xlsx");
const INVENTORY_PATH = path.join(DATA_DIR, "在庫管理表.xlsx");
const HISTORY_PATH = path.join(DATA_DIR, "在庫変更履歴.xlsx");
const SNAPSHOT_DIR = path.join(DATA_DIR, "snapshots");
const INVENTORY_IMPORT_BACKUP_DIR = path.join(DATA_DIR, "backups", "inventory-import");
const ORDER_PARSER_PATH = path.join(__dirname, "orderParser.js");

const HEADERS = [
  "注文番号", "日付", "サイト", "商品メモ", "数量", "商品ID",
  "収益USD", "ドル円レート", "収益円", "手数料(円)",
  "仕入原価(円)", "送料(円)", "梱包費(円)", "最終利益(円)", "利益率",
];
const COL = { 注文番号: 1, 日付: 2, サイト: 3, 商品メモ: 4, 数量: 5, 商品ID: 6, 収益USD: 7, ドル円レート: 8, 収益円: 9, 手数料: 10, 仕入原価: 11, 送料: 12, 梱包費: 13, 最終利益: 14, 利益率: 15 };

const INV_HEADERS = [
  "商品ID", "商品名", "バリエーション詳細", "在庫数(現物)", "出品国数", "在庫数不一致",
  "US_出品ID", "US価格(USD)", "US累計売却数", "UK_出品ID", "UK価格(GBP)", "UK累計売却数",
  "AU_出品ID", "AU価格(AUD)", "AU累計売却数", "仕入価格(円)", "仕入日", "仕入先", "備考",
  "UK在庫数", "AU在庫数", "リアル在庫", "リアル在庫確認日", "棚卸入力数量", "棚卸入力日時",
  "画像URL", "日本語商品名", "SKU", "棚卸チェック",
];
const INV_COL = {
  商品ID: 1, 商品名: 2, バリエーション詳細: 3, 在庫数: 4, 出品国数: 5, 在庫数不一致: 6, 仕入価格: 16, 仕入日: 17, 仕入先: 18, 備考: 19,
  UK在庫数: 20, AU在庫数: 21, リアル在庫: 22, リアル在庫確認日: 23, 棚卸入力数量: 24, 棚卸入力日時: 25,
  画像URL: 26, 日本語商品名: 27, SKU: 28, 棚卸チェック: 29,
};

if (!API_TOKEN) {
  console.error("API_TOKEN が設定されていません(.env を確認してください)。起動を中止します。");
  process.exit(1);
}

fs.mkdirSync(DATA_DIR, { recursive: true });

function isDataSheet(name) {
  return !name.includes("について") && !name.includes("要確認");
}

function styleHeaderRow(ws, headers) {
  ws.addRow(headers);
  ws.getRow(1).font = { bold: true };
}

const { loadSalesWorkbook: loadWorkbook, loadSalesWorkbookLocked: loadWorkbookLocked } = createSalesWorkbookLoaders({
  LEDGER_PATH,
  HEADERS,
});

const { loadInventoryWorkbook, loadInventoryWorkbookLocked } = createInventoryWorkbookLoaders({
  INVENTORY_PATH,
  INV_HEADERS,
});

const handleEbayInventoryRebuild = createInventoryRebuildHandler({ INV_HEADERS, INVENTORY_PATH, loadInventoryWorkbook: loadInventoryWorkbookLocked });

function dataSheets(wb) {
  return wb.worksheets.filter((ws) => isDataSheet(ws.name));
}

function monthSheetName(dateStr) {
  const m = /^(\d{4})-(\d{1,2})/.exec(String(dateStr || ""));
  if (!m) return null;
  return `${m[1]}年${Number(m[2])}月`;
}

function getOrCreateMonthSheet(wb, dateStr) {
  const name = monthSheetName(dateStr) || "記録";
  let ws = wb.getWorksheet(name);
  if (!ws) {
    ws = wb.addWorksheet(name);
    styleHeaderRow(ws, HEADERS);
  }
  return ws;
}

function isAuthorized(req) {
  const auth = req.headers["authorization"] || "";
  const m = /^Bearer (.+)$/.exec(auth);
  let given = null;
  if (m) {
    given = m[1];
  } else if (req.method === "GET") {
    // GET専用: 外部ツールが簡単にポーリングできるよう、?token=... でも認証可能にする
    try {
      given = new URL(req.url, "http://localhost").searchParams.get("token");
    } catch (e) {
      given = null;
    }
  }
  if (!given) return false;
  const givenBuf = Buffer.from(given);
  const expected = Buffer.from(API_TOKEN);
  if (givenBuf.length !== expected.length) return false;
  return crypto.timingSafeEqual(givenBuf, expected);
}

function parseCsv(text) {
  const rows = [];
  let row = [], field = "", inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inQuotes) {
      if (c === '"') {
        if (text[i + 1] === '"') { field += '"'; i++; }
        else inQuotes = false;
      } else field += c;
    } else if (c === '"') {
      inQuotes = true;
    } else if (c === ",") {
      row.push(field); field = "";
    } else if (c === "\r") {
      // skip
    } else if (c === "\n") {
      row.push(field); rows.push(row); row = []; field = "";
    } else {
      field += c;
    }
  }
  if (field.length || row.length) { row.push(field); rows.push(row); }
  return rows;
}

function csvToObjects(text) {
  const rows = parseCsv(text).filter((r) => !(r.length === 1 && r[0] === ""));
  if (!rows.length) return [];
  const headers = rows[0];
  return rows.slice(1).map((r) => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = r[i] !== undefined ? r[i] : ""; });
    return obj;
  });
}

function readRawBody(req, maxBytes) {
  return new Promise((resolve, reject) => {
    let size = 0;
    const chunks = [];
    req.on("data", (chunk) => {
      size += chunk.length;
      if (size > maxBytes) {
        reject(new Error("payload too large"));
        req.destroy();
        return;
      }
      chunks.push(chunk);
    });
    req.on("end", () => resolve(Buffer.concat(chunks)));
    req.on("error", reject);
  });
}

function recompute(revenueJpy, fee, cost, shipping, packing) {
  if (cost === null || shipping === null || packing === null) return { profit: "", margin: "" };
  const profit = Math.round(revenueJpy - fee - cost - shipping - packing);
  const margin = revenueJpy !== 0 ? Number((profit / revenueJpy).toFixed(4)) : "";
  return { profit, margin };
}

function numOrNull(v) {
  return v !== undefined && v !== null && v !== "" && Number.isFinite(Number(v)) ? Number(v) : null;
}

async function handleAddOrder(req, res) {
  let body;
  try {
    body = JSON.parse((await readRawBody(req, 1024 * 1024)).toString("utf8"));
  } catch (e) {
    return sendJson(res, 400, { error: "リクエストの内容を読み取れませんでした" });
  }

  const required = ["注文番号", "日付", "サイト", "商品メモ", "収益USD", "ドル円レート"];
  for (const key of required) {
    if (body[key] === undefined || body[key] === null || body[key] === "") {
      return sendJson(res, 400, { error: `${key} は必須です` });
    }
  }

  const usd = Number(body["収益USD"]);
  const rate = Number(body["ドル円レート"]);
  if (!Number.isFinite(usd) || !Number.isFinite(rate)) {
    return sendJson(res, 400, { error: "収益USD / ドル円レート は数値で指定してください" });
  }
  const revenueJpy = usd * rate;
  const fee = Math.round(revenueJpy * 0.03);
  const cost = numOrNull(body["仕入原価円"]);
  const shipping = numOrNull(body["送料円"]);
  const packing = numOrNull(body["梱包費円"]);
  const { profit, margin } = recompute(revenueJpy, fee, cost, shipping, packing);

  const salesWorkbook = await loadWorkbookLocked(); // 呼び出し元がsales lockを保持する。
  const hasRawOrderText = Object.prototype.hasOwnProperty.call(body, "注文貼付テキスト");

  // 貼り付け原文がない旧方式は従来互換。第1段階の安全ガードにより自動在庫連動は行われない。
  if (!hasRawOrderText) {
    if (findOrderRowByNo(salesWorkbook, body["注文番号"])) {
      return sendJson(res, 409, { error: `注文番号「${body["注文番号"]}」は既に登録されています(二重登録防止のため中止しました)` });
    }
    const inventoryLink = resolveManualInventoryLink(body);
    const ws = getOrCreateMonthSheet(salesWorkbook, body["日付"]);
    ws.addRow([
      body["注文番号"], body["日付"], body["サイト"], body["商品メモ"], numOrNull(body["数量"]) === null ? "" : numOrNull(body["数量"]), inventoryLink.pid,
      usd, rate, Math.round(revenueJpy), fee,
      cost === null ? "" : cost, shipping === null ? "" : shipping, packing === null ? "" : packing,
      profit, margin,
    ]);
    await atomicWriteWorkbook(salesWorkbook, LEDGER_PATH);
    return sendJson(res, 200, { status: "ok", 収益円: Math.round(revenueJpy), 手数料円: fee, 最終利益円: profit, 利益率: margin, warning: inventoryLink.reason || null });
  }

  const parsed = parseOrderText(body["注文貼付テキスト"]);
  if (!parsed.orderNo || String(parsed.orderNo) !== String(body["注文番号"])) {
    return sendJson(res, 400, { error: "貼り付け原文の注文番号と登録内容が一致しません" });
  }
  const salesOrderExists = Boolean(findOrderRowByNo(salesWorkbook, parsed.orderNo));
  const inventoryResult = await processOrderInventoryTransaction({
    withInventoryLock, loadInventoryWorkbookLocked, atomicWriteWorkbook, INVENTORY_PATH,
    parsed, salesOrderExists, INV_HEADERS,
  });

  if (inventoryResult.legacyOrder) {
    return sendJson(res, 409, { error: "売上登録済みですが注文明細がない旧方式の注文です。自動在庫処理は行いません" });
  }
  if (inventoryResult.alreadyRegistered) {
    return sendJson(res, 200, { status: "already_registered", inventory: inventoryResult });
  }

  // inventory保存後にsales保存が失敗しても、retry時は保存済み明細一致により在庫再減算されない。
  const ws = getOrCreateMonthSheet(salesWorkbook, body["日付"] || parsed.date);
  const salesSite = parsed.parseStatus === "OK" ? parsed.site : body["サイト"];
  const salesNote = parsed.items.length ? parsed.note : body["商品メモ"];
  const salesQty = parsed.items.length ? parsed.quantityTotal : numOrNull(body["数量"]);
  ws.addRow([
    parsed.orderNo, body["日付"] || parsed.date, salesSite, salesNote,
    salesQty === null ? "" : salesQty, "",
    usd, rate, Math.round(revenueJpy), fee,
    cost === null ? "" : cost, shipping === null ? "" : shipping, packing === null ? "" : packing,
    profit, margin,
  ]);
  await atomicWriteWorkbook(salesWorkbook, LEDGER_PATH);
  sendJson(res, 200, {
    status: inventoryResult.retry ? "recovered" : "ok",
    収益円: Math.round(revenueJpy), 手数料円: fee, 最終利益円: profit, 利益率: margin,
    inventory: inventoryResult,
  });
}

function findOrderRowByNo(wb, orderNo) {
  for (const ws of dataSheets(wb)) {
    for (let r = 2; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      if (String(row.getCell(COL.注文番号).value || "") === String(orderNo)) return row;
    }
  }
  return null;
}

async function handlePatchOrder(req, res) {
  let body;
  try {
    body = JSON.parse((await readRawBody(req, 1024 * 1024)).toString("utf8"));
  } catch (e) {
    return sendJson(res, 400, { error: "リクエストの内容を読み取れませんでした" });
  }
  const orderNo = body["注文番号"];
  if (!orderNo) return sendJson(res, 400, { error: "注文番号 は必須です" });

  const wb = await loadWorkbookLocked();
  const row = findOrderRowByNo(wb, orderNo);
  if (!row) return sendJson(res, 404, { error: "該当する注文番号が見つかりません" });

  const hasRawEdit = Object.prototype.hasOwnProperty.call(body, "注文貼付テキスト");
  const parsedEdit = hasRawEdit ? parseOrderText(body["注文貼付テキスト"]) : null;
  const managedResult = await editManagedOrderTransaction({
    withInventoryLock, loadInventoryWorkbookLocked, atomicWriteWorkbook, INVENTORY_PATH,
    orderNo, parsed: parsedEdit,
    quantityOverride: !hasRawEdit && "数量" in body && !("商品ID" in body) ? body["数量"] : undefined,
    INV_HEADERS,
  });

  if (managedResult.managed && "商品ID" in body && !hasRawEdit) {
    return sendJson(res, 409, { error: "管理対象注文の商品変更は、Item IDを含む注文詳細を再解析して行ってください。在庫は変更されていません" });
  }

  const oldPid = String(row.getCell(COL.商品ID).value || "");
  const oldQty = numOrNull(row.getCell(COL.数量).value) === null ? 1 : numOrNull(row.getCell(COL.数量).value);
  const requestedPid = "商品ID" in body ? String(body["商品ID"] || "") : oldPid;
  if (!managedResult.managed && !canEditOrderProductId(oldPid, requestedPid)) {
    return sendJson(res, 400, { error: "商品IDが空欄の注文には、明細確認なしでPxxxxを後付けできません" });
  }

  const cost = "仕入原価円" in body ? numOrNull(body["仕入原価円"]) : numOrNull(row.getCell(COL.仕入原価).value);
  const shipping = "送料円" in body ? numOrNull(body["送料円"]) : numOrNull(row.getCell(COL.送料).value);
  const packing = "梱包費円" in body ? numOrNull(body["梱包費円"]) : numOrNull(row.getCell(COL.梱包費).value);
  const usd = "収益USD" in body ? numOrNull(body["収益USD"]) : numOrNull(row.getCell(COL.収益USD).value);
  const rate = "ドル円レート" in body ? numOrNull(body["ドル円レート"]) : numOrNull(row.getCell(COL.ドル円レート).value);
  const revenueJpy = usd !== null && rate !== null ? usd * rate : Number(row.getCell(COL.収益円).value) || 0;
  const fee = Math.round(revenueJpy * 0.03);
  const { profit, margin } = recompute(revenueJpy, fee, cost, shipping, packing);

  if (usd !== null) row.getCell(COL.収益USD).value = usd;
  if (rate !== null) row.getCell(COL.ドル円レート).value = rate;
  if ("商品メモ" in body) row.getCell(COL.商品メモ).value = body["商品メモ"];
  if (managedResult.managed && hasRawEdit) {
    row.getCell(COL.サイト).value = parsedEdit.site;
    row.getCell(COL.商品メモ).value = parsedEdit.note;
    row.getCell(COL.数量).value = parsedEdit.quantityTotal;
    row.getCell(COL.商品ID).value = "";
  } else if ("数量" in body) {
    const q = numOrNull(body["数量"]);
    row.getCell(COL.数量).value = q === null ? "" : q;
  }
  if (!managedResult.managed && "商品ID" in body) row.getCell(COL.商品ID).value = body["商品ID"] || "";
  row.getCell(COL.収益円).value = Math.round(revenueJpy);
  row.getCell(COL.手数料).value = fee;
  row.getCell(COL.仕入原価).value = cost === null ? "" : cost;
  row.getCell(COL.送料).value = shipping === null ? "" : shipping;
  row.getCell(COL.梱包費).value = packing === null ? "" : packing;
  row.getCell(COL.最終利益).value = profit;
  row.getCell(COL.利益率).value = margin;
  row.commit();
  await atomicWriteWorkbook(wb, LEDGER_PATH);

  const warnings = [];
  if (!managedResult.managed) {
    const newPid = requestedPid;
    const newQty = "数量" in body ? (numOrNull(body["数量"]) === null ? 1 : numOrNull(body["数量"])) : oldQty;
    const stockAdjustments = [];
    if (newPid !== oldPid) {
      if (oldPid) stockAdjustments.push({ pid: oldPid, delta: oldQty, reason: "注文編集(商品ID変更・戻し)" });
      if (newPid) stockAdjustments.push({ pid: newPid, delta: -newQty, reason: "注文編集(商品ID変更)" });
    } else if (newQty !== oldQty && oldPid) stockAdjustments.push({ pid: oldPid, delta: oldQty - newQty, reason: "注文編集(数量変更)" });
    for (const adj of stockAdjustments) {
      const result = await realInv.adjustRealStock({
        loadInventoryWorkbook: loadInventoryWorkbookLocked, INVENTORY_PATH, HISTORY_PATH, INV_COL,
        pid: adj.pid, delta: adj.delta, reason: adj.reason, orderNo,
      });
      if (result.warning) warnings.push(result.warning);
    }
  }

  sendJson(res, 200, {
    status: managedResult.retry ? "already_applied" : "ok", 注文番号: orderNo, 最終利益円: profit, 利益率: margin,
    managed: managedResult.managed, inventory: managedResult.managed ? managedResult : undefined,
    warning: warnings.length ? warnings.join(" / ") : null,
  });
}

async function handleDeleteOrders(req, res) {
  let body;
  try {
    body = JSON.parse((await readRawBody(req, 1024 * 1024)).toString("utf8"));
  } catch (e) {
    return sendJson(res, 400, { error: "リクエストの内容を読み取れませんでした" });
  }
  const targets = new Set((body["注文番号"] || []).map(String));
  if (!targets.size) return sendJson(res, 400, { error: "削除する注文番号を指定してください" });

  const wb = await loadWorkbookLocked();
  // sales lockは呼び出し元が保持している。必ず sales → inventory の順で取得し、
  // inventory側を先に冪等更新することでsales保存失敗後のretryでも二重復元しない。
  const managedResult = await deleteManagedOrdersTransaction({
    withInventoryLock, loadInventoryWorkbookLocked, atomicWriteWorkbook, INVENTORY_PATH,
    orderNumbers: targets, INV_HEADERS,
  });
  const managedTargets = new Set(managedResult.managedOrderNumbers);
  let deleted = 0;
  const restores = [];
  for (const ws of dataSheets(wb)) {
    const toRemove = [];
    for (let r = 2; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      if (!targets.has(String(row.getCell(COL.注文番号).value || ""))) continue;
      toRemove.push(r);
      const targetOrderNo = String(row.getCell(COL.注文番号).value || "");
      if (!managedTargets.has(targetOrderNo)) {
        const pid = String(row.getCell(COL.商品ID).value || "");
        const qty = numOrNull(row.getCell(COL.数量).value) === null ? 1 : numOrNull(row.getCell(COL.数量).value);
        if (pid) restores.push({ pid, qty, orderNo: targetOrderNo });
      }
    }
    for (let i = toRemove.length - 1; i >= 0; i--) {
      ws.spliceRows(toRemove[i], 1);
      deleted++;
    }
  }
  if (!deleted) {
    if (managedTargets.size) {
      const status = managedResult.alreadyDeleted.length === managedTargets.size ? "already_deleted" : "recovered_deleted";
      return sendJson(res, 200, { status, deleted: 0, managed: managedResult });
    }
    return sendJson(res, 404, { error: "該当する注文が見つかりません" });
  }
  await atomicWriteWorkbook(wb, LEDGER_PATH);

  const warnings = [];
  for (const r of restores) {
    const result = await realInv.adjustRealStock({
      loadInventoryWorkbook: loadInventoryWorkbookLocked, INVENTORY_PATH, HISTORY_PATH, INV_COL,
      pid: r.pid, delta: r.qty, reason: "注文削除", orderNo: r.orderNo,
    });
    if (result.warning) warnings.push(result.warning);
  }

  sendJson(res, 200, { status: "ok", deleted, managed: managedResult, warning: warnings.length ? warnings.join(" / ") : null });
}

async function handleDeleteInventory(req, res) {
  let body;
  try {
    body = JSON.parse((await readRawBody(req, 1024 * 1024)).toString("utf8"));
  } catch (e) {
    return sendJson(res, 400, { error: "リクエストの内容を読み取れませんでした" });
  }
  const targets = new Set((body["商品ID"] || []).map(String));
  if (!targets.size) return sendJson(res, 400, { error: "削除する商品IDを指定してください" });

  const deletedCount = await withInventoryLock(async () => {
    const wb = await loadInventoryWorkbookLocked();
    const ws = wb.getWorksheet("在庫管理表") || wb.worksheets[0];
    const toRemove = [];
    for (let r = 2; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      if (targets.has(String(row.getCell(INV_COL.商品ID).value || ""))) toRemove.push(r);
    }
    for (let i = toRemove.length - 1; i >= 0; i--) ws.spliceRows(toRemove[i], 1);
    if (!toRemove.length) return 0;
    await atomicWriteWorkbook(wb, INVENTORY_PATH);
    return toRemove.length;
  });
  if (!deletedCount) return sendJson(res, 404, { error: "該当する商品IDが見つかりません" });
  sendJson(res, 200, { status: "ok", deleted: deletedCount });
}

function isOrderRow(row) {
  return Boolean(row.getCell(COL.注文番号).value) && typeof row.getCell(COL.収益USD).value === "number";
}

async function handleListOrders(req, res) {
  const wb = await loadWorkbook();
  const rows = [];
  for (const ws of dataSheets(wb)) {
    ws.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      if (!isOrderRow(row)) return;
      rows.push(row.values.slice(1, HEADERS.length + 1));
    });
  }
  sendJson(res, 200, { headers: HEADERS, rows });
}

async function handleDownload(req, res) {
  await loadWorkbook();
  res.writeHead(200, {
    "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "Content-Disposition": contentDispositionAttachment("売上管理表.xlsx"),
    "Access-Control-Allow-Origin": "*",
  });
  fs.createReadStream(LEDGER_PATH).pipe(res);
}

async function handleImport(req, res) {
  const buf = await readRawBody(req, 30 * 1024 * 1024);
  const wb = new ExcelJS.Workbook();
  try {
    await wb.xlsx.load(buf);
  } catch (e) {
    return sendJson(res, 400, { error: "有効なxlsxファイルではありません" });
  }
  await withSalesLock(async () => {
    // 検証済みbufferだけを、一意tmp + renameの共通atomic writerで置換する。
    atomicWriteBuffer(buf, LEDGER_PATH);
  });
  sendJson(res, 200, { status: "ok", message: "取り込みが完了しました" });
}

async function handleListInventory(req, res) {
  const wb = await loadInventoryWorkbook();
  const ws = wb.getWorksheet("在庫管理表") || wb.worksheets[0];
  const rows = [];
  ws.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    rows.push(row.values.slice(1, INV_HEADERS.length + 1));
  });
  sendJson(res, 200, { headers: INV_HEADERS, rows });
}

async function handleSummary(req, res) {
  const wb = await loadWorkbook();
  const orderRows = [];
  for (const ws of dataSheets(wb)) {
    ws.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      if (!isOrderRow(row)) return;
      orderRows.push(row.values.slice(1, HEADERS.length + 1));
    });
  }

  const invWb = await loadInventoryWorkbook();
  const invWs = invWb.getWorksheet("在庫管理表") || invWb.worksheets[0];
  const invRows = [];
  invWs.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    invRows.push(row.values.slice(1, INV_HEADERS.length + 1));
  });

  let totalRevenue = 0, totalCost = 0, totalProfit = 0, totalQty = 0;
  orderRows.forEach((r) => {
    totalRevenue += Number(r[8]) || 0;
    totalCost += Number(r[10]) || 0;
    totalProfit += Number(r[13]) || 0;
    totalQty += Number(r[4]) || 0;
  });

  const monthlyMap = new Map();
  orderRows.forEach((r) => {
    const ym = String(r[1] || "").slice(0, 7);
    if (!ym) return;
    if (!monthlyMap.has(ym)) monthlyMap.set(ym, { month: ym, count: 0, qty: 0, revenue: 0, cost: 0, profit: 0 });
    const m = monthlyMap.get(ym);
    m.count += 1;
    m.qty += Number(r[4]) || 0;
    m.revenue += Number(r[8]) || 0;
    m.cost += Number(r[10]) || 0;
    m.profit += Number(r[13]) || 0;
  });
  const monthly = Array.from(monthlyMap.values())
    .sort((a, b) => a.month.localeCompare(b.month))
    .map((m) => ({
      month: m.month, count: m.count, qty: m.qty,
      revenue: Math.round(m.revenue), cost: Math.round(m.cost), profit: Math.round(m.profit),
      margin: m.revenue ? Math.round((m.profit / m.revenue) * 1000) / 1000 : null,
    }));

  const siteMap = new Map();
  orderRows.forEach((r) => {
    const site = r[2] || "不明";
    siteMap.set(site, (siteMap.get(site) || 0) + (Number(r[8]) || 0));
  });
  const bySite = Array.from(siteMap.entries())
    .map(([site, revenue]) => ({ site, revenue: Math.round(revenue) }))
    .sort((a, b) => b.revenue - a.revenue);

  let inventoryValue = 0, pricedCount = 0;
  invRows.forEach((r) => {
    const price = Number(r[15]);
    const qty = Number(r[3]) || 0;
    if (price > 0) {
      inventoryValue += price * qty;
      pricedCount += 1;
    }
  });

  const recentOrders = orderRows
    .slice()
    .sort((a, b) => String(b[1] || "").localeCompare(String(a[1] || "")))
    .slice(0, 30)
    .map((r) => ({
      注文番号: r[0], 日付: r[1], サイト: r[2], 商品メモ: r[3], 数量: r[4],
      収益円: r[8], 最終利益円: r[13], 利益率: r[14],
    }));

  sendJson(res, 200, {
    generatedAt: new Date().toISOString(),
    summary: {
      totalOrders: orderRows.length,
      totalQty,
      totalRevenue: Math.round(totalRevenue),
      totalCost: Math.round(totalCost),
      totalProfit: Math.round(totalProfit),
      totalMargin: totalRevenue ? Math.round((totalProfit / totalRevenue) * 1000) / 1000 : null,
    },
    monthly,
    bySite,
    byCategory: null,
    inventory: {
      totalValue: Math.round(inventoryValue),
      pricedProductCount: pricedCount,
      totalProductCount: invRows.length,
    },
    recentOrders,
    notes: [
      "byCategory は商品カテゴリ分類のデータが現状ないため未対応です(null)。",
      "bySite は eBay出品先サイト(US/UK/AU等)ごとの集計です。Shopee連携実装後はそちらのサイトもここに含まれます。",
    ],
  });
}

async function handlePatchInventory(req, res) {
  let body;
  try {
    body = JSON.parse((await readRawBody(req, 1024 * 1024)).toString("utf8"));
  } catch (e) {
    return sendJson(res, 400, { error: "リクエストの内容を読み取れませんでした" });
  }
  const pid = body["商品ID"];
  if (!pid) return sendJson(res, 400, { error: "商品ID は必須です" });

  // ロック取得前に弾けるバリデーションは先に済ませる(不正な値でロックを無駄に占有しない)。
  let stagedQty; // undefined = 対象外, null = クリア, それ以外 = 0以上の整数
  if ("棚卸入力数量" in body) {
    const v = realInv.validateStocktakeQty(body["棚卸入力数量"]);
    if (!v.ok) return sendJson(res, 400, { error: v.error });
    stagedQty = v.value;
  }
  let realStockValue; // undefined = 対象外
  if ("リアル在庫" in body) {
    realStockValue = numOrNull(body["リアル在庫"]);
    if (realStockValue === null) return sendJson(res, 400, { error: "リアル在庫 は数値で指定してください" });
  }

  // 在庫管理表.xlsxへの読み込み→変更→書き込みは1リクエストにつき1回だけ行う
  // (以前は汎用フィールド・リアル在庫・棚卸入力数量でそれぞれ個別に読み書きしており、
  //  Excel全体への書き込みが最大3回発生していた。同時更新の競合リスクを減らすため統合する)。
  const outcome = await withInventoryLock(async () => {
    const wb = await loadInventoryWorkbookLocked();
    const ws = wb.getWorksheet("在庫管理表") || wb.worksheets[0];
    const row = realInv.findInventoryRow(ws, INV_COL, pid);
    if (!row) return { found: false };

    if ("仕入価格円" in body) {
      const v = numOrNull(body["仕入価格円"]);
      row.getCell(INV_COL.仕入価格).value = v === null ? "" : v;
    }
    if ("仕入日" in body) row.getCell(INV_COL.仕入日).value = body["仕入日"];
    if ("仕入先" in body) row.getCell(INV_COL.仕入先).value = body["仕入先"];
    if ("備考" in body) row.getCell(INV_COL.備考).value = body["備考"];
    if ("日本語商品名" in body) row.getCell(INV_COL.日本語商品名).value = body["日本語商品名"];

    // リアル在庫・棚卸入力数量は当社独自管理の値のため、専用の変更ロジック(realInventory.js)を
    // 同じトランザクション内で(読み込み・書き込みを増やさずに)適用する。
    let realStockChange = null;
    if (realStockValue !== undefined) {
      realStockChange = realInv.applyRealStockToRow(row, INV_COL, realStockValue, body["リアル在庫確認日"] || null);
    }
    if (stagedQty !== undefined) {
      realInv.applyStagedQtyToRow(row, INV_COL, stagedQty);
    }

    row.commit();
    await atomicWriteWorkbook(wb, INVENTORY_PATH);
    return { found: true, name: row.getCell(INV_COL.商品名).value, realStockChange };
  });

  if (!outcome.found) return sendJson(res, 404, { error: "該当する商品IDが見つかりません" });

  if (outcome.realStockChange) {
    await appendHistory(HISTORY_PATH, {
      pid, name: outcome.name, before: outcome.realStockChange.before, after: outcome.realStockChange.after, reason: "手動変更",
    });
  }

  sendJson(res, 200, { status: "ok", 商品ID: pid });
}

async function handleImportInventory(req, res) {
  const buf = await readRawBody(req, 30 * 1024 * 1024);
  const uploadedWorkbook = new ExcelJS.Workbook();
  try {
    await uploadedWorkbook.xlsx.load(buf);
  } catch (e) {
    return sendJson(res, 400, { error: "有効なxlsxファイルではありません" });
  }
  const requiredHeaders = ["商品ID", "商品名", "US_出品ID", "UK_出品ID", "AU_出品ID"];
  let uploadValidation;
  try {
    uploadValidation = validateInventorySheet(uploadedWorkbook, requiredHeaders);
  } catch (e) {
    return sendJson(res, 400, { error: e.message });
  }

  const result = await withInventoryLock(async () => {
    const serverWorkbook = await loadInventoryWorkbookLocked();
    // 両方未作成の旧形式は許可するが、一方だけ存在する/壊れている状態は取り込みで隠さない。
    validateProtectedSheets(serverWorkbook, { allowMissing: true });
    const merged = buildProtectedImportWorkbook({ uploadedWorkbook, serverWorkbook, requiredInventoryHeaders: requiredHeaders });

    fs.mkdirSync(INVENTORY_IMPORT_BACKUP_DIR, { recursive: true });
    let backupPath = null;
    if (fs.existsSync(INVENTORY_PATH)) {
      const stamp = new Date().toISOString().replace(/[:.]/g, "-");
      backupPath = path.join(INVENTORY_IMPORT_BACKUP_DIR, `在庫管理表.${stamp}.${process.pid}-${crypto.randomBytes(4).toString("hex")}.bak.xlsx`);
      fs.copyFileSync(INVENTORY_PATH, backupPath);
    }
    await atomicWriteWorkbook(merged.workbook, INVENTORY_PATH);
    return { backupPath, productCount: uploadValidation.productCount };
  });
  sendJson(res, 200, { status: "ok", message: "取り込みが完了しました", productCount: result.productCount, backupCreated: Boolean(result.backupPath) });
}

async function handleRebuildInventory(req, res) {
  const buf = await readRawBody(req, 30 * 1024 * 1024);
  const text = buf.toString("utf8").replace(/^﻿/, "");
  const records = csvToObjects(text);
  if (!records.length) return sendJson(res, 400, { error: "CSVにデータがありません" });
  for (const col of ["Title", "Listing site", "Item number"]) {
    if (!(col in records[0])) {
      return sendJson(res, 400, { error: `CSVに ${col} 列が見つかりません。eBayの『All active listings report』形式か確認してください` });
    }
  }

  const today = new Date().toISOString().slice(0, 10);
  const summary = await rebuildInventoryFromRecords({
    records,
    sourceNote: `・eBayの『All active listings report』(${today}時点)を元に、サーバー上で自動更新しました。`,
    removedNoteLabel: "eBay出品CSV",
    INV_HEADERS,
    INVENTORY_PATH,
    loadInventoryWorkbook: loadInventoryWorkbookLocked,
  });
  sendJson(res, 200, { status: "ok", ...summary });
}

async function handleDiscrepancies(req, res) {
  const wb = await loadInventoryWorkbook();
  const ws = wb.getWorksheet("在庫管理表") || wb.worksheets[0];
  const rows = realInv.computeDiscrepancies({ ws, INV_HEADERS, INV_COL });
  sendJson(res, 200, { count: rows.length, rows });
}

// eBay同期で「どのUS出品にも紐付かなかった」UK/AU出品の一覧(在庫管理表には反映されない未紐付け分)。
// 無理な統合はせず、そのまま「要確認」として提示するためのエンドポイント。CSV取込時はこのシートが無いため空配列を返す。
async function handleUnlinkedInventory(req, res) {
  const wb = await loadInventoryWorkbook();
  const ws = wb.getWorksheet("UK_AU保管(US未紐付け)");
  if (!ws) return sendJson(res, 200, { count: 0, rows: [] });
  const headers = ["サイト", "出品ID", "商品名", "在庫数", "累計売却数", "価格", "通貨", "出品開始日時"];
  const rows = [];
  ws.eachRow((row, r) => {
    if (r === 1) return;
    rows.push(headers.map((_, i) => row.getCell(i + 1).value));
  });
  sendJson(res, 200, { count: rows.length, headers, rows });
}

async function handleStocktakePreview(req, res) {
  const wb = await loadInventoryWorkbook();
  const ws = wb.getWorksheet("在庫管理表") || wb.worksheets[0];
  const preview = realInv.previewStocktake({ ws, INV_HEADERS });
  sendJson(res, 200, preview);
}

async function handleStocktakeConfirm(req, res) {
  let body;
  try {
    body = JSON.parse((await readRawBody(req, 1024 * 1024)).toString("utf8"));
  } catch (e) {
    return sendJson(res, 400, { error: "リクエストの内容を読み取れませんでした" });
  }
  const result = await realInv.confirmStocktake({
    loadInventoryWorkbook: loadInventoryWorkbookLocked, INVENTORY_PATH, HISTORY_PATH, INV_HEADERS, INV_COL,
    targetPids: Array.isArray(body["商品ID"]) ? body["商品ID"] : null,
  });
  sendJson(res, 200, { status: "ok", ...result });
}

// 棚卸チェックの一括解除。リアル在庫・棚卸入力数量・日本語商品名などは一切変更しない。
async function handleStocktakeResetChecks(req, res) {
  const result = await realInv.resetStocktakeChecks({ loadInventoryWorkbook: loadInventoryWorkbookLocked, INVENTORY_PATH, INV_COL });
  sendJson(res, 200, { status: "ok", ...result });
}

async function handleInventoryHistory(req, res) {
  const url = new URL(req.url, "http://localhost");
  const pid = url.searchParams.get("商品ID") || undefined;
  const rows = await listHistory(HISTORY_PATH, { pid });
  sendJson(res, 200, { headers: require("./inventoryHistory").HISTORY_HEADERS, rows });
}

async function handleListUnresolvedOrderLines(req, res) {
  const workbook = await loadInventoryWorkbook();
  validateProtectedSheets(workbook, { allowMissing: true });
  const lines = readOrderLines(workbook)
    .filter((line) => [ORDER_LINE_STATUS.UNAPPLIED, ORDER_LINE_STATUS.REVIEW, ORDER_LINE_STATUS.CONFLICT].includes(line["適用状態"]))
    .map((line) => ({
      lineKey: line["明細キー"], orderNo: line["注文番号"], site: line["販売サイト"],
      ebayItemId: line["eBay Item ID"], title: line["商品タイトル"], quantity: line["注文数量"],
      status: line["適用状態"], actionable: line["適用状態"] === ORDER_LINE_STATUS.UNAPPLIED,
    }));
  const sheet = workbook.getWorksheet("在庫管理表") || workbook.worksheets[0];
  const products = [];
  for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber++) {
    const row = sheet.getRow(rowNumber);
    const pid = String(row.getCell(INV_COL.商品ID).value || "").trim();
    if (!pid) continue;
    products.push({
      pid,
      title: String(row.getCell(INV_COL.商品名).value || ""),
      japaneseTitle: String(row.getCell(INV_COL.日本語商品名).value || ""),
      sku: String(row.getCell(INV_COL.SKU).value || ""),
    });
  }
  sendJson(res, 200, { count: lines.length, lines, products });
}

async function handleResolveOrderLine(req, res) {
  let body;
  try {
    body = JSON.parse((await readRawBody(req, 1024 * 1024)).toString("utf8"));
  } catch (e) {
    return sendJson(res, 400, { error: "リクエストの内容を読み取れませんでした" });
  }
  if (body.confirmed !== true) return sendJson(res, 400, { error: "確認操作が完了していません" });
  const result = await resolveOrderLineTransaction({
    withInventoryLock, loadInventoryWorkbookLocked, atomicWriteWorkbook, INVENTORY_PATH,
    lineKey: body.lineKey, pid: body.pid, INV_HEADERS,
  });
  sendJson(res, 200, { status: result.alreadyApplied ? "already_applied" : "applied", ...result });
}

async function handleClosingChecklist(req, res) {
  const url = new URL(req.url, "http://localhost");
  const asOf = url.searchParams.get("asOf") || new Date().toISOString().slice(0, 10);
  const wb = await loadInventoryWorkbook();
  const ws = wb.getWorksheet("在庫管理表") || wb.worksheets[0];
  const rows = realInv.closingChecklist({ ws, INV_HEADERS, asOf });
  sendJson(res, 200, { asOf, count: rows.length, rows });
}

async function handleClosingExport(req, res) {
  const url = new URL(req.url, "http://localhost");
  const asOf = url.searchParams.get("asOf") || new Date().toISOString().slice(0, 10);
  const wb = await loadInventoryWorkbook();
  const ws = wb.getWorksheet("在庫管理表") || wb.worksheets[0];
  const { workbook } = await realInv.exportClosingXlsx({ ws, INV_HEADERS, asOf });
  res.writeHead(200, {
    "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "Content-Disposition": contentDispositionAttachment(`棚卸資産_${asOf}.xlsx`),
    "Access-Control-Allow-Origin": "*",
  });
  await workbook.xlsx.write(res);
  res.end();
}

async function handleClosingSnapshot(req, res) {
  let body;
  try {
    body = JSON.parse((await readRawBody(req, 1024 * 1024)).toString("utf8"));
  } catch (e) {
    return sendJson(res, 400, { error: "リクエストの内容を読み取れませんでした" });
  }
  const asOf = body["asOf"] || new Date().toISOString().slice(0, 10);
  const wb = await loadInventoryWorkbook();
  const ws = wb.getWorksheet("在庫管理表") || wb.worksheets[0];
  const { workbook, totalValue, rowCount } = await realInv.exportClosingXlsx({ ws, INV_HEADERS, asOf });
  fs.mkdirSync(SNAPSHOT_DIR, { recursive: true });
  const fileName = `棚卸資産_${asOf}.xlsx`;
  const snapshotPath = path.join(SNAPSHOT_DIR, fileName);
  if (fs.existsSync(snapshotPath)) {
    return sendJson(res, 409, { error: `基準日「${asOf}」のスナップショットは既に保存されています(上書きはできません。別の基準日を指定するか、既存ファイルを確認してください)` });
  }
  const tmpPath = snapshotPath + ".tmp";
  await workbook.xlsx.writeFile(tmpPath);
  fs.renameSync(tmpPath, snapshotPath);
  sendJson(res, 200, { status: "ok", asOf, fileName, totalValue, rowCount });
}

async function handleListSnapshots(req, res) {
  fs.mkdirSync(SNAPSHOT_DIR, { recursive: true });
  const files = fs.readdirSync(SNAPSHOT_DIR).filter((f) => f.endsWith(".xlsx")).sort().reverse();
  sendJson(res, 200, { files });
}

async function handleDownloadSnapshot(req, res, pathname) {
  const fileName = decodeURIComponent(pathname.replace("/download/snapshots/", ""));
  if (fileName.includes("..") || fileName.includes("/") || fileName.includes("\\")) {
    return sendJson(res, 400, { error: "不正なファイル名です" });
  }
  const filePath = path.join(SNAPSHOT_DIR, fileName);
  if (!fs.existsSync(filePath)) return sendJson(res, 404, { error: "ファイルが見つかりません" });
  res.writeHead(200, {
    "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "Content-Disposition": contentDispositionAttachment(fileName),
    "Access-Control-Allow-Origin": "*",
  });
  fs.createReadStream(filePath).pipe(res);
}

const DASHBOARD_PAGE = `<!doctype html>
<html lang="ja"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Agate Trade Hub</title>
<style>
  :root {
    color-scheme: light;
    --page:        #f9f9f7;
    --surface:     #fcfcfb;
    --surface-2:   #f3f2ee;
    --ink:         #0b0b0b;
    --ink-2:       #52514e;
    --ink-muted:   #898781;
    --border:      rgba(11,11,11,0.10);
    --grid:        #e1e0d9;
    --baseline:    #c3c2b7;
    --good:        #006300;
    --series-rev:  #2a78d6;
    --series-cost: #eb6834;
    --series-prof: #1baf7a;
    --tooltip-bg:  #0b0b0b;
    --tooltip-ink: #ffffff;
    --accent-wash: rgba(42,120,214,0.10);
  }
  @media (prefers-color-scheme: dark) {
    :root:not([data-theme="light"]) {
      color-scheme: dark;
      --page: #0d0d0d; --surface: #1a1a19; --surface-2: #232322;
      --ink: #ffffff; --ink-2: #c3c2b7; --ink-muted: #898781;
      --border: rgba(255,255,255,0.10); --grid: #2c2c2a; --baseline: #383835; --good: #0ca30c;
      --series-rev: #3987e5; --series-cost: #d95926; --series-prof: #199e70;
      --tooltip-bg: #ffffff; --tooltip-ink: #0b0b0b;
      --accent-wash: rgba(57,135,229,0.14);
    }
  }
  * { box-sizing: border-box; }
  body {
    margin: 0; background: var(--page); color: var(--ink);
    font-family: -apple-system, BlinkMacSystemFont, "Hiragino Sans", "Hiragino Kaku Gothic ProN", "Yu Gothic UI", "Segoe UI", Roboto, sans-serif;
    -webkit-font-smoothing: antialiased;
  }
  .wrap { max-width: 1900px; margin: 0 auto; padding: 32px 32px 60px; display: flex; flex-direction: column; gap: 20px; }
  .hdr { display: flex; justify-content: space-between; align-items: flex-end; flex-wrap: wrap; gap: 12px; border-bottom: 1px solid var(--border); padding-bottom: 18px; }
  .hdr h1 { margin: 0 0 6px; font-size: 24px; font-weight: 700; letter-spacing: -0.01em; }
  .hdr .sub { margin: 0; color: var(--ink-2); font-size: 13px; }

  .auth-row { display: flex; gap: 8px; align-items: center; flex-wrap: wrap; }
  .auth-row input[type=password] { padding: 8px 10px; border-radius: 7px; border: 1px solid var(--border); background: var(--surface-2); color: var(--ink); font-size: 13px; min-width: 240px; }
  .btn { font: inherit; font-size: 12.5px; font-weight: 600; color: var(--ink); background: var(--surface-2); border: 1px solid var(--border); border-radius: 7px; padding: 7px 14px; cursor: pointer; }
  .btn:hover { background: var(--accent-wash); }
  .btn:disabled { opacity: 0.45; cursor: default; }
  .btn:disabled:hover { background: var(--surface-2); }
  th.sortable { cursor: pointer; user-select: none; }
  th.sortable:hover { color: var(--ink-2); }
  th.sortable .sort-arrow { margin-left: 3px; opacity: 0.6; }
  td.checkbox-col, th.checkbox-col { width: 30px; padding-left: 10px; padding-right: 0; }
  #status { font-size: 12.5px; color: var(--ink-muted); }

  .tabnav { display: flex; gap: 4px; border-bottom: 1px solid var(--border); }
  .tabbtn { font: inherit; font-size: 14px; font-weight: 600; color: var(--ink-muted); background: none; border: none; border-bottom: 2px solid transparent; padding: 10px 16px; cursor: pointer; }
  .tabbtn:hover { color: var(--ink-2); }
  .tabbtn.active { color: var(--series-rev); border-bottom-color: var(--series-rev); }
  .tabpanel { display: none; flex-direction: column; gap: 20px; }
  .tabpanel.active { display: flex; }

  .kpis { display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 1px; background: var(--border); border: 1px solid var(--border); border-radius: 12px; overflow: hidden; }
  .kpi { background: var(--surface); padding: 16px 18px; display: flex; flex-direction: column; gap: 6px; }
  .kpi .label { font-size: 12px; color: var(--ink-2); }
  .kpi .value { font-size: 22px; font-weight: 600; letter-spacing: -0.01em; }
  .kpi .value.good { color: var(--good); }
  .kpi .value.bad { color: var(--series-cost); }

  .panel { background: var(--surface); border: 1px solid var(--border); border-radius: 12px; padding: 22px 24px 20px; }
  .panel h2 { margin: 0; font-size: 15px; font-weight: 700; }
  .panel .desc { margin: 3px 0 0; font-size: 12.5px; color: var(--ink-muted); }
  .panel-body { margin-top: 16px; }

  .browser-toolbar { display: flex; align-items: center; gap: 12px; flex-wrap: wrap; }
  .search-input { font: inherit; font-size: 13px; color: var(--ink); background: var(--surface-2); border: 1px solid var(--border); border-radius: 8px; padding: 9px 12px; flex: 1; min-width: 220px; }
  .search-input:focus { outline: 2px solid var(--series-rev); outline-offset: 1px; background: var(--surface); }
  .result-count { font-size: 12.5px; color: var(--ink-muted); white-space: nowrap; }

  .table-scroll { overflow-x: auto; min-width: 0; max-height: 620px; overflow-y: auto; margin-top: 14px; padding-right: 24px; }
  table { width: 100%; border-collapse: collapse; font-size: 12.5px; }
  thead th { text-align: left; font-size: 11.5px; color: var(--ink-muted); font-weight: 600; padding: 8px 10px; border-bottom: 1px solid var(--border); white-space: nowrap; position: sticky; top: 0; background: var(--surface); z-index: 2; }
  thead th.num, td.num { text-align: right; font-variant-numeric: tabular-nums; }
  tbody td { padding: 8px 10px; border-bottom: 1px solid var(--border); white-space: nowrap; }
  tbody tr:hover { background: var(--surface-2); }
  .sticky-col { position: sticky; left: 0; background: var(--surface); z-index: 1; box-shadow: 2px 0 4px -2px var(--border); }
  tbody tr:hover td.sticky-col { background: var(--surface-2); }
  .site-chip { display: inline-block; padding: 2px 8px; border-radius: 100px; background: var(--accent-wash); color: var(--series-rev); font-size: 11.5px; font-weight: 600; }
  .item-link { color: var(--series-rev); text-decoration: underline; }
  .item-link:hover { text-decoration: none; }
  td.truncate { max-width: 220px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  td.stk-en-name { max-width: 160px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  textarea.ja-name-input { font: inherit; font-size: 13px; color: var(--ink); background: var(--surface); border: 1px solid var(--border); border-radius: 7px; padding: 6px 8px; width: 260px; min-height: 40px; max-height: 78px; resize: vertical; line-height: 1.35; white-space: normal; overflow-wrap: break-word; }
  textarea.ja-name-input:focus { outline: 2px solid var(--series-rev); outline-offset: 1px; background: var(--surface); }
  .profit-cell { color: var(--good); font-weight: 600; }
  .profit-cell.bad { color: var(--series-cost); }
  .stock-zero { color: var(--series-cost); font-weight: 600; }
  .row-unconfirmed { background: rgba(255,196,0,0.12); }
  .mismatch-cell { color: var(--series-cost); font-weight: 700; }
  .row-abnormal { background: rgba(235,104,52,0.14); }
  .stk-check-badge { display: none; align-items: center; justify-content: center; width: 16px; height: 16px; margin-left: 6px; border-radius: 50%; background: rgba(27,175,122,0.16); color: var(--good); font-size: 10px; font-weight: 700; vertical-align: middle; }
  .pager { display: flex; align-items: center; gap: 6px; margin: 10px 0; flex-wrap: wrap; }
  .pager button { font: inherit; font-size: 12.5px; color: var(--ink); background: var(--surface-2); border: 1px solid var(--border); border-radius: 6px; padding: 5px 10px; cursor: pointer; min-width: 32px; }
  .pager button:hover:not(:disabled) { background: var(--accent-wash); }
  .pager button:disabled { opacity: 0.4; cursor: default; }
  .pager button.active { background: var(--series-rev); color: #fff; border-color: var(--series-rev); font-weight: 700; }
  .pager .pager-info { font-size: 12px; color: var(--ink-muted); margin-left: 6px; }
  .thumb-btn { display: inline-flex; align-items: center; justify-content: center; width: 40px; height: 40px; border-radius: 6px; border: 1px solid var(--border); background: var(--surface-2); cursor: pointer; padding: 0; overflow: hidden; }
  .thumb-btn img { width: 100%; height: 100%; object-fit: cover; }
  .thumb-btn.no-image { font-size: 9px; color: var(--ink-muted); text-align: center; line-height: 1.2; }
  .img-modal-overlay { display: none; position: fixed; inset: 0; background: rgba(0,0,0,0.6); z-index: 100; align-items: center; justify-content: center; }
  .img-modal-overlay.show { display: flex; }
  .img-modal { background: var(--surface); border-radius: 10px; padding: 16px; max-width: 90vw; max-height: 90vh; display: flex; flex-direction: column; gap: 10px; align-items: center; }
  .img-modal img { max-width: 80vw; max-height: 70vh; object-fit: contain; border-radius: 6px; }
  .img-modal-caption { font-size: 13px; color: var(--ink-2); text-align: center; }
  .img-modal-close { align-self: flex-end; }

  .row2 { display: grid; grid-template-columns: 1.3fr 1fr; gap: 20px; }
  @media (max-width: 860px) { .row2 { grid-template-columns: 1fr; } }
  .legend { display: flex; gap: 16px; flex-wrap: wrap; margin-bottom: 8px; }
  .legend-item { display: flex; align-items: center; gap: 6px; font-size: 12.5px; color: var(--ink-2); }
  .legend-swatch { width: 10px; height: 10px; border-radius: 2px; flex: none; }
  svg text { font-family: inherit; fill: var(--ink-muted); }
  .grid-line { stroke: var(--grid); stroke-width: 1; }
  .baseline-line { stroke: var(--baseline); stroke-width: 1; }
  .bar-hit { cursor: pointer; }
  .val-label { font-size: 11px; fill: var(--ink-2); font-variant-numeric: tabular-nums; }
  .axis-label { font-size: 11px; fill: var(--ink-muted); }
  .viz-tooltip {
    position: fixed; pointer-events: none; background: var(--tooltip-bg); color: var(--tooltip-ink);
    font-size: 12px; line-height: 1.5; padding: 8px 10px; border-radius: 6px;
    box-shadow: 0 4px 14px rgba(0,0,0,0.25); opacity: 0; transform: translate(-50%, -100%);
    transition: opacity 0.1s ease; z-index: 50; white-space: nowrap;
  }
  .viz-tooltip.show { opacity: 1; }
  .viz-tooltip b { font-weight: 700; }

  .scroll-mirror-top { overflow-x: scroll; overflow-y: hidden; height: 14px; margin-top: 14px; scrollbar-color: var(--baseline) var(--surface-2); scrollbar-width: auto; }
  .scroll-mirror-top-inner { height: 1px; }
  .scroll-mirror-top::-webkit-scrollbar { height: 12px; }
  .scroll-mirror-top::-webkit-scrollbar-track { background: var(--surface-2); border-radius: 6px; }
  .scroll-mirror-top::-webkit-scrollbar-thumb { background: var(--baseline); border-radius: 6px; }
  .scroll-mirror-top::-webkit-scrollbar-thumb:hover { background: var(--ink-muted); }
  .table-scroll { margin-top: 0 !important; }

  td input { font: inherit; font-variant-numeric: tabular-nums; font-size: 12.5px; color: var(--ink); background: var(--surface-2); border: 1px solid var(--border); border-radius: 5px; padding: 5px 6px; width: 78px; text-align: right; }
  td input:focus { outline: 2px solid var(--series-rev); outline-offset: 1px; background: var(--surface); }
  td input.wide-input { width: 160px; text-align: left; }
  td input.wide-input.wider { width: 280px; }
  /* 数量入力欄の誤操作防止: 上下スピナーを非表示にする(WebKit系・Firefox両対応)。
     価格・為替等の金額入力には適用しない(qty-inputクラスを付けた要素のみ対象)。 */
  input.qty-input::-webkit-outer-spin-button,
  input.qty-input::-webkit-inner-spin-button { -webkit-appearance: none; margin: 0; }
  input.qty-input { -moz-appearance: textfield; appearance: textfield; }
  tr.saving td { background: #fff7e0 !important; }
  tr.saved td { background: #e9f7ec !important; }
  tr.error td { background: #fde8e8 !important; }

  .paste-parse { display: flex; flex-direction: column; gap: 10px; padding: 14px; margin-bottom: 16px; background: var(--surface-2); border: 1px dashed var(--border); border-radius: 10px; }
  .paste-parse textarea { font: inherit; font-size: 12.5px; color: var(--ink); background: var(--surface); border: 1px solid var(--border); border-radius: 7px; padding: 10px; width: 100%; resize: vertical; line-height: 1.5; }
  .paste-parse textarea:focus { outline: 2px solid var(--series-rev); outline-offset: 1px; }
  .entry-form { display: grid; grid-template-columns: repeat(3, minmax(120px, 1fr)); gap: 12px 14px; }
  .entry-form label { display: flex; flex-direction: column; gap: 4px; font-size: 12px; color: var(--ink-2); }
  .entry-form input { font: inherit; font-size: 13px; color: var(--ink); background: var(--surface); border: 1px solid var(--border); border-radius: 7px; padding: 7px 9px; }
  .entry-form input:focus { outline: 2px solid var(--series-rev); outline-offset: 1px; }
  .hint { font-size: 12px; color: var(--ink-muted); }
  .hint.ok { color: var(--good); }
  .hint.ng { color: var(--series-cost); }
  .order-parse-summary { display: flex; flex-wrap: wrap; gap: 8px 18px; padding: 10px 12px; border: 1px solid var(--border); border-radius: 7px; background: var(--surface); font-size: 12.5px; }
  .order-parse-summary.ok { border-color: var(--good); }
  .order-parse-summary.ng { border-color: var(--series-cost); background: #fde8e8; }
  .order-items-review { max-height: 360px; overflow: auto; border: 1px solid var(--border); border-radius: 7px; background: var(--surface); }
  .order-items-review table { min-width: 850px; }
  .order-item-status-ok { color: var(--good); font-weight: 700; }
  .order-item-status-ng { color: var(--series-cost); font-weight: 700; }
  input[type=file] { font-size: 12.5px; color: var(--ink-2); }
</style></head>
<body>
<div class="wrap">
  <div class="hdr">
    <div>
      <h1>Agate Trade Hub</h1>
      <p class="sub">株式会社アゲイト — eBay 注文・在庫・売上。注文タブは書き換えると即座にサーバーに保存されます</p>
    </div>
    <div class="auth-row">
      <input id="token" type="password" placeholder="アクセストークン">
      <button class="btn" id="toggleToken" type="button">表示</button>
      <button class="btn" id="saveToken">トークンを記憶</button>
      <span id="status"></span>
    </div>
  </div>

  <div class="tabnav">
    <button class="tabbtn active" data-tab="sales">売上分析</button>
    <button class="tabbtn" data-tab="inventory">在庫</button>
    <button class="tabbtn" data-tab="orders">注文</button>
    <button class="tabbtn" data-tab="discrepancy">相違</button>
    <button class="tabbtn" data-tab="stocktake">棚卸</button>
    <button class="tabbtn" data-tab="closing">決算</button>
  </div>

  <div class="tabpanel active" id="tab-sales">
    <div class="kpis" id="kpis"></div>

    <div class="panel">
      <h2>月次データ</h2>
      <p class="desc">サーバー上の最新データから自動集計(注文の日付ごとに月単位でまとめています)</p>
      <div class="panel-body table-scroll">
        <table>
          <thead><tr><th>月</th><th class="num">件数</th><th class="num">総個数</th><th class="num">受取額(円)</th><th class="num">仕入回収額</th><th class="num">純利益(円)</th><th class="num">利益率</th></tr></thead>
          <tbody id="monthly-body"></tbody>
        </table>
      </div>
    </div>

    <div class="row2">
      <div class="panel">
        <h2>月次推移(受取額・仕入回収・純利益)</h2>
        <div class="panel-body">
          <div class="legend">
            <span class="legend-item"><span class="legend-swatch" style="background:var(--series-rev)"></span>受取額</span>
            <span class="legend-item"><span class="legend-swatch" style="background:var(--series-cost)"></span>仕入回収額</span>
            <span class="legend-item"><span class="legend-swatch" style="background:var(--series-prof)"></span>純利益</span>
          </div>
          <svg id="chart-monthly" viewBox="0 0 780 300" width="100%"></svg>
        </div>
      </div>
      <div class="panel">
        <h2>月次利益率</h2>
        <div class="panel-body">
          <svg id="chart-margin" viewBox="0 0 480 300" width="100%"></svg>
        </div>
      </div>
    </div>

    <div class="panel">
      <h2>サイト別 受取額</h2>
      <p class="desc">出品先サイトごとの受取額合計(円)</p>
      <div class="panel-body">
        <svg id="chart-site" viewBox="0 0 780 220" width="100%"></svg>
      </div>
    </div>

    <div class="panel">
      <h2>利益TOP10の注文</h2>
      <p class="desc">最終利益が大きかった注文(サーバー上の最新データ)</p>
      <div class="panel-body table-scroll">
        <table>
          <thead><tr><th>日付</th><th>サイト</th><th>商品メモ</th><th class="num">総個数</th><th class="num">収益(円)</th><th class="num">最終利益(円)</th><th class="num">利益率</th></tr></thead>
          <tbody id="top10-body"></tbody>
        </table>
      </div>
    </div>
  </div>

  <div class="tabpanel" id="tab-inventory">
    <div class="panel">
      <h2>在庫一覧(閲覧専用)</h2>
      <p class="desc">商品名・商品IDで検索できます。eBayの「All active listings report」CSVをアップロードすると、名寄せ・フラグ付けまで自動でやり直します。</p>
      <div class="panel-body">
        <div class="kpis" id="inv-kpis" style="margin-bottom:16px;"></div>
        <div class="paste-parse">
          <div class="browser-toolbar">
            <button class="btn" id="ebay-rebuild-btn">eBay最新情報を取り込む</button>
            <span id="ebay-rebuild-status" class="hint"></span>
          </div>
          <div class="browser-toolbar">
            <input type="file" id="inv-csv-file" accept=".csv">
            <button class="btn" id="inv-csv-upload">CSVで在庫を更新</button>
            <span id="inv-csv-status" class="hint"></span>
          </div>
        </div>
        <div class="browser-toolbar">
          <input type="text" class="search-input" id="inv-q" placeholder="商品名・商品IDで検索…">
          <button class="btn" id="inv-delete-btn" disabled>選択した行を削除(<span id="inv-selected-count">0</span>)</button>
          <button class="btn" id="inv-csv-export">CSVでダウンロード</button>
          <span class="result-count" id="inv-count"></span>
        </div>
        <div class="pager" id="inv-pager-top"></div>
        <div class="scroll-mirror-top" id="inv-table-mirror"><div class="scroll-mirror-top-inner"></div></div>
        <div class="table-scroll" id="inv-table-scroll">
          <table>
            <thead><tr id="inv-thead"></tr></thead>
            <tbody id="inv-tbody"></tbody>
          </table>
        </div>
        <div class="pager" id="inv-pager-bottom"></div>
      </div>
    </div>
  </div>

  <div class="tabpanel" id="tab-orders">
    <div class="panel">
      <h2>新しい注文を登録</h2>
      <p class="desc">eBayの注文詳細ページの内容をそのまま貼り付けてください。読み取れる範囲を自動で埋めます。内容を確認・修正してから登録してください(空欄は自分で埋めてください)。</p>
      <div class="panel-body">
        <div class="paste-parse">
          <textarea id="ne-paste" rows="6" placeholder="ここに貼り付け"></textarea>
          <div class="browser-toolbar">
            <button class="btn" id="ne-parse-btn">読み取る</button>
          </div>
          <div class="entry-form" id="ne-review" style="display:none;">
            <label>注文番号<input id="ne-order" type="text"></label>
            <label>日付<input id="ne-date" type="date"></label>
            <label>サイト<input id="ne-site" type="text"></label>
            <label>収益USD<input id="ne-usd" type="number" step="any"></label>
            <label>ドル円レート<input id="ne-rate" type="number" step="any" value="155"></label>
            <label>手数料(円)<input id="ne-fee-preview" type="text" disabled placeholder="自動計算(収益円の3%)"></label>
            <label>仕入原価(円)<input id="ne-cost" type="number" step="any" placeholder="わかれば入力"></label>
            <label>送料(円)<input id="ne-shipping" type="number" step="any" placeholder="わかれば入力"></label>
            <label>梱包費(円)<input id="ne-packing" type="number" step="any" placeholder="わかれば入力" value="50"></label>
            <label>数量<input id="ne-qty" type="number" step="1" min="0" class="qty-input"></label>
            <label style="grid-column:span 2;">商品メモ<input id="ne-note" type="text"></label>
            <label style="grid-column:span 3;">商品ID(リアル在庫と連動させたい場合は指定。商品名で検索できます)
              <input id="ne-pid" type="text" list="ne-pid-list" placeholder="例: P0001、または商品名で検索">
              <datalist id="ne-pid-list"></datalist>
            </label>
          </div>
          <div id="ne-items-review" style="display:none;">
            <div id="ne-parse-summary" class="order-parse-summary"></div>
            <div class="order-items-review" style="margin-top:10px;">
              <table>
                <thead><tr><th>No.</th><th>商品タイトル</th><th>eBay Item ID</th><th class="num">数量</th><th>SKU</th><th>解析状態</th></tr></thead>
                <tbody id="ne-items-body"></tbody>
              </table>
            </div>
          </div>
          <div class="browser-toolbar" id="ne-submit-row" style="display:none;">
            <button class="btn" id="ne-submit-btn">この内容で登録する</button>
            <span id="ne-status" class="hint"></span>
          </div>
        </div>
      </div>
    </div>

    <div class="panel">
      <h2>未紐付け明細</h2>
      <p class="desc">未適用のeBay Item IDを、確認したAgate商品へ手動で紐付けます。矛盾・要確認はこの画面では変更できません。</p>
      <div class="panel-body">
        <div class="browser-toolbar">
          <button class="btn" id="unresolved-refresh">再読み込み</button>
          <span id="unresolved-status" class="hint"></span>
        </div>
        <datalist id="unresolved-products"></datalist>
        <div class="table-scroll">
          <table>
            <thead><tr><th>注文番号</th><th>サイト</th><th>eBay Item ID</th><th>商品タイトル</th><th class="num">数量</th><th>状態</th><th>確認したPxxxx</th><th>操作</th></tr></thead>
            <tbody id="unresolved-body"></tbody>
          </table>
        </div>
      </div>
    </div>

    <div class="panel">
      <h2>注文一覧</h2>
      <p class="desc">注文番号・商品メモ・サイトで検索できます。<b>収益USD・ドル円レート・仕入原価・送料・梱包費は直接書き換えられます</b>(最終利益はその場で再計算・保存されます)。</p>
      <div class="panel-body">
        <div class="paste-parse">
          <div class="browser-toolbar">
            <input type="file" id="ord-xlsx-file" accept=".xlsx">
            <button class="btn" id="ord-xlsx-upload">売上管理表(xlsx)を丸ごと取り込む</button>
            <span id="ord-xlsx-status" class="hint"></span>
          </div>
          <p class="hint">サーバー上の注文データがこのファイルの内容で完全に置き換わります。手元で作り直した売上管理表を反映したいときに使ってください。</p>
        </div>
        <div class="browser-toolbar">
          <input type="text" class="search-input" id="ord-q" placeholder="注文番号・商品メモ・サイトで検索…">
          <button class="btn" id="ord-delete-btn" disabled>選択した行を削除(<span id="ord-selected-count">0</span>)</button>
          <button class="btn" id="ord-csv-export">CSVでダウンロード</button>
          <span class="result-count" id="ord-count"></span>
        </div>
        <datalist id="ord-pid-list"></datalist>
        <div class="scroll-mirror-top" id="ord-table-mirror"><div class="scroll-mirror-top-inner"></div></div>
        <div class="table-scroll" id="ord-table-scroll">
          <table>
            <thead><tr id="ord-thead"></tr></thead>
            <tbody id="ord-tbody"></tbody>
          </table>
        </div>
      </div>
    </div>
  </div>

  <div class="tabpanel" id="tab-discrepancy">
    <div class="panel">
      <h2>相違一覧</h2>
      <p class="desc">2種類の比較のうち<b>どちらか一方でもズレていれば</b>表示します: ①「リアル在庫」(当社独自管理の実在庫)と「US在庫」の比較、②「US在庫」を基準にした「UK/AU在庫」との比較(eBaymagの国間同期が壊れていないかのチェック)。出品されている国だけを比較します。まだ実地棚卸で確認していない行(リアル在庫確認日が空欄)は黄色で表示します。</p>
      <div class="panel-body">
        <div class="browser-toolbar">
          <button class="btn" id="disc-refresh-btn">再読み込み</button>
          <span class="result-count" id="disc-count"></span>
        </div>
        <div class="table-scroll">
          <table>
            <thead><tr><th>商品ID</th><th>商品名</th><th class="num">リアル在庫</th><th class="num">US在庫</th><th class="num">UK在庫</th><th class="num">AU在庫</th><th>リアル在庫確認日</th></tr></thead>
            <tbody id="disc-tbody"></tbody>
          </table>
        </div>
      </div>
    </div>

    <div class="panel">
      <h2>未紐付け(要確認)</h2>
      <p class="desc">UK/AUの出品のうち、タイトルの一致でどのUS出品にも紐付けられなかったものです。<b>同一商品と確実に判断できないため、自動では統合していません。</b>実際には既存商品と同一の可能性があるので、手動でご確認ください。</p>
      <div class="panel-body">
        <div class="browser-toolbar">
          <span class="result-count" id="unlinked-count"></span>
        </div>
        <div class="table-scroll">
          <table>
            <thead><tr><th>サイト</th><th>出品ID</th><th>商品名</th><th class="num">在庫数</th><th class="num">価格</th><th>出品開始日時</th></tr></thead>
            <tbody id="unlinked-tbody"></tbody>
          </table>
        </div>
      </div>
    </div>
  </div>

  <div class="tabpanel" id="tab-stocktake">
    <div class="panel">
      <h2>棚卸入力</h2>
      <p class="desc">「棚卸入力数量」に実際に数えた個数を入力すると、その場で一時保存されます(この時点ではリアル在庫は変わりません)。入力が終わったら下の「一括確定」で内容を確認してからリアル在庫へ反映してください。日本語商品名はその場で編集・保存できます。「画像」をクリックすると商品画像を拡大表示します。</p>
      <div class="panel-body">
        <div class="browser-toolbar">
          <span class="result-count" id="stk-progress"></span>
          <button class="btn" id="stk-reset-checks-btn">棚卸チェックをすべて解除</button>
          <span id="stk-reset-status" class="hint"></span>
        </div>
        <div class="browser-toolbar">
          <input type="text" class="search-input" id="stk-q" placeholder="英語商品名・日本語商品名・商品ID・SKUで検索…">
          <label class="hint">並び替え
            <select id="stk-sort">
              <option value="pid">商品ID順</option>
              <option value="name">英語商品名順</option>
              <option value="ja-name">日本語商品名順</option>
              <option value="price">仕入価格が安い順</option>
              <option value="real">リアル在庫が少ない順</option>
              <option value="unchecked">未棚卸を先頭</option>
            </select>
          </label>
          <span class="result-count" id="stk-count"></span>
        </div>
        <div class="pager" id="stk-pager-top"></div>
        <div class="table-scroll">
          <table>
            <thead><tr id="stk-thead"></tr></thead>
            <tbody id="stk-tbody"></tbody>
          </table>
        </div>
        <div class="pager" id="stk-pager-bottom"></div>
      </div>
    </div>

    <div class="panel">
      <h2>棚卸の一括確定</h2>
      <p class="desc">「棚卸入力数量」を入力した商品だけが対象です。内容を確認してから確定してください(確定するとリアル在庫が書き換わります)。</p>
      <div class="panel-body">
        <div class="browser-toolbar">
          <button class="btn" id="stk-preview-btn">確定前に内容を確認する</button>
          <span id="stk-preview-status" class="hint"></span>
        </div>
        <div id="stk-preview-summary" class="kpis" style="display:none; margin-bottom:16px;"></div>
        <div class="table-scroll" id="stk-preview-table-wrap" style="display:none;">
          <table>
            <thead><tr><th>商品ID</th><th>商品名</th><th class="num">リアル在庫(現在)</th><th class="num">棚卸入力数量</th><th class="num">差異</th><th>異常値</th></tr></thead>
            <tbody id="stk-preview-tbody"></tbody>
          </table>
        </div>
        <div class="browser-toolbar" id="stk-confirm-row" style="display:none;">
          <button class="btn" id="stk-confirm-btn">この内容で一括確定する</button>
          <span id="stk-confirm-status" class="hint"></span>
        </div>
      </div>
    </div>
  </div>

  <div class="tabpanel" id="tab-closing">
    <div class="panel">
      <h2>決算棚卸資産</h2>
      <p class="desc">「リアル在庫」を基準に棚卸資産(商品ID・商品名・リアル在庫・仕入価格・評価額)をまとめます。決算日(基準日)を指定してください。</p>
      <div class="panel-body">
        <div class="browser-toolbar">
          <label class="hint">基準日 <input type="date" id="cls-asof"></label>
          <button class="btn" id="cls-checklist-btn">未確認商品をチェック</button>
          <button class="btn" id="cls-export-btn">棚卸資産をエクスポート(xlsx)</button>
          <button class="btn" id="cls-snapshot-btn">この基準日でスナップショット保存</button>
          <span id="cls-status" class="hint"></span>
        </div>
        <p class="hint">「未確認商品をチェック」は、基準日より前に実地棚卸(リアル在庫確認日)が済んでいない商品を一覧化します。「スナップショット保存」は、後から見返しても数字が変わらない確定記録として基準日ごとに1回だけ保存できます。</p>
        <div class="table-scroll" id="cls-checklist-wrap" style="display:none;">
          <table>
            <thead><tr><th>商品ID</th><th>商品名</th><th class="num">リアル在庫</th><th>リアル在庫確認日</th></tr></thead>
            <tbody id="cls-checklist-tbody"></tbody>
          </table>
        </div>
      </div>
    </div>

    <div class="panel">
      <h2>保存済みスナップショット</h2>
      <div class="panel-body">
        <ul id="cls-snapshot-list" class="hint"></ul>
      </div>
    </div>
  </div>
</div>
<div class="viz-tooltip" id="tooltip"></div>
<div class="img-modal-overlay" id="img-modal-overlay">
  <div class="img-modal">
    <button class="btn img-modal-close" id="img-modal-close">閉じる ×</button>
    <img id="img-modal-img" src="" alt="商品画像">
    <div class="img-modal-caption" id="img-modal-caption"></div>
  </div>
</div>
<script src="/order-parser.js"></script>
<script>
const ORD_HEADERS = ["注文番号","日付","サイト","商品メモ","数量","商品ID","収益USD","ドル円レート","収益円","手数料(円)","仕入原価(円)","送料(円)","梱包費(円)","最終利益(円)","利益率"];
const ORD_EDITABLE = ["収益USD","ドル円レート","数量","仕入原価(円)","送料(円)","梱包費(円)"];
const ORD_EDITABLE_TEXT = ["商品メモ", "商品ID"];
const ORD_NUM_COLS = ["収益USD","ドル円レート","数量","収益円","手数料(円)","仕入原価(円)","送料(円)","梱包費(円)","最終利益(円)","利益率"];
const ORD_FIELD_KEY = { "収益USD": "収益USD", "ドル円レート": "ドル円レート", "仕入原価(円)": "仕入原価円", "送料(円)": "送料円", "梱包費(円)": "梱包費円", "商品メモ": "商品メモ", "数量": "数量", "商品ID": "商品ID" };
const INV_NUM_COLS = ["在庫数(現物)","出品国数","US価格(USD)","US累計売却数","UK価格(GBP)","UK累計売却数","AU価格(AUD)","AU累計売却数","仕入価格(円)"];
const INV_EDITABLE_TEXT = ["仕入日","備考"];
const INV_EDITABLE_NUM = ["仕入価格(円)"];
const INV_FIELD_KEY = { "仕入価格(円)": "仕入価格円", "仕入日": "仕入日", "仕入先": "仕入先", "備考": "備考" };
let orderRows = [];
let invRows = [];
let invHeaders = [];
let unresolvedProducts = [];

function getToken() { return localStorage.getItem("agate_token") || ""; }
(() => {
  // URLに ?token=... が付いていれば自動的に記憶する(手入力不要でブラウザ確認できるようにするため)
  const urlToken = new URLSearchParams(window.location.search).get("token");
  if (urlToken) localStorage.setItem("agate_token", urlToken);
})();
document.getElementById("token").value = getToken();
document.getElementById("saveToken").addEventListener("click", () => {
  localStorage.setItem("agate_token", document.getElementById("token").value.trim());
  loadAll();
});

document.querySelectorAll(".tabbtn").forEach(btn => {
  btn.addEventListener("click", () => {
    document.querySelectorAll(".tabbtn").forEach(b => b.classList.remove("active"));
    document.querySelectorAll(".tabpanel").forEach(p => p.classList.remove("active"));
    btn.classList.add("active");
    document.getElementById("tab-" + btn.dataset.tab).classList.add("active");
    if (btn.dataset.tab === "orders") setupScrollMirror("ord-table-mirror", "ord-table-scroll");
    if (btn.dataset.tab === "inventory") setupScrollMirror("inv-table-mirror", "inv-table-scroll");
    if (btn.dataset.tab === "discrepancy") loadDiscrepancies();
    if (btn.dataset.tab === "stocktake") loadStocktakeList();
    if (btn.dataset.tab === "closing") loadSnapshotList();
  });
});

function fmt(v, isPercent) {
  if (v === null || v === undefined || v === "") return "";
  if (isPercent) return (Number(v) * 100).toFixed(1) + "%";
  if (typeof v === "number") return v.toLocaleString("ja-JP", { maximumFractionDigits: 2 });
  return v;
}

async function loadAll() {
  const token = getToken();
  const statusEl = document.getElementById("status");
  if (!token) { statusEl.textContent = "トークンを入力してください"; return; }
  statusEl.textContent = "読み込み中...";
  try {
    const [ordRes, invRes, unresolvedRes] = await Promise.all([
      fetch("/api/orders", { headers: { Authorization: "Bearer " + token } }),
      fetch("/api/inventory", { headers: { Authorization: "Bearer " + token } }),
      fetch("/api/order-lines/unresolved", { headers: { Authorization: "Bearer " + token } }),
    ]);
    if (!ordRes.ok) { statusEl.textContent = "エラー: " + ordRes.status + "(トークンを確認してください)"; return; }
    const ordData = await ordRes.json();
    orderRows = ordData.rows;
    if (invRes.ok) {
      const invData = await invRes.json();
      invRows = invData.rows;
      invHeaders = invData.headers;
    }
    if (unresolvedRes.ok) renderUnresolvedOrderLines(await unresolvedRes.json());
    renderKpis();
    renderSalesTab();
    renderOrders();
    renderInventory();
    statusEl.textContent = "";
  } catch (e) {
    statusEl.textContent = "通信エラー: " + e.message;
  }
}

function renderUnresolvedOrderLines(data) {
  const body = document.getElementById("unresolved-body");
  const productList = document.getElementById("unresolved-products");
  const statusEl = document.getElementById("unresolved-status");
  unresolvedProducts = Array.isArray(data.products) ? data.products : [];
  productList.replaceChildren();
  unresolvedProducts.forEach((product) => {
    const option = document.createElement("option");
    option.value = product.pid + " | " + product.title + " | " + product.japaneseTitle + " | " + product.sku;
    option.label = product.pid + " / " + product.title + (product.japaneseTitle ? " / " + product.japaneseTitle : "");
    productList.appendChild(option);
  });
  body.replaceChildren();
  (data.lines || []).forEach((line) => {
    const tr = document.createElement("tr");
    [line.orderNo, line.site, line.ebayItemId, line.title, line.quantity, line.status].forEach((value, index) => {
      const td = document.createElement("td");
      td.textContent = value === null || value === undefined ? "" : value;
      if (index === 4) td.className = "num";
      tr.appendChild(td);
    });
    const selectTd = document.createElement("td");
    const input = document.createElement("input");
    input.type = "text";
    input.setAttribute("list", "unresolved-products");
    input.placeholder = "Pxxxx・商品名・SKUで検索";
    input.disabled = !line.actionable;
    selectTd.appendChild(input);
    tr.appendChild(selectTd);
    const actionTd = document.createElement("td");
    const button = document.createElement("button");
    button.className = "btn";
    button.textContent = line.actionable ? "確認して適用" : "要確認";
    button.disabled = !line.actionable;
    button.addEventListener("click", async () => {
      const pidMatch = /^\s*(P\d+)\b/.exec(input.value);
      if (!pidMatch) { statusEl.textContent = "候補からPxxxxを選択してください"; statusEl.className = "hint ng"; return; }
      const pid = pidMatch[1];
      const product = unresolvedProducts.find((candidate) => candidate.pid === pid);
      if (!product) { statusEl.textContent = "選択したPxxxxが候補にありません"; statusEl.className = "hint ng"; return; }
      const message = line.site + "\\nItem ID: " + line.ebayItemId + "\\n商品: " + line.title + "\\n数量: " + line.quantity
        + "\\n\\n適用先: " + product.pid + "\\n英語商品名: " + product.title + "\\n日本語商品名: " + product.japaneseTitle
        + "\\n\\nこの対応でリアル在庫を" + line.quantity + "減らし、今後このItem IDを" + product.pid + "として使用します。よろしいですか？";
      if (!window.confirm(message)) return;
      statusEl.textContent = "適用中...";
      statusEl.className = "hint";
      try {
        const response = await fetch("/api/order-lines/resolve", {
          method: "POST",
          headers: { "Content-Type": "application/json", Authorization: "Bearer " + getToken() },
          body: JSON.stringify({ lineKey: line.lineKey, pid, confirmed: true }),
        });
        const result = await response.json();
        if (!response.ok) { statusEl.textContent = "失敗: " + (result.error || response.status); statusEl.className = "hint ng"; return; }
        await loadAll();
        statusEl.textContent = result.status === "already_applied"
          ? "この明細はすでに適用済みです"
          : pid + "へ数量" + result.quantity + "を適用しました（リアル在庫 " + result.before + " → " + result.after + "）";
        statusEl.className = "hint ok";
      } catch (error) {
        statusEl.textContent = "通信エラー: " + error.message;
        statusEl.className = "hint ng";
      }
    });
    actionTd.appendChild(button);
    tr.appendChild(actionTd);
    body.appendChild(tr);
  });
  statusEl.textContent = "未処理・要確認 " + (data.count || 0) + "件";
  statusEl.className = "hint";
}

document.getElementById("unresolved-refresh").addEventListener("click", loadAll);

function renderKpis() {
  let revenue = 0, cost = 0, profit = 0, profitKnown = 0, qty = 0;
  orderRows.forEach(r => {
    revenue += Number(r[8]) || 0;
    cost += Number(r[10]) || 0;
    qty += Number(r[4]) || 0;
    if (r[13] !== "" && r[13] !== null && r[13] !== undefined) { profit += Number(r[13]); profitKnown++; }
  });
  const margin = revenue !== 0 ? profit / revenue : 0;
  const tiles = [
    { label: "受取総額", value: "¥" + fmt(Math.round(revenue)) },
    { label: "仕入回収額", value: "¥" + fmt(Math.round(cost)) },
    { label: "純利益", value: "¥" + fmt(Math.round(profit)), cls: profit >= 0 ? "good" : "bad" },
    { label: "利益率", value: (margin * 100).toFixed(1) + "%", cls: margin >= 0 ? "good" : "bad" },
    { label: "総注文数", value: orderRows.length.toLocaleString("ja-JP") + " 件" },
    { label: "総販売個数", value: qty.toLocaleString("ja-JP") + " 個" },
    { label: "在庫評価額(仕入ベース)", value: "¥" + fmt(Math.round(computeInventoryValue().total)) },
    { label: "総在庫個数", value: computeTotalStock().toLocaleString("ja-JP") + " 個" },
  ];
  document.getElementById("kpis").innerHTML = tiles.map(t =>
    '<div class="kpi"><div class="label">' + t.label + '</div><div class="value' + (t.cls ? " " + t.cls : "") + '">' + t.value + '</div></div>'
  ).join("");

  const top10 = orderRows.filter(r => r[13] !== "" && r[13] !== null && r[13] !== undefined)
    .slice().sort((a, b) => Number(b[13]) - Number(a[13])).slice(0, 10);
  const top10Body = document.getElementById("top10-body");
  top10Body.replaceChildren();
  top10.forEach((r) => {
    const tr = document.createElement("tr");
    const dateTd = document.createElement("td");
    dateTd.textContent = r[1] || "";
    tr.appendChild(dateTd);

    const siteTd = document.createElement("td");
    const siteChip = document.createElement("span");
    siteChip.className = "site-chip";
    siteChip.textContent = r[2] || "不明";
    siteTd.appendChild(siteChip);
    tr.appendChild(siteTd);

    const noteTd = document.createElement("td");
    noteTd.className = "truncate";
    noteTd.textContent = r[3] || "";
    noteTd.title = String(r[3] || "");
    tr.appendChild(noteTd);

    const values = [[r[4], "num", false], [r[8], "num", false], [r[13], "num profit-cell" + (Number(r[13]) < 0 ? " bad" : ""), false], [r[14], "num", true]];
    values.forEach(([value, className, percent]) => {
      const td = document.createElement("td");
      td.className = className;
      td.textContent = fmt(value, percent);
      tr.appendChild(td);
    });
    top10Body.appendChild(tr);
  });
}

function computeMonthly() {
  const map = new Map();
  orderRows.forEach((r) => {
    const d = String(r[1] || "");
    if (d.length < 7) return;
    const month = d.slice(0, 7);
    if (!map.has(month)) map.set(month, { revenue: 0, cost: 0, profit: 0, count: 0, qty: 0 });
    const m = map.get(month);
    m.revenue += Number(r[8]) || 0;
    m.cost += Number(r[10]) || 0;
    m.qty += Number(r[4]) || 0;
    if (r[13] !== "" && r[13] !== null && r[13] !== undefined) m.profit += Number(r[13]);
    m.count++;
  });
  return Array.from(map.keys()).sort().map((k) => {
    const v = map.get(k);
    return { month: k, revenue: v.revenue, cost: v.cost, profit: v.profit, count: v.count, qty: v.qty, margin: v.revenue ? v.profit / v.revenue : 0 };
  });
}

function renderMonthlyTable(monthly) {
  const newestFirst = monthly.slice().reverse();
  document.getElementById("monthly-body").innerHTML = newestFirst.map((m) =>
    "<tr><td>" + m.month + "</td><td class='num'>" + m.count + "</td><td class='num'>" + m.qty + "</td><td class='num'>" + fmt(Math.round(m.revenue)) +
    "</td><td class='num'>" + fmt(Math.round(m.cost)) + "</td><td class='num profit-cell" + (m.profit < 0 ? " bad" : "") + "'>" + fmt(Math.round(m.profit)) +
    "</td><td class='num'>" + (m.margin * 100).toFixed(1) + "%</td></tr>"
  ).join("");
}

function svgNode(tag, attrs) {
  const el = document.createElementNS("http://www.w3.org/2000/svg", tag);
  for (const k in attrs) el.setAttribute(k, attrs[k]);
  return el;
}

function showTooltip(evt, html) {
  const tip = document.getElementById("tooltip");
  tip.innerHTML = html;
  tip.style.left = evt.clientX + "px";
  tip.style.top = (evt.clientY - 10) + "px";
  tip.classList.add("show");
}
function hideTooltip() {
  document.getElementById("tooltip").classList.remove("show");
}

function renderMonthlyChart(monthly) {
  const svg = document.getElementById("chart-monthly");
  svg.innerHTML = "";
  if (!monthly.length) return;
  const W = 780, H = 300, padL = 55, padR = 10, padT = 10, padB = 30;
  const innerW = W - padL - padR, innerH = H - padT - padB;
  const maxV = Math.max(1, ...monthly.map((m) => Math.max(m.revenue, m.cost, Math.abs(m.profit))));
  const n = monthly.length;
  const bw = innerW / n;
  const colors = { revenue: "var(--series-rev)", cost: "var(--series-cost)", profit: "var(--series-prof)" };
  const labels = { revenue: "受取額", cost: "仕入回収額", profit: "純利益" };

  for (let i = 0; i <= 4; i++) {
    const y = padT + innerH - (innerH * i) / 4;
    svg.appendChild(svgNode("line", { x1: padL, y1: y, x2: padL + innerW, y2: y, class: "grid-line" }));
    const lab = svgNode("text", { x: padL - 8, y: y + 4, "text-anchor": "end", class: "val-label" });
    lab.textContent = fmt(Math.round((maxV * i) / 4));
    svg.appendChild(lab);
  }
  svg.appendChild(svgNode("line", { x1: padL, y1: padT + innerH, x2: padL + innerW, y2: padT + innerH, class: "baseline-line" }));

  monthly.forEach((m, i) => {
    const x0 = padL + i * bw;
    const groupW = bw * 0.78;
    const barW = groupW / 3;
    ["revenue", "cost", "profit"].forEach((k, j) => {
      const v = m[k];
      const h = Math.max(0, (Math.abs(v) / maxV) * innerH);
      const y = padT + innerH - h;
      const x = x0 + bw * 0.11 + j * barW;
      const rect = svgNode("rect", { x: x, y: y, width: Math.max(1, barW * 0.85), height: h, fill: colors[k], rx: 1.5, class: "bar-hit" });
      rect.addEventListener("mousemove", (evt) => showTooltip(evt, "<b>" + m.month + "</b><br>" + labels[k] + ": ¥" + fmt(Math.round(v))));
      rect.addEventListener("mouseleave", hideTooltip);
      svg.appendChild(rect);
    });
    const label = svgNode("text", { x: x0 + bw / 2, y: padT + innerH + 16, "text-anchor": "middle", class: "axis-label" });
    label.textContent = m.month.slice(2).replace("-", "/");
    svg.appendChild(label);
  });
}

function renderMarginChart(monthly) {
  const svg = document.getElementById("chart-margin");
  svg.innerHTML = "";
  if (!monthly.length) return;
  const W = 480, H = 300, padL = 50, padR = 10, padT = 10, padB = 30;
  const innerW = W - padL - padR, innerH = H - padT - padB;
  const dataMin = Math.min(...monthly.map((m) => m.margin));
  const dataMax = Math.max(...monthly.map((m) => m.margin));
  const minV = Math.min(0, dataMin);
  const maxV = Math.max(0.01, dataMax);
  const range = (maxV - minV) || 1;
  const n = monthly.length;
  const stepX = n > 1 ? innerW / (n - 1) : 0;
  const yFor = (v) => padT + innerH - ((v - minV) / range) * innerH;

  for (let i = 0; i <= 4; i++) {
    const v = minV + (range * i) / 4;
    const y = yFor(v);
    svg.appendChild(svgNode("line", { x1: padL, y1: y, x2: padL + innerW, y2: y, class: "grid-line" }));
    const lab = svgNode("text", { x: padL - 8, y: y + 4, "text-anchor": "end", class: "val-label" });
    lab.textContent = (v * 100).toFixed(0) + "%";
    svg.appendChild(lab);
  }
  svg.appendChild(svgNode("line", { x1: padL, y1: yFor(0), x2: padL + innerW, y2: yFor(0), class: "baseline-line" }));

  let d = "";
  monthly.forEach((m, i) => {
    const x = padL + i * stepX;
    const y = yFor(m.margin);
    d += (i === 0 ? "M" : "L") + x + "," + y + " ";
  });
  svg.appendChild(svgNode("path", { d: d.trim(), fill: "none", stroke: "var(--series-rev)", "stroke-width": 2 }));

  const everyN = Math.ceil(n / 8) || 1;
  monthly.forEach((m, i) => {
    const x = padL + i * stepX;
    const y = yFor(m.margin);
    const c = svgNode("circle", { cx: x, cy: y, r: 3, fill: "var(--series-rev)", class: "bar-hit" });
    c.addEventListener("mousemove", (evt) => showTooltip(evt, "<b>" + m.month + "</b><br>利益率: " + (m.margin * 100).toFixed(1) + "%"));
    c.addEventListener("mouseleave", hideTooltip);
    svg.appendChild(c);
    if (i % everyN === 0) {
      const label = svgNode("text", { x: x, y: padT + innerH + 16, "text-anchor": "middle", class: "axis-label" });
      label.textContent = m.month.slice(2).replace("-", "/");
      svg.appendChild(label);
    }
  });
}

function renderSiteChart() {
  const svg = document.getElementById("chart-site");
  svg.innerHTML = "";
  const bySite = new Map();
  orderRows.forEach((r) => {
    const site = r[2] || "不明";
    bySite.set(site, (bySite.get(site) || 0) + (Number(r[8]) || 0));
  });
  const entries = Array.from(bySite.entries()).sort((a, b) => b[1] - a[1]);
  if (!entries.length) return;
  const W = 780, H = Math.max(120, entries.length * 30), padL = 70, padR = 70, padT = 10, padB = 10;
  const innerW = W - padL - padR, innerH = H - padT - padB;
  const maxV = Math.max(1, ...entries.map((e) => e[1]));
  const bh = innerH / entries.length;
  svg.setAttribute("viewBox", "0 0 " + W + " " + H);
  entries.forEach(([site, val], i) => {
    const y = padT + i * bh + bh * 0.15;
    const w = (val / maxV) * innerW;
    const rect = svgNode("rect", { x: padL, y: y, width: Math.max(1, w), height: bh * 0.7, fill: "var(--series-rev)", rx: 2, class: "bar-hit" });
    rect.addEventListener("mousemove", (evt) => showTooltip(evt, "<b>" + site + "</b><br>受取額: ¥" + fmt(Math.round(val))));
    rect.addEventListener("mouseleave", hideTooltip);
    svg.appendChild(rect);
    const label = svgNode("text", { x: padL - 8, y: y + bh * 0.35 + 4, "text-anchor": "end", class: "axis-label" });
    label.textContent = site;
    svg.appendChild(label);
    const valLabel = svgNode("text", { x: padL + w + 6, y: y + bh * 0.35 + 4, class: "val-label" });
    valLabel.textContent = "¥" + fmt(Math.round(val));
    svg.appendChild(valLabel);
  });
}

function renderSalesTab() {
  const monthly = computeMonthly();
  renderMonthlyTable(monthly);
  renderMonthlyChart(monthly);
  renderMarginChart(monthly);
  renderSiteChart();
}

function setupScrollMirror(mirrorId, scrollId) {
  const mirror = document.getElementById(mirrorId);
  const scrollEl = document.getElementById(scrollId);
  if (!mirror || !scrollEl) return;
  const inner = mirror.querySelector(".scroll-mirror-top-inner");
  const table = scrollEl.querySelector("table");
  if (!table) return;
  // table.scrollWidthではなくscrollEl.scrollWidthを使う。table-scrollの右側paddingが
  // 実スクロール範囲に含まれるため、ミラー側もそれを含めないと最大までずれてしまう。
  inner.style.width = scrollEl.scrollWidth + "px";
  let syncing = false;
  mirror.onscroll = () => { if (syncing) return; syncing = true; scrollEl.scrollLeft = mirror.scrollLeft; syncing = false; };
  scrollEl.onscroll = () => { if (syncing) return; syncing = true; mirror.scrollLeft = scrollEl.scrollLeft; syncing = false; };
}

// 数量入力欄の誤操作防止(棚卸のような連続入力でホイールによる意図しない増減を防ぐ)。
// preventDefault()は呼ばない(呼ぶとページ・表のスクロールごと止まってしまうため)。
// フォーカスを外すだけで、その回のホイールイベントによる値の増減は発生しなくなり、
// ホイール自体はそのまま素通りしてページ/表は通常どおりスクロールする。
function preventQtyWheelChange(inp) {
  inp.addEventListener("wheel", () => inp.blur(), { passive: true });
}

// 商品IDはリアル在庫連動・注文登録処理では引き続き使用するが、日常確認では
// 商品メモの方が重要なため画面表示(一覧・CSV)からのみ外す。Excel・API・
// orderRows自体からは削除しない(在庫タブのINV_HIDDENと同じ仕組み)。
const ORD_HIDDEN = ["商品ID"];
let ordSort = { idx: 1, dir: -1 };
let ordSelected = new Set();

function csvEscapeCell(v) {
  const s = v === null || v === undefined ? "" : String(v);
  if (/["\\n]/.test(s) || s.indexOf(",") !== -1) return '"' + s.replace(/"/g, '""') + '"';
  return s;
}

function downloadCsv(filename, headers, rows) {
  const lines = [headers.map(csvEscapeCell).join(",")];
  rows.forEach((r) => lines.push(r.map(csvEscapeCell).join(",")));
  const blob = new Blob(["﻿" + lines.join("\\r\\n")], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function sortRows(rows, headers, numCols, sortState) {
  const idx = sortState.idx, dir = sortState.dir;
  const isNum = numCols.includes(headers[idx]);
  return rows.slice().sort((a, b) => {
    const av = a[idx], bv = b[idx];
    if (isNum) {
      const an = av === "" || av === null || av === undefined ? -Infinity : Number(av);
      const bn = bv === "" || bv === null || bv === undefined ? -Infinity : Number(bv);
      return (an - bn) * dir;
    }
    return String(av || "").localeCompare(String(bv || "")) * dir;
  });
}

function updateOrdDeleteBtn() {
  document.getElementById("ord-selected-count").textContent = ordSelected.size;
  document.getElementById("ord-delete-btn").disabled = ordSelected.size === 0;
}

function renderOrders() {
  const displayOrder = ORD_HEADERS.map((h, i) => i).filter((i) => !ORD_HIDDEN.includes(ORD_HEADERS[i]));
  const thead = document.getElementById("ord-thead");
  const checkAllTh = '<th class="checkbox-col"><input type="checkbox" id="ord-select-all"></th>';
  thead.innerHTML = checkAllTh + displayOrder.map((i, pos) => {
    const isSortCol = i === ordSort.idx;
    const arrow = isSortCol ? '<span class="sort-arrow">' + (ordSort.dir === 1 ? "▲" : "▼") + '</span>' : "";
    return '<th class="sortable ' + (ORD_NUM_COLS.includes(ORD_HEADERS[i]) ? "num" : "") + (pos === 0 ? " sticky-col" : "") + '" data-idx="' + i + '">' + ORD_HEADERS[i] + arrow + '</th>';
  }).join("");
  thead.querySelectorAll("th.sortable").forEach((th) => {
    th.addEventListener("click", () => {
      const idx = Number(th.dataset.idx);
      ordSort = { idx, dir: ordSort.idx === idx ? -ordSort.dir : -1 };
      renderOrders();
    });
  });

  const q = document.getElementById("ord-q").value.trim().toLowerCase();
  const tbody = document.getElementById("ord-tbody");
  tbody.innerHTML = "";
  let shown = 0;
  const visible = [];
  const sortedOrders = sortRows(orderRows, ORD_HEADERS, ORD_NUM_COLS, ordSort);
  sortedOrders.forEach(row => {
    const orderNo = row[0];
    const searchable = ((row[0]||"") + " " + (row[2]||"") + " " + (row[3]||"")).toLowerCase();
    if (q && searchable.indexOf(q) === -1) return;
    shown++;
    visible.push(orderNo);
    const tr = document.createElement("tr");
    const cbTd = document.createElement("td");
    cbTd.className = "checkbox-col";
    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.checked = ordSelected.has(orderNo);
    cb.addEventListener("change", () => {
      if (cb.checked) ordSelected.add(orderNo); else ordSelected.delete(orderNo);
      updateOrdDeleteBtn();
    });
    cbTd.appendChild(cb);
    tr.appendChild(cbTd);
    displayOrder.forEach((i, pos) => {
      const h = ORD_HEADERS[i];
      const td = document.createElement("td");
      if (pos === 0) td.className = "sticky-col";
      if (ORD_EDITABLE.includes(h)) {
        td.className = (td.className ? td.className + " " : "") + "num";
        const inp = document.createElement("input");
        const isQty = h === "数量";
        inp.type = "number"; inp.step = isQty ? "1" : "any";
        if (isQty) { inp.min = "0"; inp.className = "qty-input"; preventQtyWheelChange(inp); }
        inp.value = row[i] === null || row[i] === undefined ? "" : row[i];
        inp.addEventListener("change", () => saveOrderField(tr, orderNo, h, inp.value));
        td.appendChild(inp);
      } else if (ORD_EDITABLE_TEXT.includes(h)) {
        const inp = document.createElement("input");
        inp.type = "text";
        inp.className = "wide-input";
        if (h === "商品ID") { inp.setAttribute("list", "ord-pid-list"); inp.placeholder = "商品名で検索…"; }
        inp.value = row[i] === null || row[i] === undefined ? "" : row[i];
        inp.addEventListener("change", () => saveOrderField(tr, orderNo, h, inp.value));
        td.appendChild(inp);
      } else if (h === "サイト") {
        const chip = document.createElement("span"); chip.className = "site-chip"; chip.textContent = row[i] || "不明"; td.appendChild(chip);
      } else if (h === "最終利益(円)") {
        td.className = (td.className ? td.className + " " : "") + "num profit-cell" + (Number(row[i]) < 0 ? " bad" : "");
        td.textContent = fmt(row[i]);
      } else if (h === "利益率") {
        const known = row[i] !== "" && row[i] !== null && row[i] !== undefined;
        td.className = (td.className ? td.className + " " : "") + "num profit-cell" + (known && Number(row[i]) < 0 ? " bad" : "");
        td.textContent = fmt(row[i], true);
      } else {
        td.className = (td.className ? td.className + " " : "") + (ORD_NUM_COLS.includes(h) ? "num" : "");
        td.textContent = fmt(row[i], h === "利益率");
      }
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  document.getElementById("ord-count").textContent = shown.toLocaleString("ja-JP") + " / " + orderRows.length.toLocaleString("ja-JP") + " 件";
  const selectAllCb = document.getElementById("ord-select-all");
  selectAllCb.checked = visible.length > 0 && visible.every((id) => ordSelected.has(id));
  selectAllCb.addEventListener("change", () => {
    if (selectAllCb.checked) visible.forEach((id) => ordSelected.add(id));
    else visible.forEach((id) => ordSelected.delete(id));
    renderOrders();
  });
  updateOrdDeleteBtn();
  setupScrollMirror("ord-table-mirror", "ord-table-scroll");
}

document.getElementById("ord-delete-btn").addEventListener("click", async () => {
  if (!ordSelected.size) return;
  if (!confirm(ordSelected.size + "件の注文を削除します。よろしいですか?(元に戻せません)")) return;
  const token = getToken();
  try {
    const r = await fetch("/api/orders", {
      method: "DELETE",
      headers: { "Content-Type": "application/json", Authorization: "Bearer " + token },
      body: JSON.stringify({ 注文番号: Array.from(ordSelected) }),
    });
    const data = await r.json().catch(() => ({}));
    if (!r.ok) { alert("削除に失敗しました。在庫変更は安全停止しました: " + (data.error || r.status)); return; }
    ordSelected.clear();
    await loadAll();
    if (data.warning) alert(data.warning);
  } catch (e) {
    alert("通信エラー: " + e.message);
  }
});

document.getElementById("ord-csv-export").addEventListener("click", () => {
  const displayOrder = ORD_HEADERS.map((h, i) => i).filter((i) => !ORD_HIDDEN.includes(ORD_HEADERS[i]));
  const q = document.getElementById("ord-q").value.trim().toLowerCase();
  const sortedOrders = sortRows(orderRows, ORD_HEADERS, ORD_NUM_COLS, ordSort);
  const rows = sortedOrders.filter((row) => {
    const searchable = ((row[0]||"") + " " + (row[2]||"") + " " + (row[3]||"")).toLowerCase();
    return !q || searchable.indexOf(q) !== -1;
  }).map((row) => displayOrder.map((i) => row[i]));
  downloadCsv("注文一覧_" + new Date().toISOString().slice(0, 10) + ".csv", displayOrder.map((i) => ORD_HEADERS[i]), rows);
});

async function saveOrderField(tr, orderNo, header, value) {
  tr.className = "saving";
  const token = getToken();
  const body = { 注文番号: orderNo };
  body[ORD_FIELD_KEY[header]] = ORD_EDITABLE_TEXT.includes(header) ? value : (value === "" ? "" : Number(value));
  try {
    const r = await fetch("/api/orders", {
      method: "PATCH",
      headers: { "Content-Type": "application/json", Authorization: "Bearer " + token },
      body: JSON.stringify(body),
    });
    const data = await r.json().catch(() => ({}));
    if (!r.ok) { tr.className = "error"; alert("編集できませんでした。在庫変更は安全停止しました: " + (data.error || r.status)); return; }
    tr.className = "saved";
    setTimeout(() => { tr.className = ""; }, 1500);
    if (data.warning) alert(data.warning);
  } catch (e) {
    tr.className = "error";
  }
}

const INV_HIDDEN = [
  "UK_出品ID", "UK価格(GBP)", "AU_出品ID", "AU価格(AUD)", "在庫数不一致", "バリエーション詳細", "仕入先",
  // リアル在庫機能の追加列。通常の「在庫」タブ(eBayのCSVをそのまま反映する場所)の見た目は変更しないため非表示にする。
  // 各値は「相違」「棚卸」「決算」タブや検索(SKU)から個別に利用する。
  "UK在庫数", "AU在庫数", "リアル在庫", "リアル在庫確認日", "棚卸入力数量", "棚卸入力日時", "画像URL", "日本語商品名", "SKU", "棚卸チェック",
];
let invSort = { idx: 0, dir: -1 };
let invSelected = new Set();
const INV_PAGE_SIZE = 500;
let invPage = 1;
let invPrevQ = "";

function sortInvRows(rows) {
  const idx = invSort.idx, dir = invSort.dir;
  const header = invHeaders[idx];
  if (header === "商品ID") {
    return rows.slice().sort((a, b) => {
      const na = parseInt(String(a[0] || "").slice(1), 10) || 0;
      const nb = parseInt(String(b[0] || "").slice(1), 10) || 0;
      return (na - nb) * dir;
    });
  }
  return sortRows(rows, invHeaders, INV_NUM_COLS, invSort);
}

function updateInvDeleteBtn() {
  document.getElementById("inv-selected-count").textContent = invSelected.size;
  document.getElementById("inv-delete-btn").disabled = invSelected.size === 0;
}

function computeTotalStock() {
  if (!invHeaders.length) return 0;
  const stockIdx = invHeaders.indexOf("在庫数(現物)");
  let total = 0;
  invRows.forEach((row) => { total += Number(row[stockIdx]) || 0; });
  return total;
}

function computeInventoryValue() {
  if (!invHeaders.length) return { total: 0, priced: 0 };
  const priceIdx = invHeaders.indexOf("仕入価格(円)");
  const stockIdx = invHeaders.indexOf("在庫数(現物)");
  let total = 0, priced = 0;
  invRows.forEach((row) => {
    const price = Number(row[priceIdx]);
    if (!Number.isFinite(price) || price <= 0) return;
    const stock = Number(row[stockIdx]) || 0;
    total += price * stock;
    priced++;
  });
  return { total, priced };
}

function renderInvKpis() {
  const { total, priced } = computeInventoryValue();
  document.getElementById("inv-kpis").innerHTML =
    '<div class="kpi"><div class="label">在庫評価額(仕入単価×在庫数の合計)</div><div class="value">¥' + fmt(Math.round(total)) + '</div></div>' +
    '<div class="kpi"><div class="label">仕入価格が入っている商品数</div><div class="value">' + priced.toLocaleString("ja-JP") + " / " + invRows.length.toLocaleString("ja-JP") + '</div></div>';
}

function renderPidDatalist() {
  const pidIdx = invHeaders.indexOf("商品ID"), nameIdx = invHeaders.indexOf("商品名");
  if (pidIdx === -1) return;
  const options = invRows.map((row) => '<option value="' + row[pidIdx] + '">' + (row[nameIdx] || "").replace(/"/g, "&quot;") + '</option>').join("");
  document.getElementById("ne-pid-list").innerHTML = options;
  document.getElementById("ord-pid-list").innerHTML = options;
}

function renderInventory() {
  if (!invHeaders.length) return;
  renderInvKpis();
  renderPidDatalist();
  const displayOrder = invHeaders.map((h, i) => i).filter((i) => !INV_HIDDEN.includes(invHeaders[i]));
  const thead = document.getElementById("inv-thead");
  const checkAllTh = '<th class="checkbox-col"><input type="checkbox" id="inv-select-all"></th>';
  thead.innerHTML = checkAllTh + displayOrder.map((i, pos) => {
    const isSortCol = i === invSort.idx;
    const arrow = isSortCol ? '<span class="sort-arrow">' + (invSort.dir === 1 ? "▲" : "▼") + '</span>' : "";
    return '<th class="sortable ' + (INV_NUM_COLS.includes(invHeaders[i]) ? "num" : "") + (pos === 0 ? " sticky-col" : "") + '" data-idx="' + i + '">' + invHeaders[i] + arrow + '</th>';
  }).join("");
  thead.querySelectorAll("th.sortable").forEach((th) => {
    th.addEventListener("click", () => {
      const idx = Number(th.dataset.idx);
      invSort = { idx, dir: invSort.idx === idx ? -invSort.dir : -1 };
      invPage = 1;
      renderInventory();
    });
  });

  const q = document.getElementById("inv-q").value.trim().toLowerCase();
  if (!q && invPrevQ) invPage = 1; // 検索を解除したら通常のページ表示(1ページ目)に戻す
  invPrevQ = q;

  const tbody = document.getElementById("inv-tbody");
  const pagerTop = document.getElementById("inv-pager-top");
  const pagerBottom = document.getElementById("inv-pager-bottom");
  tbody.innerHTML = "";
  const stockIdx = invHeaders.indexOf("在庫数(現物)");
  const visible = [];
  const sortedInv = sortInvRows(invRows);
  const matched = sortedInv.filter((row) => {
    const searchable = ((row[0]||"") + " " + (row[1]||"")).toLowerCase();
    return !q || searchable.indexOf(q) !== -1;
  });

  let pageRows, totalPages;
  if (q) {
    // 検索時: 全商品を対象に絞り込み、ページ分けせずヒット件数だけ表示する
    pageRows = matched;
    totalPages = 1;
    pagerTop.innerHTML = "";
    pagerBottom.innerHTML = "";
  } else {
    totalPages = Math.max(1, Math.ceil(matched.length / INV_PAGE_SIZE));
    if (invPage > totalPages) invPage = totalPages;
    if (invPage < 1) invPage = 1;
    const start = (invPage - 1) * INV_PAGE_SIZE;
    pageRows = matched.slice(start, start + INV_PAGE_SIZE);
    const goToInvPage = (p) => { invPage = p; renderInventory(); };
    renderPager(pagerTop, invPage, totalPages, goToInvPage);
    renderPager(pagerBottom, invPage, totalPages, goToInvPage);
  }

  pageRows.forEach(row => {
    const pid = row[0];
    visible.push(pid);
    const tr = document.createElement("tr");
    const cbTd = document.createElement("td");
    cbTd.className = "checkbox-col";
    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.checked = invSelected.has(pid);
    cb.addEventListener("change", () => {
      if (cb.checked) invSelected.add(pid); else invSelected.delete(pid);
      updateInvDeleteBtn();
    });
    cbTd.appendChild(cb);
    tr.appendChild(cbTd);
    displayOrder.forEach((i, pos) => {
      const h = invHeaders[i];
      const td = document.createElement("td");
      if (pos === 0) td.className = "sticky-col";
      const isNum = INV_NUM_COLS.includes(h);
      td.className = (td.className ? td.className + " " : "") + (isNum ? "num" : "");
      if (i === stockIdx && (row[i] === 0 || row[i] === null || row[i] === undefined)) td.className += " stock-zero";
      if (INV_EDITABLE_NUM.includes(h)) {
        const inp = document.createElement("input");
        inp.type = "number"; inp.step = "any";
        inp.value = row[i] === null || row[i] === undefined ? "" : row[i];
        inp.addEventListener("change", () => saveInventoryField(tr, pid, h, inp.value, row));
        td.appendChild(inp);
      } else if (INV_EDITABLE_TEXT.includes(h)) {
        const inp = document.createElement("input");
        inp.type = "text";
        inp.className = h === "備考" ? "wide-input wider" : "wide-input";
        inp.value = row[i] === null || row[i] === undefined ? "" : row[i];
        inp.addEventListener("change", () => saveInventoryField(tr, pid, h, inp.value, row));
        td.appendChild(inp);
      } else if (h === "US_出品ID" && row[i] !== null && row[i] !== undefined && row[i] !== "") {
        const a = document.createElement("a");
        a.className = "item-link";
        a.href = "https://www.ebay.com/itm/" + row[i];
        a.target = "_blank";
        a.rel = "noopener noreferrer";
        a.textContent = row[i];
        td.appendChild(a);
      } else {
        td.textContent = row[i] === null || row[i] === undefined ? "" : (isNum && typeof row[i] === "number" ? fmt(row[i]) : row[i]);
      }
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  if (q) {
    document.getElementById("inv-count").textContent = matched.length.toLocaleString("ja-JP") + " 件ヒット(全" + invRows.length.toLocaleString("ja-JP") + "件中)";
  } else {
    const start = (invPage - 1) * INV_PAGE_SIZE;
    document.getElementById("inv-count").textContent =
      (matched.length ? (start + 1) : 0).toLocaleString("ja-JP") + "〜" + Math.min(start + INV_PAGE_SIZE, matched.length).toLocaleString("ja-JP") +
      " / " + invRows.length.toLocaleString("ja-JP") + " 件(" + invPage + " / " + totalPages + " ページ)";
  }
  const selectAllCb = document.getElementById("inv-select-all");
  selectAllCb.checked = visible.length > 0 && visible.every((id) => invSelected.has(id));
  selectAllCb.addEventListener("change", () => {
    if (selectAllCb.checked) visible.forEach((id) => invSelected.add(id));
    else visible.forEach((id) => invSelected.delete(id));
    renderInventory();
  });
  updateInvDeleteBtn();
  setupScrollMirror("inv-table-mirror", "inv-table-scroll");
}

document.getElementById("inv-delete-btn").addEventListener("click", async () => {
  if (!invSelected.size) return;
  if (!confirm(invSelected.size + "件の商品を削除します。よろしいですか?(元に戻せません)")) return;
  const token = getToken();
  try {
    const r = await fetch("/api/inventory", {
      method: "DELETE",
      headers: { "Content-Type": "application/json", Authorization: "Bearer " + token },
      body: JSON.stringify({ 商品ID: Array.from(invSelected) }),
    });
    if (!r.ok) { alert("削除に失敗しました"); return; }
    invSelected.clear();
    await loadAll();
  } catch (e) {
    alert("通信エラー: " + e.message);
  }
});

document.getElementById("inv-csv-export").addEventListener("click", () => {
  if (!invHeaders.length) return;
  const displayOrder = invHeaders.map((h, i) => i).filter((i) => !INV_HIDDEN.includes(invHeaders[i]));
  const q = document.getElementById("inv-q").value.trim().toLowerCase();
  const sortedInv = sortInvRows(invRows);
  const rows = sortedInv.filter((row) => {
    const searchable = ((row[0]||"") + " " + (row[1]||"")).toLowerCase();
    return !q || searchable.indexOf(q) !== -1;
  }).map((row) => displayOrder.map((i) => row[i]));
  downloadCsv("在庫一覧_" + new Date().toISOString().slice(0, 10) + ".csv", displayOrder.map((i) => invHeaders[i]), rows);
});

// row(invRows内の該当行オブジェクト、任意)を渡すと保存成功時にそのままキャッシュへ反映する。
// 在庫タブ・棚卸タブは同じ invRows 配列を参照して描画しているため、これだけで両タブが
// リロードなしに最新値を表示できる(仕入価格(円)はどちらのタブから編集しても同じ列を更新する)。
async function saveInventoryField(tr, pid, header, value, row) {
  tr.className = "saving";
  const token = getToken();
  const isNum = INV_EDITABLE_NUM.includes(header);
  const body = { 商品ID: pid };
  body[INV_FIELD_KEY[header]] = isNum ? (value === "" ? "" : Number(value)) : value;
  try {
    const r = await fetch("/api/inventory", {
      method: "PATCH",
      headers: { "Content-Type": "application/json", Authorization: "Bearer " + token },
      body: JSON.stringify(body),
    });
    if (!r.ok) { tr.className = "error"; return; }
    tr.className = "saved";
    setTimeout(() => { tr.className = ""; }, 1500);
    if (row) {
      const idx = invHeaders.indexOf(header);
      if (idx !== -1) row[idx] = isNum ? (value === "" ? null : Number(value)) : value;
    }
  } catch (e) {
    tr.className = "error";
  }
}

document.getElementById("ord-q").addEventListener("input", renderOrders);
document.getElementById("inv-q").addEventListener("input", renderInventory);

// ---- 相違 ----
async function loadDiscrepancies() {
  const tbody = document.getElementById("disc-tbody");
  tbody.innerHTML = "<tr><td colspan='7'>読み込み中...</td></tr>";
  const token = getToken();
  try {
    const r = await fetch("/api/inventory/discrepancies", { headers: { Authorization: "Bearer " + token } });
    const data = await r.json();
    if (!r.ok) { tbody.innerHTML = "<tr><td colspan='7'>エラー: " + (data.error || r.status) + "</td></tr>"; return; }
    renderDiscrepancies(data.rows);
  } catch (e) {
    tbody.innerHTML = "<tr><td colspan='7'>通信エラー: " + e.message + "</td></tr>";
  }
  loadUnlinked();
}

async function loadUnlinked() {
  const tbody = document.getElementById("unlinked-tbody");
  tbody.innerHTML = "<tr><td colspan='6'>読み込み中...</td></tr>";
  const token = getToken();
  try {
    const r = await fetch("/api/inventory/unlinked", { headers: { Authorization: "Bearer " + token } });
    const data = await r.json();
    if (!r.ok) { tbody.innerHTML = "<tr><td colspan='6'>エラー: " + (data.error || r.status) + "</td></tr>"; return; }
    document.getElementById("unlinked-count").textContent = data.count.toLocaleString("ja-JP") + " 件";
    if (!data.count) { tbody.innerHTML = "<tr><td colspan='6'>未紐付けの出品はありません</td></tr>"; return; }
    tbody.innerHTML = data.rows.map((r) =>
      "<tr><td><span class='site-chip'>" + (r[0] || "") + "</span></td><td>" + (r[1] || "") + "</td><td class='truncate'>" + (r[2] || "") +
      "</td><td class='num'>" + fmt(r[3]) + "</td><td class='num'>" + fmt(r[5]) + "</td><td>" + (r[7] ? String(r[7]).slice(0, 10) : "") + "</td></tr>"
    ).join("");
  } catch (e) {
    tbody.innerHTML = "<tr><td colspan='6'>通信エラー: " + e.message + "</td></tr>";
  }
}
function renderDiscrepancies(rows) {
  document.getElementById("disc-count").textContent = rows.length.toLocaleString("ja-JP") + " 件";
  const tbody = document.getElementById("disc-tbody");
  if (!rows.length) { tbody.innerHTML = "<tr><td colspan='7'>相違はありません</td></tr>"; return; }
  const bySite = (row, site) => row.countries.find((c) => c.site === site);
  tbody.innerHTML = rows.map((row) => {
    const cell = (site) => {
      const c = bySite(row, site);
      if (!c) return "<td class='num'>-</td>";
      return "<td class='num" + (c.mismatch ? " mismatch-cell" : "") + "'>" + c.value.toLocaleString("ja-JP") + "</td>";
    };
    return "<tr class='" + (row.unconfirmed ? "row-unconfirmed" : "") + "'><td>" + row.商品ID + "</td><td class='truncate'>" + (row.商品名 || "") + "</td><td class='num'>" +
      row.リアル在庫.toLocaleString("ja-JP") + "</td>" + cell("US") + cell("UK") + cell("AU") + "<td>" + (row.リアル在庫確認日 || "(未確認)") + "</td></tr>";
  }).join("");
}
document.getElementById("disc-refresh-btn").addEventListener("click", loadDiscrepancies);

// ---- 棚卸 ----
const STK_PAGE_SIZE = 300;
let stkPage = 1;
let stkPrevQ = "";
// 列見出しクリックによる並び替え。nullの間は既存のプルダウン(stk-sort)が有効。
// {key, dir} が入っている間はこちらを優先する(プルダウンはstk-sort変更時にクリアする)。
let stkHeaderSort = null;

function loadStocktakeList() { stkPage = 1; renderStocktake(); }

function stkRowMatches(row, idx, q) {
  const pid = row[idx.pid];
  const name = row[idx.name] || "";
  const jaName = row[idx.jaName] || "";
  const sku = row[idx.sku] || ""; // 検索対象のみ。一覧には表示しない
  const searchable = (String(pid || "") + " " + name + " " + jaName + " " + sku).toLowerCase();
  return !q || searchable.indexOf(q) !== -1;
}

// 列見出しクリックで並び替え可能な列(画像・操作列は対象外)。idxオブジェクトのキー名と対応させる。
const STK_SORTABLE_COLUMNS = [
  { key: "pid", label: "商品ID" },
  { key: "name", label: "商品名(英語)" },
  { key: "jaName", label: "商品名(日本語)" },
  null,
  { key: "price", label: "仕入価格(円)", num: true },
  { key: "real", label: "リアル在庫", num: true },
  { key: "staged", label: "棚卸入力数量", num: true },
  { key: "stagedAt", label: "入力日時" },
];
const STK_HEADER_NUM_KEYS = new Set(["price", "real", "staged"]);

// 空欄/未入力は昇順・降順にかかわらず必ず末尾に置く(0は有効な数値として通常どおり比較する)。
function compareStkHeaderValue(av, bv, isNum, dir) {
  const aEmpty = av === "" || av === null || av === undefined;
  const bEmpty = bv === "" || bv === null || bv === undefined;
  if (aEmpty && bEmpty) return 0;
  if (aEmpty) return 1;
  if (bEmpty) return -1;
  if (isNum) return (Number(av) - Number(bv)) * dir;
  return String(av).localeCompare(String(bv), "ja", { numeric: true, sensitivity: "base" }) * dir;
}

function sortStocktakeRowsByHeader(rows, idx, headerSort) {
  const field = idx[headerSort.key];
  const isNum = STK_HEADER_NUM_KEYS.has(headerSort.key);
  return rows.slice().sort((a, b) => compareStkHeaderValue(a[field], b[field], isNum, headerSort.dir));
}

function renderStocktakeHead() {
  const thead = document.getElementById("stk-thead");
  thead.innerHTML = STK_SORTABLE_COLUMNS.map((col) => {
    if (!col) return "<th>画像</th>";
    const isSortCol = stkHeaderSort && stkHeaderSort.key === col.key;
    const arrow = isSortCol ? '<span class="sort-arrow">' + (stkHeaderSort.dir === 1 ? "▲" : "▼") + '</span>' : "";
    return '<th class="sortable' + (col.num ? " num" : "") + '" data-key="' + col.key + '">' + col.label + arrow + '</th>';
  }).join("");
  thead.querySelectorAll("th.sortable").forEach((th) => {
    th.addEventListener("click", () => {
      const key = th.dataset.key;
      // 1回目クリックは昇順、同じ列の再クリックで昇順/降順を切り替える。別の列は昇順から開始する。
      stkHeaderSort = { key, dir: stkHeaderSort && stkHeaderSort.key === key ? -stkHeaderSort.dir : 1 };
      // ヘッダーソートを優先させ、プルダウンは通常状態(先頭の商品ID順)へ戻す。
      document.getElementById("stk-sort").value = "pid";
      stkPage = 1;
      renderStocktake();
    });
  });
}

function sortStocktakeRows(rows, idx, sortKey) {
  const sorted = rows.slice();
  const textCompare = (a, b, field) => String(a[field] || "").localeCompare(String(b[field] || ""), "ja", { numeric: true, sensitivity: "base" });
  const numberCompare = (a, b, field) => {
    const av = a[field] === "" || a[field] === null || a[field] === undefined ? Infinity : Number(a[field]);
    const bv = b[field] === "" || b[field] === null || b[field] === undefined ? Infinity : Number(b[field]);
    return av - bv;
  };
  sorted.sort((a, b) => {
    if (sortKey === "name") return textCompare(a, b, idx.name);
    if (sortKey === "ja-name") return textCompare(a, b, idx.jaName);
    if (sortKey === "price") return numberCompare(a, b, idx.price);
    if (sortKey === "real") return numberCompare(a, b, idx.real);
    if (sortKey === "unchecked") return Number(Boolean(a[idx.checked])) - Number(Boolean(b[idx.checked]));
    return textCompare(a, b, idx.pid);
  });
  return sorted;
}

function buildStocktakeRow(row, idx) {
  const pid = row[idx.pid];
  const name = row[idx.name] || "";
  const jaName = row[idx.jaName] || "";
  const tr = document.createElement("tr");
  const pidTd = document.createElement("td");
  pidTd.textContent = pid;
  tr.appendChild(pidTd);
  const nameTd = document.createElement("td");
  nameTd.className = "stk-en-name";
  nameTd.textContent = name;
  nameTd.title = name;
  tr.appendChild(nameTd);

  const jaTd = document.createElement("td");
  const jaInp = document.createElement("textarea");
  jaInp.className = "ja-name-input";
  jaInp.rows = 2;
  jaInp.value = jaName;
  jaInp.addEventListener("change", () => saveInventoryTextField(tr, pid, "日本語商品名", jaInp.value, () => { row[idx.jaName] = jaInp.value; }));
  jaTd.appendChild(jaInp);
  tr.appendChild(jaTd);

  const imgTd = document.createElement("td");
  const imageUrl = row[idx.image];
  const thumbBtn = document.createElement("button");
  if (imageUrl) {
    thumbBtn.className = "thumb-btn";
    const img = document.createElement("img");
    img.src = imageUrl; img.loading = "lazy"; img.alt = name;
    thumbBtn.appendChild(img);
    thumbBtn.addEventListener("click", () => openImageModal(imageUrl, pid, name, jaName));
  } else {
    thumbBtn.className = "thumb-btn no-image";
    thumbBtn.textContent = "画像なし";
    thumbBtn.disabled = true;
  }
  imgTd.appendChild(thumbBtn);
  tr.appendChild(imgTd);

  const priceTd = document.createElement("td");
  priceTd.className = "num";
  const priceInp = document.createElement("input");
  priceInp.type = "number"; priceInp.step = "any";
  priceInp.value = row[idx.price] === null || row[idx.price] === undefined ? "" : row[idx.price];
  priceInp.addEventListener("change", () => saveInventoryField(tr, pid, "仕入価格(円)", priceInp.value, row));
  priceTd.appendChild(priceInp);
  tr.appendChild(priceTd);

  const realTd = document.createElement("td");
  realTd.className = "num";
  realTd.textContent = row[idx.real] === null || row[idx.real] === undefined ? "(未設定)" : Number(row[idx.real]).toLocaleString("ja-JP");
  tr.appendChild(realTd);

  const stagedTd = document.createElement("td");
  stagedTd.className = "num";
  const inp = document.createElement("input");
  inp.type = "number"; inp.step = "1"; inp.min = "0"; inp.className = "qty-input";
  preventQtyWheelChange(inp);
  inp.value = row[idx.staged] === null || row[idx.staged] === undefined ? "" : row[idx.staged];
  const checkBadge = document.createElement("span");
  checkBadge.className = "stk-check-badge";
  checkBadge.title = "棚卸済み";
  checkBadge.textContent = "✓";
  checkBadge.style.display = row[idx.checked] ? "inline-flex" : "none";
  inp.addEventListener("change", () => saveStocktakeQty(tr, pid, inp.value, (savedAt) => {
    row[idx.staged] = inp.value === "" ? null : Number(inp.value);
    row[idx.stagedAt] = savedAt;
    atTd.textContent = savedAt ? savedAt.slice(0, 16).replace("T", " ") : "";
    // 0を含め、数量が正常保存されたら棚卸済みチェックを立てる(未入力に戻しても自動では解除しない)
    if (inp.value !== "" && !row[idx.checked]) {
      row[idx.checked] = true;
      checkBadge.style.display = "inline-flex";
      renderStocktakeProgress();
    }
  }));
  stagedTd.appendChild(inp);
  stagedTd.appendChild(checkBadge);
  tr.appendChild(stagedTd);

  const atTd = document.createElement("td");
  atTd.textContent = row[idx.stagedAt] ? String(row[idx.stagedAt]).slice(0, 16).replace("T", " ") : "";
  tr.appendChild(atTd);
  return tr;
}

// 汎用ページャー描画。currentPageは現在ページ、onGoToPageはページ番号を受け取って移動する関数。
function renderPager(container, currentPage, totalPages, onGoToPage) {
  container.innerHTML = "";
  if (totalPages <= 1) return;
  const mkBtn = (label, page, opts) => {
    const b = document.createElement("button");
    b.textContent = label;
    if (opts && opts.active) b.className = "active";
    if (opts && opts.disabled) b.disabled = true;
    b.addEventListener("click", () => onGoToPage(page));
    return b;
  };
  container.appendChild(mkBtn("前へ", currentPage - 1, { disabled: currentPage <= 1 }));
  for (let p = 1; p <= totalPages; p++) {
    container.appendChild(mkBtn(String(p), p, { active: p === currentPage }));
  }
  container.appendChild(mkBtn("次へ", currentPage + 1, { disabled: currentPage >= totalPages }));
}

function renderStocktake() {
  if (!invHeaders.length) { document.getElementById("stk-tbody").innerHTML = "<tr><td colspan='8'>在庫データが読み込まれていません</td></tr>"; return; }
  const idx = {
    pid: invHeaders.indexOf("商品ID"), name: invHeaders.indexOf("商品名"), jaName: invHeaders.indexOf("日本語商品名"),
    image: invHeaders.indexOf("画像URL"), sku: invHeaders.indexOf("SKU"),
    price: invHeaders.indexOf("仕入価格(円)"),
    real: invHeaders.indexOf("リアル在庫"), staged: invHeaders.indexOf("棚卸入力数量"), stagedAt: invHeaders.indexOf("棚卸入力日時"),
    checked: invHeaders.indexOf("棚卸チェック"),
  };
  renderStocktakeProgress(idx);
  renderStocktakeHead();
  const q = document.getElementById("stk-q").value.trim().toLowerCase();
  if (q !== stkPrevQ) stkPage = 1;
  stkPrevQ = q;
  const sortKey = document.getElementById("stk-sort").value;

  const tbody = document.getElementById("stk-tbody");
  const pagerTop = document.getElementById("stk-pager-top");
  const pagerBottom = document.getElementById("stk-pager-bottom");
  tbody.innerHTML = "";

  // 元のinvRowsは変更せず、検索結果のコピーだけを並び替えてからページ分けする。
  // 検索→ソート→ページネーションの順を維持し、表示中のページだけでなく
  // 検索結果全体を並び替えてから300件ずつに分割する。
  const filteredRows = invRows.filter((row) => stkRowMatches(row, idx, q));
  const displayRows = stkHeaderSort
    ? sortStocktakeRowsByHeader(filteredRows, idx, stkHeaderSort)
    : sortStocktakeRows(filteredRows, idx, sortKey);

  const totalPages = Math.max(1, Math.ceil(displayRows.length / STK_PAGE_SIZE));
  if (stkPage > totalPages) stkPage = totalPages;
  if (stkPage < 1) stkPage = 1;
  const start = (stkPage - 1) * STK_PAGE_SIZE;
  const pageRows = displayRows.slice(start, start + STK_PAGE_SIZE);
  pageRows.forEach((row) => tbody.appendChild(buildStocktakeRow(row, idx)));

  const goToStkPage = (p) => { stkPage = p; renderStocktake(); };
  renderPager(pagerTop, stkPage, totalPages, goToStkPage);
  renderPager(pagerBottom, stkPage, totalPages, goToStkPage);
  const rangeText = (displayRows.length ? (start + 1) : 0).toLocaleString("ja-JP") + "〜" + Math.min(start + STK_PAGE_SIZE, displayRows.length).toLocaleString("ja-JP");
  document.getElementById("stk-count").textContent = q
    ? rangeText + " / " + displayRows.length.toLocaleString("ja-JP") + " 件ヒット(全" + invRows.length.toLocaleString("ja-JP") + "件中、" + stkPage + " / " + totalPages + " ページ)"
    : rangeText + " / " + displayRows.length.toLocaleString("ja-JP") + " 件(" + stkPage + " / " + totalPages + " ページ)";
}
document.getElementById("stk-q").addEventListener("input", renderStocktake);
document.getElementById("stk-sort").addEventListener("change", () => { stkHeaderSort = null; stkPage = 1; renderStocktake(); });

// 検索条件やページに関わらず、全商品を対象に棚卸チェックの進捗を集計する
function renderStocktakeProgress(idx) {
  const checkedIdx = idx ? idx.checked : invHeaders.indexOf("棚卸チェック");
  const total = invRows.length;
  const checked = checkedIdx === -1 ? 0 : invRows.filter((row) => Boolean(row[checkedIdx])).length;
  const pct = total ? (checked / total * 100).toFixed(1) : "0.0";
  document.getElementById("stk-progress").textContent =
    "棚卸進捗　" + checked.toLocaleString("ja-JP") + " / " + total.toLocaleString("ja-JP") + "件(" + pct + "%)";
}

document.getElementById("stk-reset-checks-btn").addEventListener("click", async () => {
  if (!confirm("棚卸済みチェックをすべて解除しますか?\\n棚卸入力数量やリアル在庫は変更されません。")) return;
  const statusEl = document.getElementById("stk-reset-status");
  statusEl.textContent = "解除中...";
  statusEl.className = "hint";
  const token = getToken();
  try {
    const r = await fetch("/api/inventory/stocktake/reset-checks", {
      method: "POST",
      headers: { Authorization: "Bearer " + token },
    });
    const data = await r.json();
    if (!r.ok) { statusEl.textContent = "失敗: " + (data.error || r.status); statusEl.className = "hint ng"; return; }
    const checkedIdx = invHeaders.indexOf("棚卸チェック");
    if (checkedIdx !== -1) invRows.forEach((row) => { row[checkedIdx] = null; });
    renderStocktake();
    statusEl.textContent = data.reset.toLocaleString("ja-JP") + "件のチェックを解除しました";
    statusEl.className = "hint ok";
  } catch (e) {
    statusEl.textContent = "通信エラー: " + e.message;
    statusEl.className = "hint ng";
  }
});

function openImageModal(url, pid, name, jaName) {
  document.getElementById("img-modal-img").src = url;
  document.getElementById("img-modal-caption").textContent = pid + " / " + name + (jaName ? " / " + jaName : "");
  document.getElementById("img-modal-overlay").classList.add("show");
}
function closeImageModal() {
  document.getElementById("img-modal-overlay").classList.remove("show");
  document.getElementById("img-modal-img").src = "";
}
document.getElementById("img-modal-close").addEventListener("click", closeImageModal);
document.getElementById("img-modal-overlay").addEventListener("click", (evt) => {
  if (evt.target.id === "img-modal-overlay") closeImageModal();
});

async function saveStocktakeQty(tr, pid, value, onSuccess) {
  tr.className = "saving";
  const token = getToken();
  const body = { 商品ID: pid, 棚卸入力数量: value === "" ? "" : Number(value) };
  try {
    const r = await fetch("/api/inventory", {
      method: "PATCH",
      headers: { "Content-Type": "application/json", Authorization: "Bearer " + token },
      body: JSON.stringify(body),
    });
    if (!r.ok) { tr.className = "error"; return; }
    tr.className = "saved";
    setTimeout(() => { tr.className = ""; }, 1500);
    // ページ切り替え後も入力値が消えないよう、保存成功時にクライアント側のキャッシュ(invRows)も更新する
    if (onSuccess) onSuccess(value === "" ? null : new Date().toISOString());
  } catch (e) {
    tr.className = "error";
  }
}

async function saveInventoryTextField(tr, pid, field, value, onSuccess) {
  tr.className = "saving";
  const token = getToken();
  const body = { 商品ID: pid };
  body[field] = value;
  try {
    const r = await fetch("/api/inventory", {
      method: "PATCH",
      headers: { "Content-Type": "application/json", Authorization: "Bearer " + token },
      body: JSON.stringify(body),
    });
    if (!r.ok) { tr.className = "error"; return; }
    tr.className = "saved";
    setTimeout(() => { tr.className = ""; }, 1500);
    if (onSuccess) onSuccess();
  } catch (e) {
    tr.className = "error";
  }
}

document.getElementById("stk-preview-btn").addEventListener("click", async () => {
  const statusEl = document.getElementById("stk-preview-status");
  statusEl.textContent = "読み込み中...";
  statusEl.className = "hint";
  const token = getToken();
  try {
    const r = await fetch("/api/inventory/stocktake/preview", { headers: { Authorization: "Bearer " + token } });
    const data = await r.json();
    if (!r.ok) { statusEl.textContent = "失敗: " + (data.error || r.status); statusEl.className = "hint ng"; return; }
    statusEl.textContent = "";
    renderStocktakePreview(data);
  } catch (e) {
    statusEl.textContent = "通信エラー: " + e.message;
    statusEl.className = "hint ng";
  }
});

let stocktakePreviewRows = [];
function renderStocktakePreview(data) {
  stocktakePreviewRows = data.rows;
  const summaryEl = document.getElementById("stk-preview-summary");
  summaryEl.style.display = "grid";
  summaryEl.innerHTML = [
    { label: "対象件数", value: data.targetCount },
    { label: "差異あり", value: data.diffCount },
    { label: "増加", value: data.increased },
    { label: "減少", value: data.decreased },
    { label: "異常値の疑い", value: data.abnormalCount, cls: data.abnormalCount > 0 ? "bad" : "" },
  ].map((t) => '<div class="kpi"><div class="label">' + t.label + '</div><div class="value' + (t.cls ? " " + t.cls : "") + '">' + t.value + '</div></div>').join("");

  const wrap = document.getElementById("stk-preview-table-wrap");
  wrap.style.display = "block";
  document.getElementById("stk-preview-tbody").innerHTML = data.rows.map((r) =>
    "<tr class='" + (r.異常値 ? "row-abnormal" : "") + "'><td>" + r.商品ID + "</td><td class='truncate'>" + (r.商品名 || "") + "</td><td class='num'>" +
    r.リアル在庫.toLocaleString("ja-JP") + "</td><td class='num'>" + r.棚卸入力数量.toLocaleString("ja-JP") + "</td><td class='num" + (r.差異 < 0 ? " mismatch-cell" : "") + "'>" +
    (r.差異 > 0 ? "+" : "") + r.差異.toLocaleString("ja-JP") + "</td><td>" + (r.異常値 ? "⚠ 要確認" : "") + "</td></tr>"
  ).join("");

  document.getElementById("stk-confirm-row").style.display = data.targetCount > 0 ? "flex" : "none";
}

document.getElementById("stk-confirm-btn").addEventListener("click", async () => {
  if (!stocktakePreviewRows.length) return;
  const abnormal = stocktakePreviewRows.filter((r) => r.異常値).length;
  const msg = "対象" + stocktakePreviewRows.length + "件をリアル在庫へ反映します。" +
    (abnormal > 0 ? "うち異常値の疑いがある行が" + abnormal + "件あります。" : "") + "よろしいですか?(元に戻せません)";
  if (!confirm(msg)) return;
  const statusEl = document.getElementById("stk-confirm-status");
  statusEl.textContent = "確定中...";
  statusEl.className = "hint";
  const token = getToken();
  try {
    const r = await fetch("/api/inventory/stocktake/confirm", {
      method: "POST",
      headers: { "Content-Type": "application/json", Authorization: "Bearer " + token },
      body: JSON.stringify({ 商品ID: stocktakePreviewRows.map((r) => r.商品ID) }),
    });
    const data = await r.json();
    if (!r.ok) { statusEl.textContent = "失敗: " + (data.error || r.status); statusEl.className = "hint ng"; return; }
    statusEl.textContent = data.confirmed + "件を確定しました";
    statusEl.className = "hint ok";
    document.getElementById("stk-preview-summary").style.display = "none";
    document.getElementById("stk-preview-table-wrap").style.display = "none";
    document.getElementById("stk-confirm-row").style.display = "none";
    stocktakePreviewRows = [];
    await loadAll();
    renderStocktake();
  } catch (e) {
    statusEl.textContent = "通信エラー: " + e.message;
    statusEl.className = "hint ng";
  }
});

// ---- 決算 ----
document.getElementById("cls-asof").value = new Date().toISOString().slice(0, 10);

document.getElementById("cls-checklist-btn").addEventListener("click", async () => {
  const asOf = document.getElementById("cls-asof").value;
  const statusEl = document.getElementById("cls-status");
  statusEl.textContent = "読み込み中...";
  statusEl.className = "hint";
  const token = getToken();
  try {
    const r = await fetch("/api/closing/checklist?asOf=" + encodeURIComponent(asOf), { headers: { Authorization: "Bearer " + token } });
    const data = await r.json();
    if (!r.ok) { statusEl.textContent = "失敗: " + (data.error || r.status); statusEl.className = "hint ng"; return; }
    statusEl.textContent = data.count + "件が未確認です";
    statusEl.className = "hint";
    document.getElementById("cls-checklist-wrap").style.display = "block";
    document.getElementById("cls-checklist-tbody").innerHTML = data.rows.map((row) =>
      "<tr><td>" + row.商品ID + "</td><td class='truncate'>" + (row.商品名 || "") + "</td><td class='num'>" +
      (row.リアル在庫 === null || row.リアル在庫 === undefined ? "" : Number(row.リアル在庫).toLocaleString("ja-JP")) + "</td><td>" + (row.リアル在庫確認日 || "(未確認)") + "</td></tr>"
    ).join("");
  } catch (e) {
    statusEl.textContent = "通信エラー: " + e.message;
    statusEl.className = "hint ng";
  }
});

document.getElementById("cls-export-btn").addEventListener("click", () => {
  const asOf = document.getElementById("cls-asof").value;
  const token = getToken();
  window.location.href = "/api/closing/export?asOf=" + encodeURIComponent(asOf) + "&token=" + encodeURIComponent(token);
});

document.getElementById("cls-snapshot-btn").addEventListener("click", async () => {
  const asOf = document.getElementById("cls-asof").value;
  if (!confirm("基準日「" + asOf + "」の棚卸資産スナップショットを保存します。同じ基準日では1回しか保存できません。よろしいですか?")) return;
  const statusEl = document.getElementById("cls-status");
  statusEl.textContent = "保存中...";
  statusEl.className = "hint";
  const token = getToken();
  try {
    const r = await fetch("/api/closing/snapshot", {
      method: "POST",
      headers: { "Content-Type": "application/json", Authorization: "Bearer " + token },
      body: JSON.stringify({ asOf }),
    });
    const data = await r.json();
    if (!r.ok) { statusEl.textContent = "失敗: " + (data.error || r.status); statusEl.className = "hint ng"; return; }
    statusEl.textContent = "保存しました(合計評価額 ¥" + Math.round(data.totalValue).toLocaleString("ja-JP") + ")";
    statusEl.className = "hint ok";
    loadSnapshotList();
  } catch (e) {
    statusEl.textContent = "通信エラー: " + e.message;
    statusEl.className = "hint ng";
  }
});

async function loadSnapshotList() {
  const listEl = document.getElementById("cls-snapshot-list");
  const token = getToken();
  try {
    const r = await fetch("/api/closing/snapshots", { headers: { Authorization: "Bearer " + token } });
    const data = await r.json();
    if (!r.ok) { listEl.innerHTML = "<li>エラー: " + (data.error || r.status) + "</li>"; return; }
    if (!data.files.length) { listEl.innerHTML = "<li>保存済みのスナップショットはありません</li>"; return; }
    listEl.innerHTML = data.files.map((f) =>
      "<li><a href='#' data-file='" + f + "' class='snapshot-link'>" + f + "</a></li>"
    ).join("");
    listEl.querySelectorAll(".snapshot-link").forEach((a) => {
      a.addEventListener("click", (evt) => {
        evt.preventDefault();
        window.location.href = "/download/snapshots/" + encodeURIComponent(a.dataset.file) + "?token=" + encodeURIComponent(getToken());
      });
    });
  } catch (e) {
    listEl.innerHTML = "<li>通信エラー: " + e.message + "</li>";
  }
}

document.getElementById("ebay-rebuild-btn").addEventListener("click", async () => {
  const btn = document.getElementById("ebay-rebuild-btn");
  const statusEl = document.getElementById("ebay-rebuild-status");
  btn.disabled = true;
  statusEl.className = "hint";
  statusEl.textContent = "読み込み中...(30秒ほどかかります)";
  const token = getToken();
  try {
    const r = await fetch("/api/ebay/inventory/rebuild", {
      method: "POST",
      headers: { Authorization: "Bearer " + token },
    });
    const data = await r.json();
    if (!r.ok) { statusEl.textContent = "失敗: " + (data.error || r.status); statusEl.className = "hint ng"; return; }
    statusEl.textContent = "更新完了(USベース" + data.total + "件・新規" + data.new + "件・削除" + data.deletedCount + "件・要確認" + data.ambiguousMatches + "件・UK/AU保管" + data.ukAuOrphanCount + "件)";
    statusEl.className = "hint ok";
    await loadAll();
  } catch (e) {
    statusEl.textContent = "通信エラー: " + e.message;
    statusEl.className = "hint ng";
  } finally {
    btn.disabled = false;
  }
});

document.getElementById("inv-csv-upload").addEventListener("click", async () => {
  const fileInput = document.getElementById("inv-csv-file");
  const statusEl = document.getElementById("inv-csv-status");
  statusEl.className = "hint"; statusEl.textContent = "";
  if (!fileInput.files[0]) { statusEl.textContent = "CSVファイルを選んでください"; statusEl.className = "hint ng"; return; }
  const token = getToken();
  statusEl.textContent = "更新中...(数秒かかります)";
  try {
    const text = await fileInput.files[0].text();
    const r = await fetch("/api/inventory/rebuild", {
      method: "POST",
      headers: { Authorization: "Bearer " + token, "Content-Type": "text/csv; charset=utf-8" },
      body: text,
    });
    const data = await r.json();
    if (!r.ok) { statusEl.textContent = "失敗: " + (data.error || r.status); statusEl.className = "hint ng"; return; }
    statusEl.textContent = "更新完了(全" + data.total + "件・新規" + data.new + "件・削除扱い" + data.removedNew + "件・要確認" + data.anomalous + "件)";
    statusEl.className = "hint ok";
    await loadAll();
  } catch (e) {
    statusEl.textContent = "通信エラー: " + e.message;
    statusEl.className = "hint ng";
  }
});

document.getElementById("ord-xlsx-upload").addEventListener("click", async () => {
  const fileInput = document.getElementById("ord-xlsx-file");
  const statusEl = document.getElementById("ord-xlsx-status");
  statusEl.className = "hint"; statusEl.textContent = "";
  if (!fileInput.files[0]) { statusEl.textContent = "xlsxファイルを選んでください"; statusEl.className = "hint ng"; return; }
  const token = getToken();
  statusEl.textContent = "取り込み中...";
  try {
    const buf = await fileInput.files[0].arrayBuffer();
    const r = await fetch("/api/import", {
      method: "POST",
      headers: { Authorization: "Bearer " + token, "Content-Type": "application/octet-stream" },
      body: buf,
    });
    const data = await r.json();
    if (!r.ok) { statusEl.textContent = "失敗: " + (data.error || r.status); statusEl.className = "hint ng"; return; }
    statusEl.textContent = "成功しました: " + (data.message || "取り込み完了");
    statusEl.className = "hint ok";
    await loadAll();
  } catch (e) {
    statusEl.textContent = "通信エラー: " + e.message;
    statusEl.className = "hint ng";
  }
});

const parseOrderText = AgateOrderParser.parseOrderText;

function renderOrderParseReview(parsed) {
  const review = document.getElementById("ne-items-review");
  const summary = document.getElementById("ne-parse-summary");
  const body = document.getElementById("ne-items-body");
  review.style.display = "block";
  summary.className = "order-parse-summary " + (parsed.parseStatus === "OK" ? "ok" : "ng");
  summary.replaceChildren();

  const summaryValues = [
    ["明細数", parsed.itemCount],
    ["数量合計", parsed.quantityTotal],
    ["小計", parsed.subtotalQuantity === null ? "取得なし" : parsed.subtotalQuantity],
    ["解析状態", parsed.parseStatus],
  ];
  for (const pair of summaryValues) {
    const span = document.createElement("span");
    span.textContent = pair[0] + "：" + pair[1];
    summary.appendChild(span);
  }
  if (parsed.parseErrors.length) {
    const error = document.createElement("span");
    error.textContent = parsed.parseErrors.join(" / ");
    error.className = "order-item-status-ng";
    summary.appendChild(error);
  }

  body.replaceChildren();
  parsed.items.forEach((item, index) => {
    const tr = document.createElement("tr");
    const values = [index + 1, item.title, item.ebayItemId, item.quantity === null ? "" : item.quantity, item.sku];
    values.forEach((value, columnIndex) => {
      const td = document.createElement("td");
      td.textContent = value;
      if (columnIndex === 3) td.className = "num";
      tr.appendChild(td);
    });
    const status = document.createElement("td");
    status.textContent = item.parseStatus === "OK" ? "OK" : "要確認：" + item.errors.join(" / ");
    status.className = item.parseStatus === "OK" ? "order-item-status-ok" : "order-item-status-ng";
    tr.appendChild(status);
    body.appendChild(tr);
  });
}

function updateFeePreview() {
  const usd = Number(document.getElementById("ne-usd").value);
  const rate = Number(document.getElementById("ne-rate").value);
  const preview = document.getElementById("ne-fee-preview");
  if (Number.isFinite(usd) && Number.isFinite(rate) && usd > 0 && rate > 0) {
    const revenueJpy = usd * rate;
    const fee = Math.round(revenueJpy * 0.03);
    preview.value = "¥" + fee.toLocaleString("ja-JP") + "(収益円 ¥" + Math.round(revenueJpy).toLocaleString("ja-JP") + ")";
  } else {
    preview.value = "";
  }
}
document.getElementById("ne-usd").addEventListener("input", updateFeePreview);
document.getElementById("ne-rate").addEventListener("input", updateFeePreview);
preventQtyWheelChange(document.getElementById("ne-qty"));

document.getElementById("ne-parse-btn").addEventListener("click", () => {
  const parsed = parseOrderText(document.getElementById("ne-paste").value);
  document.getElementById("ne-order").value = parsed.orderNo;
  document.getElementById("ne-date").value = parsed.date;
  document.getElementById("ne-site").value = parsed.site;
  document.getElementById("ne-usd").value = parsed.usd;
  document.getElementById("ne-note").value = parsed.note;
  document.getElementById("ne-qty").value = parsed.qty;
  document.getElementById("ne-cost").value = "";
  document.getElementById("ne-shipping").value = "";
  document.getElementById("ne-packing").value = "50";
  const pidInput = document.getElementById("ne-pid");
  pidInput.value = "";
  // 貼り付け原文がある新規注文は明細方式だけを在庫変更の正とし、旧単一Pxxxx経路と併用しない。
  pidInput.disabled = true;
  updateFeePreview();
  renderOrderParseReview(parsed);
  document.getElementById("ne-review").style.display = "grid";
  document.getElementById("ne-submit-row").style.display = "flex";
  document.getElementById("ne-status").textContent = parsed.parseStatus === "OK"
    ? "内容を確認して登録してください。在庫は販売サイト + eBay Item IDの完全一致だけで明細単位に反映します"
    : "売上情報は登録できますが、解析状態が要確認のためリアル在庫には反映しません";
  document.getElementById("ne-status").className = "hint " + (parsed.parseStatus === "OK" ? "" : "ng");
});

document.getElementById("ne-submit-btn").addEventListener("click", async () => {
  const statusEl = document.getElementById("ne-status");
  const body = {
    注文番号: document.getElementById("ne-order").value.trim(),
    日付: document.getElementById("ne-date").value.trim(),
    サイト: document.getElementById("ne-site").value.trim() || "不明",
    商品メモ: document.getElementById("ne-note").value.trim(),
    収益USD: document.getElementById("ne-usd").value,
    ドル円レート: document.getElementById("ne-rate").value,
    注文貼付テキスト: document.getElementById("ne-paste").value,
  };
  const cost = document.getElementById("ne-cost").value;
  const shipping = document.getElementById("ne-shipping").value;
  const packing = document.getElementById("ne-packing").value;
  const qty = document.getElementById("ne-qty").value;
  if (cost !== "") body.仕入原価円 = cost;
  if (shipping !== "") body.送料円 = shipping;
  if (packing !== "") body.梱包費円 = packing;
  if (qty !== "") body.数量 = qty;
  const pid = document.getElementById("ne-pid").value.trim();
  if (pid !== "") body.商品ID = pid;
  if (!body.注文番号 || !body.日付 || !body.収益USD || !body.ドル円レート) {
    statusEl.textContent = "注文番号・日付・収益USD・ドル円レートは必須です";
    statusEl.className = "hint ng";
    return;
  }
  statusEl.textContent = "登録中...";
  statusEl.className = "hint";
  try {
    const token = getToken();
    const r = await fetch("/api/orders", {
      method: "POST",
      headers: { "Content-Type": "application/json", Authorization: "Bearer " + token },
      body: JSON.stringify(body),
    });
    const data = await r.json();
    if (!r.ok) { statusEl.textContent = "失敗: " + (data.error || r.status); statusEl.className = "hint ng"; return; }
    const inv = data.inventory;
    const inventorySummary = inv
      ? "・在庫反映 " + inv.applied + "/" + inv.total + "明細・未反映 " + inv.unapplied + "明細" + (inv.conflict ? "・矛盾 " + inv.conflict + "明細" : "")
      : "";
    statusEl.textContent = (data.status === "already_registered" ? "既に登録済みです" : "登録しました")
      + (data.収益円 === undefined ? "" : "(収益円: " + data.収益円 + "・手数料: " + data.手数料円 + ")")
      + inventorySummary + (data.warning ? " ※" + data.warning : "");
    statusEl.className = "hint " + ((data.warning || (inv && inv.unapplied)) ? "ng" : "ok");
    document.getElementById("ne-paste").value = "";
    document.getElementById("ne-review").style.display = "none";
    document.getElementById("ne-items-review").style.display = "none";
    document.getElementById("ne-submit-row").style.display = "none";
    document.getElementById("ne-pid").disabled = false;
    await loadAll();
  } catch (e) {
    statusEl.textContent = "通信エラー: " + e.message;
    statusEl.className = "hint ng";
  }
});

if (getToken()) loadAll();
</script>
</body></html>`;

const server = http.createServer(async (req, res) => {
  if (req.method === "OPTIONS") {
    res.writeHead(204, {
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "GET, POST, PATCH, DELETE, OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type, Authorization",
    });
    return res.end();
  }

  let pathname = req.url;
  try {
    // URLのpathnameはパーセントエンコードされたまま返る(自動デコードされない)ため、
    // 日本語などを含むルート("/download/売上管理表.xlsx"等)と比較できるよう明示的にデコードする。
    pathname = decodeURIComponent(new URL(req.url, "http://localhost").pathname);
  } catch (e) {
    // フォールバック: req.urlをそのまま使う
  }

  if (req.method === "GET" && pathname === "/") {
    res.writeHead(200, { "Content-Type": "text/plain; charset=utf-8" });
    return res.end("agate-tool-server: OK");
  }
  if (req.method === "GET" && (pathname === "/orders" || pathname === "/dashboard")) {
    return sendHtml(res, DASHBOARD_PAGE);
  }
  if (req.method === "GET" && pathname === "/order-parser.js") {
    res.writeHead(200, { "Content-Type": "application/javascript; charset=utf-8", "Cache-Control": "no-store" });
    return fs.createReadStream(ORDER_PARSER_PATH).pipe(res);
  }

  const protectedRoutes = [
    "/api/orders", "/api/inventory", "/download/売上管理表.xlsx", "/api/import", "/api/import/inventory", "/api/inventory/rebuild", "/api/summary",
    "/ebay/connect", "/api/ebay/inventory", "/api/ebay/inventory/rebuild",
    "/api/inventory/discrepancies", "/api/inventory/unlinked", "/api/inventory/stocktake/preview", "/api/inventory/stocktake/confirm", "/api/inventory/stocktake/reset-checks", "/api/inventory/history",
    "/api/order-lines/unresolved", "/api/order-lines/resolve",
    "/api/closing/checklist", "/api/closing/export", "/api/closing/snapshot", "/api/closing/snapshots",
  ];
  const isProtected = protectedRoutes.includes(pathname) || pathname.startsWith("/download/snapshots/");
  if (isProtected && !isAuthorized(req)) {
    return sendJson(res, 401, { error: "認証に失敗しました(トークンを確認してください)" });
  }

  try {
    if (req.method === "GET" && pathname === "/ebay/connect") return await handleEbayConnect(req, res);
    if (req.method === "GET" && pathname === "/ebay/callback") return await handleEbayCallback(req, res);
    if (req.method === "GET" && pathname === "/api/ebay/inventory") return await handleEbayInventory(req, res);
    if (req.method === "POST" && pathname === "/api/ebay/inventory/rebuild") return await handleEbayInventoryRebuild(req, res);
    if (req.method === "POST" && pathname === "/api/orders") return await withSalesLock(() => handleAddOrder(req, res));
    if (req.method === "PATCH" && pathname === "/api/orders") return await withSalesLock(() => handlePatchOrder(req, res));
    if (req.method === "GET" && pathname === "/api/orders") return await handleListOrders(req, res);
    if (req.method === "GET" && pathname === "/download/売上管理表.xlsx") return await handleDownload(req, res);
    if (req.method === "POST" && pathname === "/api/import") return await handleImport(req, res);
    if (req.method === "GET" && pathname === "/api/inventory") return await handleListInventory(req, res);
    if (req.method === "GET" && pathname === "/api/summary") return await handleSummary(req, res);
    if (req.method === "PATCH" && pathname === "/api/inventory") return await handlePatchInventory(req, res);
    if (req.method === "DELETE" && pathname === "/api/orders") return await withSalesLock(() => handleDeleteOrders(req, res));
    if (req.method === "DELETE" && pathname === "/api/inventory") return await handleDeleteInventory(req, res);
    if (req.method === "POST" && pathname === "/api/import/inventory") return await handleImportInventory(req, res);
    if (req.method === "POST" && pathname === "/api/inventory/rebuild") return await handleRebuildInventory(req, res);
    if (req.method === "GET" && pathname === "/api/inventory/discrepancies") return await handleDiscrepancies(req, res);
    if (req.method === "GET" && pathname === "/api/inventory/unlinked") return await handleUnlinkedInventory(req, res);
    if (req.method === "GET" && pathname === "/api/inventory/stocktake/preview") return await handleStocktakePreview(req, res);
    if (req.method === "POST" && pathname === "/api/inventory/stocktake/confirm") return await handleStocktakeConfirm(req, res);
    if (req.method === "POST" && pathname === "/api/inventory/stocktake/reset-checks") return await handleStocktakeResetChecks(req, res);
    if (req.method === "GET" && pathname === "/api/inventory/history") return await handleInventoryHistory(req, res);
    if (req.method === "GET" && pathname === "/api/order-lines/unresolved") return await handleListUnresolvedOrderLines(req, res);
    if (req.method === "POST" && pathname === "/api/order-lines/resolve") return await handleResolveOrderLine(req, res);
    if (req.method === "GET" && pathname === "/api/closing/checklist") return await handleClosingChecklist(req, res);
    if (req.method === "GET" && pathname === "/api/closing/export") return await handleClosingExport(req, res);
    if (req.method === "POST" && pathname === "/api/closing/snapshot") return await handleClosingSnapshot(req, res);
    if (req.method === "GET" && pathname === "/api/closing/snapshots") return await handleListSnapshots(req, res);
    if (req.method === "GET" && pathname.startsWith("/download/snapshots/")) return await handleDownloadSnapshot(req, res, pathname);
  } catch (e) {
    console.error(e);
    const statusCode = Number(e && e.statusCode) || 500;
    return sendJson(res, statusCode, { error: statusCode === 500 ? "サーバー内部でエラーが発生しました" : e.message });
  }

  sendJson(res, 404, { error: "not found" });
});

server.listen(PORT, "127.0.0.1", () => {
  console.log(`agate-tool-server listening on 127.0.0.1:${PORT}`);
});
