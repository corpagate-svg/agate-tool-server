const http = require("http");
const crypto = require("crypto");
const path = require("path");
const fs = require("fs");
const ExcelJS = require("exceljs");

const PORT = process.env.PORT || 3000;
const API_TOKEN = process.env.API_TOKEN || "";
const DATA_DIR = path.join(__dirname, "..", "data");
const LEDGER_PATH = path.join(DATA_DIR, "売上管理表.xlsx");
const INVENTORY_PATH = path.join(DATA_DIR, "在庫管理表.xlsx");

const HEADERS = [
  "注文番号", "日付", "サイト", "商品メモ", "商品ID",
  "収益USD", "ドル円レート", "収益円", "手数料(円)",
  "仕入原価(円)", "送料(円)", "梱包費(円)", "最終利益(円)", "利益率",
];
const COL = { 注文番号: 1, 日付: 2, サイト: 3, 商品メモ: 4, 商品ID: 5, 収益USD: 6, ドル円レート: 7, 収益円: 8, 手数料: 9, 仕入原価: 10, 送料: 11, 梱包費: 12, 最終利益: 13, 利益率: 14 };

const INV_HEADERS = [
  "商品ID", "商品名", "バリエーション詳細", "在庫数(現物)", "出品国数", "在庫数不一致",
  "US_出品ID", "US価格(USD)", "US累計売却数", "UK_出品ID", "UK価格(GBP)", "UK累計売却数",
  "AU_出品ID", "AU価格(AUD)", "AU累計売却数", "仕入価格(円)", "仕入日", "仕入先", "備考",
];

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

async function loadWorkbook() {
  const wb = new ExcelJS.Workbook();
  if (fs.existsSync(LEDGER_PATH)) {
    await wb.xlsx.readFile(LEDGER_PATH);
    return wb;
  }
  const ws = wb.addWorksheet("記録");
  styleHeaderRow(ws, HEADERS);
  await wb.xlsx.writeFile(LEDGER_PATH);
  return wb;
}

async function loadInventoryWorkbook() {
  const wb = new ExcelJS.Workbook();
  if (fs.existsSync(INVENTORY_PATH)) {
    await wb.xlsx.readFile(INVENTORY_PATH);
    return wb;
  }
  const ws = wb.addWorksheet("在庫管理表");
  styleHeaderRow(ws, INV_HEADERS);
  await wb.xlsx.writeFile(INVENTORY_PATH);
  return wb;
}

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
  if (!m) return false;
  const given = Buffer.from(m[1]);
  const expected = Buffer.from(API_TOKEN);
  if (given.length !== expected.length) return false;
  return crypto.timingSafeEqual(given, expected);
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

function numOrNullCsv(v) {
  const n = parseFloat(v);
  return Number.isFinite(n) ? n : null;
}

function intOrNullCsv(v) {
  const n = parseInt(v, 10);
  return Number.isFinite(n) ? n : null;
}

function sendJson(res, status, obj) {
  const body = JSON.stringify(obj);
  res.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Access-Control-Allow-Origin": "*",
  });
  res.end(body);
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

  const wb = await loadWorkbook();
  const ws = getOrCreateMonthSheet(wb, body["日付"]);
  ws.addRow([
    body["注文番号"], body["日付"], body["サイト"], body["商品メモ"], body["商品ID"] || "",
    usd, rate, Math.round(revenueJpy), fee,
    cost === null ? "" : cost, shipping === null ? "" : shipping, packing === null ? "" : packing,
    profit, margin,
  ]);
  await wb.xlsx.writeFile(LEDGER_PATH);

  sendJson(res, 200, { status: "ok", 収益円: Math.round(revenueJpy), 手数料円: fee, 最終利益円: profit, 利益率: margin });
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

  const wb = await loadWorkbook();
  const sheets = dataSheets(wb);
  let updated = null;
  for (const ws of sheets) {
    for (let r = 2; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      if (String(row.getCell(COL.注文番号).value || "") !== String(orderNo)) continue;

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
      row.getCell(COL.収益円).value = Math.round(revenueJpy);
      row.getCell(COL.手数料).value = fee;
      row.getCell(COL.仕入原価).value = cost === null ? "" : cost;
      row.getCell(COL.送料).value = shipping === null ? "" : shipping;
      row.getCell(COL.梱包費).value = packing === null ? "" : packing;
      row.getCell(COL.最終利益).value = profit;
      row.getCell(COL.利益率).value = margin;
      row.commit();
      updated = { 注文番号: orderNo, 最終利益円: profit, 利益率: margin };
    }
  }

  if (!updated) return sendJson(res, 404, { error: "該当する注文番号が見つかりません" });
  await wb.xlsx.writeFile(LEDGER_PATH);
  sendJson(res, 200, { status: "ok", ...updated });
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
    "Content-Disposition": "attachment; filename=\"売上管理表.xlsx\"",
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
  const tmpPath = LEDGER_PATH + ".tmp";
  fs.writeFileSync(tmpPath, buf);
  fs.renameSync(tmpPath, LEDGER_PATH);
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

async function handleImportInventory(req, res) {
  const buf = await readRawBody(req, 30 * 1024 * 1024);
  const wb = new ExcelJS.Workbook();
  try {
    await wb.xlsx.load(buf);
  } catch (e) {
    return sendJson(res, 400, { error: "有効なxlsxファイルではありません" });
  }
  const tmpPath = INVENTORY_PATH + ".tmp";
  fs.writeFileSync(tmpPath, buf);
  fs.renameSync(tmpPath, INVENTORY_PATH);
  sendJson(res, 200, { status: "ok", message: "取り込みが完了しました" });
}

const INV_HEADER_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F4E78" } };
const INV_HEADER_FONT = { bold: true, color: { argb: "FFFFFFFF" } };
const INV_FLAG_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF2CC" } };
const INV_REMOVED_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF2DCDB" } };

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

  const oldWb = await loadInventoryWorkbook();
  const oldWs = oldWb.getWorksheet("在庫管理表") || oldWb.worksheets[0];
  const saved = new Map();
  let maxId = 0;
  oldWs.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const pid = row.getCell(1).value;
    if (!pid) return;
    const idNum = parseInt(String(pid).slice(1), 10);
    if (Number.isFinite(idNum)) maxId = Math.max(maxId, idNum);
    const fullRow = [];
    for (let c = 1; c <= INV_HEADERS.length; c++) fullRow.push(row.getCell(c).value);
    const key = (fullRow[1] || "") + " || " + (fullRow[2] || "");
    saved.set(key, {
      商品ID: pid,
      fullRow,
      wasRemoved: String(fullRow[18] || "").includes("見当たりません"),
    });
  });

  const groups = new Map();
  records.forEach((r) => {
    const key = (r["Title"] || "") + " || " + (r["Variation details"] || "");
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(r);
  });

  const anomalous = [];
  const mainKeys = [];
  for (const [key, items] of groups) {
    if (items.length > 3) anomalous.push(...items);
    else mainKeys.push(key);
  }

  const today = new Date().toISOString().slice(0, 10);
  const rowsOut = [];
  const newProducts = [];
  let nextId = maxId + 1;
  const newKeySet = new Set(mainKeys);

  for (const key of mainKeys) {
    const items = groups.get(key);
    const first = items[0];
    const qtys = new Set(items.map((it) => it["Available quantity"]));
    const bySite = {};
    items.forEach((it) => {
      const site = it["Listing site"];
      if (["US", "UK", "AU"].includes(site) && !bySite[site]) bySite[site] = it;
    });
    const siteCount = Object.keys(bySite).length;

    let pid;
    let purchasePrice = null, purchaseDate = null, purchaseFrom = null, note = "";
    if (saved.has(key)) {
      const s = saved.get(key);
      pid = s.商品ID;
      purchasePrice = s.fullRow[15]; purchaseDate = s.fullRow[16]; purchaseFrom = s.fullRow[17];
      note = s.wasRemoved ? "" : (s.fullRow[18] || "");
    } else {
      pid = "P" + String(nextId).padStart(4, "0");
      nextId++;
      newProducts.push([pid, first["Title"]]);
    }

    rowsOut.push({
      row: [
        pid, first["Title"] || "", first["Variation details"] || null, intOrNullCsv(first["Available quantity"]),
        siteCount, qtys.size > 1 ? "要確認" : "",
        bySite.US ? numOrNullCsv(bySite.US["Item number"]) : null, bySite.US ? numOrNullCsv(bySite.US["Current price"]) : null, bySite.US ? intOrNullCsv(bySite.US["Sold quantity"]) : null,
        bySite.UK ? numOrNullCsv(bySite.UK["Item number"]) : null, bySite.UK ? numOrNullCsv(bySite.UK["Current price"]) : null, bySite.UK ? intOrNullCsv(bySite.UK["Sold quantity"]) : null,
        bySite.AU ? numOrNullCsv(bySite.AU["Item number"]) : null, bySite.AU ? numOrNullCsv(bySite.AU["Current price"]) : null, bySite.AU ? intOrNullCsv(bySite.AU["Sold quantity"]) : null,
        purchasePrice, purchaseDate, purchaseFrom, note,
      ],
      status: "current",
      flag: siteCount < 3 || qtys.size > 1,
    });
  }

  let removedNewCount = 0;
  let removedStillCount = 0;
  for (const [key, s] of saved) {
    if (newKeySet.has(key)) continue;
    let note = s.fullRow[18] || "";
    if (s.wasRemoved) {
      removedStillCount++;
    } else {
      note = (note ? note + " / " : "") + `${today}時点のeBay出品CSVに見当たりません(削除/売り切れの可能性)`;
      removedNewCount++;
    }
    const carried = s.fullRow.slice(0, 15);
    rowsOut.push({ row: carried.concat([s.fullRow[15], s.fullRow[16], s.fullRow[17], note]), status: "removed" });
  }

  rowsOut.sort((a, b) => {
    if ((a.status === "removed") !== (b.status === "removed")) return a.status === "removed" ? 1 : -1;
    return String(a.row[0]).localeCompare(String(b.row[0]));
  });

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("在庫管理表");
  ws.addRow(INV_HEADERS);
  for (let c = 1; c <= INV_HEADERS.length; c++) {
    const cell = ws.getRow(1).getCell(c);
    cell.fill = INV_HEADER_FILL;
    cell.font = INV_HEADER_FONT;
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  }
  ws.getRow(1).height = 30;
  ws.views = [{ state: "frozen", ySplit: 1 }];

  rowsOut.forEach((r) => {
    const excelRow = ws.addRow(r.row);
    if (r.status === "removed" || r.flag) {
      const fill = r.status === "removed" ? INV_REMOVED_FILL : INV_FLAG_FILL;
      for (let c = 1; c <= INV_HEADERS.length; c++) excelRow.getCell(c).fill = fill;
    }
  });
  const widths = [10, 40, 30, 12, 10, 12, 14, 12, 12, 14, 12, 12, 14, 12, 12, 14, 12, 16, 30];
  widths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });

  const ws2 = wb.addWorksheet("要確認(重複出品)");
  ws2.addRow(["名寄せキー", "出品ID", "商品名", "バリエーション詳細", "SKU", "在庫数", "出品国", "現在価格", "通貨"]);
  for (let c = 1; c <= 9; c++) {
    ws2.getRow(1).getCell(c).fill = INV_HEADER_FILL;
    ws2.getRow(1).getCell(c).font = INV_HEADER_FONT;
  }
  ws2.views = [{ state: "frozen", ySplit: 1 }];
  anomalous.forEach((it) => {
    ws2.addRow([
      (it["Title"] || "") + " || " + (it["Variation details"] || ""),
      numOrNullCsv(it["Item number"]), it["Title"], it["Variation details"], it["Custom label (SKU)"],
      intOrNullCsv(it["Available quantity"]), it["Listing site"], numOrNullCsv(it["Current price"]), it["Currency"],
    ]);
  });
  [50, 14, 40, 30, 20, 10, 10, 12, 10].forEach((w, i) => { ws2.getColumn(i + 1).width = w; });

  const ws3 = wb.addWorksheet("この表について");
  ws3.getCell("A1").value = "在庫管理表について";
  ws3.getCell("A1").font = { bold: true, size: 13 };
  const notes = [
    `・eBayの『All active listings report』(${today}時点)を元に、サーバー上で自動更新しました。`,
    "・商品IDは前回から引き継いでいます(売上管理表の商品IDとの対応を保つため)。新しく出品された商品には新しいIDを振っています。",
    "・黄色でハイライトした行は「3カ国揃っていない」または「在庫数が国ごとに違う」商品です。",
    "・赤色でハイライトした行は、出品CSVに見当たらなかった商品です。削除されたか、売り切れて自動終了した可能性があります。",
    "・「要確認(重複出品)」シートには、同じ名前の出品が3件を超えて存在した商品を出品ID単位でそのまま残しています。",
    "・仕入価格・仕入日・仕入先・備考は、前回入力済みだった内容をそのまま引き継いでいます。",
  ];
  notes.forEach((n, i) => {
    const cell = ws3.getCell(`A${i + 3}`);
    cell.value = n;
    cell.font = { italic: true, color: { argb: "FF808080" }, size: 9 };
  });
  ws3.getColumn(1).width = 110;

  const tmpPath = INVENTORY_PATH + ".tmp";
  await wb.xlsx.writeFile(tmpPath);
  fs.renameSync(tmpPath, INVENTORY_PATH);

  sendJson(res, 200, {
    status: "ok",
    total: rowsOut.length,
    new: newProducts.length,
    removedNew: removedNewCount,
    removedStill: removedStillCount,
    anomalous: anomalous.length,
  });
}

function sendHtml(res, html) {
  res.writeHead(200, { "Content-Type": "text/html; charset=utf-8" });
  res.end(html);
}

const DASHBOARD_PAGE = `<!doctype html>
<html lang="ja"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>物販事業ツール(サーバー版)</title>
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

  .table-scroll { overflow-x: auto; min-width: 0; max-height: 620px; overflow-y: auto; margin-top: 14px; }
  table { width: 100%; border-collapse: collapse; font-size: 12.5px; }
  thead th { text-align: left; font-size: 11.5px; color: var(--ink-muted); font-weight: 600; padding: 8px 10px; border-bottom: 1px solid var(--border); white-space: nowrap; position: sticky; top: 0; background: var(--surface); z-index: 2; }
  thead th.num, td.num { text-align: right; font-variant-numeric: tabular-nums; }
  tbody td { padding: 8px 10px; border-bottom: 1px solid var(--border); white-space: nowrap; }
  tbody tr:hover { background: var(--surface-2); }
  .sticky-col { position: sticky; left: 0; background: var(--surface); z-index: 1; box-shadow: 2px 0 4px -2px var(--border); }
  tbody tr:hover td.sticky-col { background: var(--surface-2); }
  .site-chip { display: inline-block; padding: 2px 8px; border-radius: 100px; background: var(--accent-wash); color: var(--series-rev); font-size: 11.5px; font-weight: 600; }
  .profit-cell { color: var(--good); font-weight: 600; }
  .profit-cell.bad { color: var(--series-cost); }
  .stock-zero { color: var(--series-cost); font-weight: 600; }

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
  input[type=file] { font-size: 12.5px; color: var(--ink-2); }
</style></head>
<body>
<div class="wrap">
  <div class="hdr">
    <div>
      <h1>物販事業ツール(サーバー版)</h1>
      <p class="sub">株式会社アゲイト — eBay 注文・在庫・売上。注文タブは書き換えると即座にサーバーに保存されます</p>
    </div>
    <div class="auth-row">
      <input id="token" type="password" placeholder="アクセストークン">
      <button class="btn" id="saveToken">トークンを記憶</button>
      <span id="status"></span>
    </div>
  </div>

  <div class="tabnav">
    <button class="tabbtn active" data-tab="sales">売上分析</button>
    <button class="tabbtn" data-tab="inventory">在庫</button>
    <button class="tabbtn" data-tab="orders">注文</button>
  </div>

  <div class="tabpanel active" id="tab-sales">
    <div class="kpis" id="kpis"></div>

    <div class="panel">
      <h2>月次データ</h2>
      <p class="desc">サーバー上の最新データから自動集計(注文の日付ごとに月単位でまとめています)</p>
      <div class="panel-body table-scroll">
        <table>
          <thead><tr><th>月</th><th class="num">件数</th><th class="num">収益(円)</th><th class="num">仕入(円)</th><th class="num">最終利益(円)</th><th class="num">利益率</th></tr></thead>
          <tbody id="monthly-body"></tbody>
        </table>
      </div>
    </div>

    <div class="row2">
      <div class="panel">
        <h2>月次推移(収益・仕入・利益)</h2>
        <div class="panel-body">
          <div class="legend">
            <span class="legend-item"><span class="legend-swatch" style="background:var(--series-rev)"></span>収益</span>
            <span class="legend-item"><span class="legend-swatch" style="background:var(--series-cost)"></span>仕入価格</span>
            <span class="legend-item"><span class="legend-swatch" style="background:var(--series-prof)"></span>最終利益</span>
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
      <h2>サイト別 収益</h2>
      <p class="desc">出品先サイトごとの収益合計(円)</p>
      <div class="panel-body">
        <svg id="chart-site" viewBox="0 0 780 220" width="100%"></svg>
      </div>
    </div>

    <div class="panel">
      <h2>利益TOP10の注文</h2>
      <p class="desc">最終利益が大きかった注文(サーバー上の最新データ)</p>
      <div class="panel-body table-scroll">
        <table>
          <thead><tr><th>日付</th><th>サイト</th><th>商品メモ</th><th class="num">収益(円)</th><th class="num">最終利益(円)</th><th class="num">利益率</th></tr></thead>
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
        <div class="paste-parse">
          <div class="browser-toolbar">
            <input type="file" id="inv-csv-file" accept=".csv">
            <button class="btn" id="inv-csv-upload">CSVで在庫を更新</button>
            <span id="inv-csv-status" class="hint"></span>
          </div>
        </div>
        <div class="browser-toolbar">
          <input type="text" class="search-input" id="inv-q" placeholder="商品名・商品IDで検索…">
          <span class="result-count" id="inv-count"></span>
        </div>
        <div class="scroll-mirror-top" id="inv-table-mirror"><div class="scroll-mirror-top-inner"></div></div>
        <div class="table-scroll" id="inv-table-scroll">
          <table>
            <thead><tr id="inv-thead"></tr></thead>
            <tbody id="inv-tbody"></tbody>
          </table>
        </div>
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
            <label style="grid-column:span 3;">商品メモ<input id="ne-note" type="text"></label>
          </div>
          <div class="browser-toolbar" id="ne-submit-row" style="display:none;">
            <button class="btn" id="ne-submit-btn">この内容で登録する</button>
            <span id="ne-status" class="hint"></span>
          </div>
        </div>
      </div>
    </div>

    <div class="panel">
      <h2>注文一覧</h2>
      <p class="desc">注文番号・商品メモ・サイトで検索できます。<b>収益USD・ドル円レート・仕入原価・送料・梱包費は直接書き換えられます</b>(最終利益はその場で再計算・保存されます)。</p>
      <div class="panel-body">
        <div class="browser-toolbar">
          <input type="text" class="search-input" id="ord-q" placeholder="注文番号・商品メモ・サイトで検索…">
          <span class="result-count" id="ord-count"></span>
        </div>
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
</div>
<div class="viz-tooltip" id="tooltip"></div>
<script>
const ORD_HEADERS = ["注文番号","日付","サイト","商品メモ","商品ID","収益USD","ドル円レート","収益円","手数料(円)","仕入原価(円)","送料(円)","梱包費(円)","最終利益(円)","利益率"];
const ORD_EDITABLE = ["収益USD","ドル円レート","仕入原価(円)","送料(円)","梱包費(円)"];
const ORD_NUM_COLS = ["収益USD","ドル円レート","収益円","手数料(円)","仕入原価(円)","送料(円)","梱包費(円)","最終利益(円)","利益率"];
const ORD_FIELD_KEY = { "収益USD": "収益USD", "ドル円レート": "ドル円レート", "仕入原価(円)": "仕入原価円", "送料(円)": "送料円", "梱包費(円)": "梱包費円" };
const INV_NUM_COLS = ["在庫数(現物)","出品国数","US価格(USD)","US累計売却数","UK価格(GBP)","UK累計売却数","AU価格(AUD)","AU累計売却数","仕入価格(円)"];
let orderRows = [];
let invRows = [];
let invHeaders = [];

function getToken() { return localStorage.getItem("agate_token") || ""; }
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
    const [ordRes, invRes] = await Promise.all([
      fetch("/api/orders", { headers: { Authorization: "Bearer " + token } }),
      fetch("/api/inventory", { headers: { Authorization: "Bearer " + token } }),
    ]);
    if (!ordRes.ok) { statusEl.textContent = "エラー: " + ordRes.status + "(トークンを確認してください)"; return; }
    const ordData = await ordRes.json();
    orderRows = ordData.rows;
    if (invRes.ok) {
      const invData = await invRes.json();
      invRows = invData.rows;
      invHeaders = invData.headers;
    }
    renderKpis();
    renderSalesTab();
    renderOrders();
    renderInventory();
    statusEl.textContent = "";
  } catch (e) {
    statusEl.textContent = "通信エラー: " + e.message;
  }
}

function renderKpis() {
  let revenue = 0, cost = 0, profit = 0, profitKnown = 0;
  orderRows.forEach(r => {
    revenue += Number(r[7]) || 0;
    cost += Number(r[9]) || 0;
    if (r[12] !== "" && r[12] !== null && r[12] !== undefined) { profit += Number(r[12]); profitKnown++; }
  });
  const margin = revenue !== 0 ? profit / revenue : 0;
  const tiles = [
    { label: "総収益", value: "¥" + fmt(Math.round(revenue)) },
    { label: "総仕入価格", value: "¥" + fmt(Math.round(cost)) },
    { label: "総最終利益", value: "¥" + fmt(Math.round(profit)), cls: profit >= 0 ? "good" : "bad" },
    { label: "利益率(判明分)", value: (margin * 100).toFixed(1) + "%", cls: margin >= 0 ? "good" : "bad" },
    { label: "総注文数", value: orderRows.length.toLocaleString("ja-JP") + " 件" },
  ];
  document.getElementById("kpis").innerHTML = tiles.map(t =>
    '<div class="kpi"><div class="label">' + t.label + '</div><div class="value' + (t.cls ? " " + t.cls : "") + '">' + t.value + '</div></div>'
  ).join("");

  const top10 = orderRows.filter(r => r[12] !== "" && r[12] !== null && r[12] !== undefined)
    .slice().sort((a, b) => Number(b[12]) - Number(a[12])).slice(0, 10);
  document.getElementById("top10-body").innerHTML = top10.map(r =>
    "<tr><td>" + (r[1] || "") + "</td><td><span class='site-chip'>" + (r[2] || "不明") + "</span></td><td>" + (r[3] || "") +
    "</td><td class='num'>" + fmt(r[7]) + "</td><td class='num profit-cell" + (Number(r[12]) < 0 ? " bad" : "") + "'>" + fmt(r[12]) +
    "</td><td class='num'>" + fmt(r[13], true) + "</td></tr>"
  ).join("");
}

function computeMonthly() {
  const map = new Map();
  orderRows.forEach((r) => {
    const d = String(r[1] || "");
    if (d.length < 7) return;
    const month = d.slice(0, 7);
    if (!map.has(month)) map.set(month, { revenue: 0, cost: 0, profit: 0, count: 0 });
    const m = map.get(month);
    m.revenue += Number(r[7]) || 0;
    m.cost += Number(r[9]) || 0;
    if (r[12] !== "" && r[12] !== null && r[12] !== undefined) m.profit += Number(r[12]);
    m.count++;
  });
  return Array.from(map.keys()).sort().map((k) => {
    const v = map.get(k);
    return { month: k, revenue: v.revenue, cost: v.cost, profit: v.profit, count: v.count, margin: v.revenue ? v.profit / v.revenue : 0 };
  });
}

function renderMonthlyTable(monthly) {
  const newestFirst = monthly.slice().reverse();
  document.getElementById("monthly-body").innerHTML = newestFirst.map((m) =>
    "<tr><td>" + m.month + "</td><td class='num'>" + m.count + "</td><td class='num'>" + fmt(Math.round(m.revenue)) +
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
  const labels = { revenue: "収益", cost: "仕入", profit: "最終利益" };

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
  const W = 480, H = 300, padL = 45, padR = 10, padT = 10, padB = 30;
  const innerW = W - padL - padR, innerH = H - padT - padB;
  const maxAbs = Math.max(0.05, ...monthly.map((m) => Math.abs(m.margin)));
  const n = monthly.length;
  const stepX = n > 1 ? innerW / (n - 1) : 0;
  const yFor = (v) => padT + innerH / 2 - (v / maxAbs) * (innerH / 2);

  for (let i = 0; i <= 4; i++) {
    const y = padT + (innerH * i) / 4;
    svg.appendChild(svgNode("line", { x1: padL, y1: y, x2: padL + innerW, y2: y, class: "grid-line" }));
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
    bySite.set(site, (bySite.get(site) || 0) + (Number(r[7]) || 0));
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
    rect.addEventListener("mousemove", (evt) => showTooltip(evt, "<b>" + site + "</b><br>収益: ¥" + fmt(Math.round(val))));
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
  inner.style.width = table.scrollWidth + "px";
  let syncing = false;
  mirror.onscroll = () => { if (syncing) return; syncing = true; scrollEl.scrollLeft = mirror.scrollLeft; syncing = false; };
  scrollEl.onscroll = () => { if (syncing) return; syncing = true; mirror.scrollLeft = scrollEl.scrollLeft; syncing = false; };
}

const ORD_HIDDEN = ["商品ID"];

function renderOrders() {
  const displayOrder = ORD_HEADERS.map((h, i) => i).filter((i) => !ORD_HIDDEN.includes(ORD_HEADERS[i]));
  const thead = document.getElementById("ord-thead");
  thead.innerHTML = displayOrder.map((i, pos) => '<th class="' + (ORD_NUM_COLS.includes(ORD_HEADERS[i]) ? "num" : "") + (pos === 0 ? " sticky-col" : "") + '">' + ORD_HEADERS[i] + '</th>').join("");
  const q = document.getElementById("ord-q").value.trim().toLowerCase();
  const tbody = document.getElementById("ord-tbody");
  tbody.innerHTML = "";
  let shown = 0;
  const sortedOrders = orderRows.slice().sort((a, b) => String(b[1] || "").localeCompare(String(a[1] || "")));
  sortedOrders.forEach(row => {
    const orderNo = row[0];
    const searchable = ((row[0]||"") + " " + (row[2]||"") + " " + (row[3]||"")).toLowerCase();
    if (q && searchable.indexOf(q) === -1) return;
    shown++;
    const tr = document.createElement("tr");
    displayOrder.forEach((i, pos) => {
      const h = ORD_HEADERS[i];
      const td = document.createElement("td");
      if (pos === 0) td.className = "sticky-col";
      if (ORD_EDITABLE.includes(h)) {
        td.className = (td.className ? td.className + " " : "") + "num";
        const inp = document.createElement("input");
        inp.type = "number"; inp.step = "any";
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
  setupScrollMirror("ord-table-mirror", "ord-table-scroll");
}

async function saveOrderField(tr, orderNo, header, value) {
  tr.className = "saving";
  const token = getToken();
  const body = { 注文番号: orderNo };
  body[ORD_FIELD_KEY[header]] = value === "" ? "" : Number(value);
  try {
    const r = await fetch("/api/orders", {
      method: "PATCH",
      headers: { "Content-Type": "application/json", Authorization: "Bearer " + token },
      body: JSON.stringify(body),
    });
    if (!r.ok) { tr.className = "error"; return; }
    tr.className = "saved";
    setTimeout(() => { tr.className = ""; }, 1500);
  } catch (e) {
    tr.className = "error";
  }
}

const INV_HIDDEN = ["UK_出品ID", "UK価格(GBP)", "AU_出品ID", "AU価格(AUD)", "在庫数不一致", "バリエーション詳細"];

function renderInventory() {
  if (!invHeaders.length) return;
  const displayOrder = invHeaders.map((h, i) => i).filter((i) => !INV_HIDDEN.includes(invHeaders[i]));
  const thead = document.getElementById("inv-thead");
  thead.innerHTML = displayOrder.map((i, pos) => '<th class="' + (INV_NUM_COLS.includes(invHeaders[i]) ? "num" : "") + (pos === 0 ? " sticky-col" : "") + '">' + invHeaders[i] + '</th>').join("");
  const q = document.getElementById("inv-q").value.trim().toLowerCase();
  const tbody = document.getElementById("inv-tbody");
  tbody.innerHTML = "";
  let shown = 0;
  const stockIdx = invHeaders.indexOf("在庫数(現物)");
  const sortedInv = invRows.slice().sort((a, b) => {
    const na = parseInt(String(a[0] || "").slice(1), 10) || 0;
    const nb = parseInt(String(b[0] || "").slice(1), 10) || 0;
    return nb - na;
  });
  sortedInv.forEach(row => {
    const searchable = ((row[0]||"") + " " + (row[1]||"")).toLowerCase();
    if (q && searchable.indexOf(q) === -1) return;
    shown++;
    const tr = document.createElement("tr");
    displayOrder.forEach((i, pos) => {
      const h = invHeaders[i];
      const td = document.createElement("td");
      if (pos === 0) td.className = "sticky-col";
      const isNum = INV_NUM_COLS.includes(h);
      td.className = (td.className ? td.className + " " : "") + (isNum ? "num" : "");
      if (i === stockIdx && (row[i] === 0 || row[i] === null || row[i] === undefined)) td.className += " stock-zero";
      td.textContent = row[i] === null || row[i] === undefined ? "" : (isNum && typeof row[i] === "number" ? fmt(row[i]) : row[i]);
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  document.getElementById("inv-count").textContent = shown.toLocaleString("ja-JP") + " / " + invRows.length.toLocaleString("ja-JP") + " 件";
  setupScrollMirror("inv-table-mirror", "inv-table-scroll");
}

document.getElementById("ord-q").addEventListener("input", renderOrders);
document.getElementById("inv-q").addEventListener("input", renderInventory);
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

const MONTH_MAP = { Jan: 1, Feb: 2, Mar: 3, Apr: 4, May: 5, Jun: 6, Jul: 7, Aug: 8, Sep: 9, Oct: 10, Nov: 11, Dec: 12 };

function parseOrderText(text) {
  const lines = text.split(/\\r?\\n/).map((l) => l.trim());
  const result = { orderNo: "", date: "", site: "", usd: "", note: "" };

  for (const line of lines) {
    const m = /^(\\d{1,3}-\\d{4,7}-\\d{4,7})$/.exec(line);
    if (m) { result.orderNo = m[1]; break; }
  }

  const saleIdx = lines.indexOf("販売");
  if (saleIdx !== -1) {
    for (let i = saleIdx + 1; i < lines.length && i < saleIdx + 4; i++) {
      const dm = /^([A-Za-z]{3})[a-z]*\\s+(\\d{1,2}),\\s*(\\d{4})/.exec(lines[i]);
      if (dm) {
        const key = dm[1][0].toUpperCase() + dm[1].slice(1, 3).toLowerCase();
        const mm = MONTH_MAP[key];
        if (mm) {
          result.date = dm[3] + "-" + String(mm).padStart(2, "0") + "-" + String(dm[2]).padStart(2, "0");
          break;
        }
      }
    }
  }

  for (const line of lines) {
    const sm = /^([A-Z]{2})\\s*[£$]/.exec(line);
    if (sm) { result.site = sm[1]; break; }
  }

  const revIdx = lines.indexOf("注文の収益");
  if (revIdx !== -1) {
    for (let i = revIdx + 1; i < lines.length && i < revIdx + 3; i++) {
      const um = /US\\s*\\$\\s*([\\d.]+)/.exec(lines[i]);
      if (um) { result.usd = um[1]; break; }
    }
  }

  const titles = [];
  for (let i = 0; i < lines.length; i++) {
    if (/^商品ID/.test(lines[i]) && i >= 2) {
      const t = lines[i - 2];
      if (t && !titles.includes(t)) titles.push(t);
    }
  }
  result.note = titles.join(" / ");

  return result;
}

document.getElementById("ne-parse-btn").addEventListener("click", () => {
  const parsed = parseOrderText(document.getElementById("ne-paste").value);
  document.getElementById("ne-order").value = parsed.orderNo;
  document.getElementById("ne-date").value = parsed.date;
  document.getElementById("ne-site").value = parsed.site;
  document.getElementById("ne-usd").value = parsed.usd;
  document.getElementById("ne-note").value = parsed.note;
  document.getElementById("ne-review").style.display = "grid";
  document.getElementById("ne-submit-row").style.display = "flex";
  document.getElementById("ne-status").textContent = "内容を確認してから登録してください(空欄は自分で埋めてください)";
  document.getElementById("ne-status").className = "hint";
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
  };
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
    statusEl.textContent = "登録しました(収益円: " + data.収益円 + ")";
    statusEl.className = "hint ok";
    document.getElementById("ne-paste").value = "";
    document.getElementById("ne-review").style.display = "none";
    document.getElementById("ne-submit-row").style.display = "none";
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
      "Access-Control-Allow-Methods": "GET, POST, PATCH, OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type, Authorization",
    });
    return res.end();
  }

  if (req.method === "GET" && req.url === "/") {
    res.writeHead(200, { "Content-Type": "text/plain; charset=utf-8" });
    return res.end("agate-tool-server: OK");
  }
  if (req.method === "GET" && (req.url === "/orders" || req.url === "/dashboard")) {
    return sendHtml(res, DASHBOARD_PAGE);
  }

  const protectedRoutes = ["/api/orders", "/api/inventory", "/download/売上管理表.xlsx", "/api/import", "/api/import/inventory", "/api/inventory/rebuild"];
  if (protectedRoutes.includes(req.url) && !isAuthorized(req)) {
    return sendJson(res, 401, { error: "認証に失敗しました(トークンを確認してください)" });
  }

  try {
    if (req.method === "POST" && req.url === "/api/orders") return await handleAddOrder(req, res);
    if (req.method === "PATCH" && req.url === "/api/orders") return await handlePatchOrder(req, res);
    if (req.method === "GET" && req.url === "/api/orders") return await handleListOrders(req, res);
    if (req.method === "GET" && req.url === "/download/売上管理表.xlsx") return await handleDownload(req, res);
    if (req.method === "POST" && req.url === "/api/import") return await handleImport(req, res);
    if (req.method === "GET" && req.url === "/api/inventory") return await handleListInventory(req, res);
    if (req.method === "POST" && req.url === "/api/import/inventory") return await handleImportInventory(req, res);
    if (req.method === "POST" && req.url === "/api/inventory/rebuild") return await handleRebuildInventory(req, res);
  } catch (e) {
    console.error(e);
    return sendJson(res, 500, { error: "サーバー内部でエラーが発生しました" });
  }

  sendJson(res, 404, { error: "not found" });
});

server.listen(PORT, "127.0.0.1", () => {
  console.log(`agate-tool-server listening on 127.0.0.1:${PORT}`);
});
