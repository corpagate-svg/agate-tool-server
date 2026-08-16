const http = require("http");
const crypto = require("crypto");
const path = require("path");
const fs = require("fs");
const ExcelJS = require("exceljs");

const PORT = process.env.PORT || 3000;
const API_TOKEN = process.env.API_TOKEN || "";
const DATA_DIR = path.join(__dirname, "..", "data");
const LEDGER_PATH = path.join(DATA_DIR, "売上管理表.xlsx");

const HEADERS = [
  "注文番号", "日付", "サイト", "商品メモ", "商品ID",
  "収益USD", "ドル円レート", "収益円", "手数料(円)",
  "仕入原価(円)", "送料(円)", "梱包費(円)", "最終利益(円)", "利益率",
];
const COL = { 注文番号: 1, 日付: 2, サイト: 3, 商品メモ: 4, 商品ID: 5, 収益USD: 6, ドル円レート: 7, 収益円: 8, 手数料: 9, 仕入原価: 10, 送料: 11, 梱包費: 12, 最終利益: 13, 利益率: 14 };

if (!API_TOKEN) {
  console.error("API_TOKEN が設定されていません(.env を確認してください)。起動を中止します。");
  process.exit(1);
}

fs.mkdirSync(DATA_DIR, { recursive: true });

function isDataSheet(name) {
  return !name.includes("について");
}

function styleHeaderRow(ws) {
  ws.addRow(HEADERS);
  ws.getRow(1).font = { bold: true };
}

async function loadWorkbook() {
  const wb = new ExcelJS.Workbook();
  if (fs.existsSync(LEDGER_PATH)) {
    await wb.xlsx.readFile(LEDGER_PATH);
    return wb;
  }
  const ws = wb.addWorksheet("記録");
  styleHeaderRow(ws);
  await wb.xlsx.writeFile(LEDGER_PATH);
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
    styleHeaderRow(ws);
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
      const revenueJpy = Number(row.getCell(COL.収益円).value) || 0;
      const fee = Number(row.getCell(COL.手数料).value) || 0;
      const { profit, margin } = recompute(revenueJpy, fee, cost, shipping, packing);

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

async function handleListOrders(req, res) {
  const wb = await loadWorkbook();
  const rows = [];
  for (const ws of dataSheets(wb)) {
    ws.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
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

function sendHtml(res, html) {
  res.writeHead(200, { "Content-Type": "text/html; charset=utf-8" });
  res.end(html);
}

const ORDERS_PAGE = `<!doctype html>
<html lang="ja"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>売上管理表(サーバー版)</title>
<style>
  body { font-family: -apple-system, "Hiragino Sans", "Yu Gothic", sans-serif; margin: 0; background: #f6f7f9; color: #1b1f24; }
  header { position: sticky; top: 0; background: #1f4e78; color: #fff; padding: 14px 18px; z-index: 5; }
  header h1 { font-size: 16px; margin: 0 0 8px; }
  .bar { display: flex; gap: 8px; flex-wrap: wrap; align-items: center; }
  input[type=text], input[type=password] { padding: 6px 8px; border-radius: 6px; border: 1px solid #ccc; font-size: 13px; }
  #status { font-size: 12px; margin-left: auto; opacity: .9; }
  main { padding: 12px; }
  table { border-collapse: collapse; width: 100%; background: #fff; font-size: 12px; }
  th, td { border: 1px solid #e2e5e9; padding: 5px 7px; white-space: nowrap; }
  th { background: #eef1f5; position: sticky; top: 66px; text-align: left; }
  td.num, th.num { text-align: right; }
  td input { width: 80px; text-align: right; border: 1px solid #ccc; border-radius: 4px; padding: 3px 5px; }
  tr.saving td { background: #fff7e0; }
  tr.saved td { background: #e9f7ec; }
  tr.error td { background: #fde8e8; }
  .count { color: #556; font-size: 12px; margin: 6px 0; }
</style></head>
<body>
<header>
  <h1>売上管理表(サーバー版) — 仕入原価・送料・梱包費はここで編集すると即座に保存されます</h1>
  <div class="bar">
    <input id="token" type="password" placeholder="アクセストークン">
    <button id="saveToken">トークンを記憶</button>
    <input id="q" type="text" placeholder="検索(注文番号・商品メモ)">
    <span id="status"></span>
  </div>
</header>
<main>
  <div class="count" id="count"></div>
  <div style="overflow:auto">
  <table>
    <thead><tr id="thead"></tr></thead>
    <tbody id="tbody"></tbody>
  </table>
  </div>
</main>
<script>
const HEADERS_META = ["注文番号","日付","サイト","商品メモ","商品ID","収益USD","ドル円レート","収益円","手数料(円)","仕入原価(円)","送料(円)","梱包費(円)","最終利益(円)","利益率"];
const EDITABLE = ["仕入原価(円)","送料(円)","梱包費(円)"];
const NUM_COLS = ["収益USD","ドル円レート","収益円","手数料(円)","仕入原価(円)","送料(円)","梱包費(円)","最終利益(円)","利益率"];
let allRows = [];

function getToken() { return localStorage.getItem("agate_token") || ""; }
document.getElementById("token").value = getToken();
document.getElementById("saveToken").addEventListener("click", () => {
  localStorage.setItem("agate_token", document.getElementById("token").value.trim());
  load();
});

function fmt(v, key) {
  if (v === null || v === undefined || v === "") return "";
  if (key === "利益率") return (Number(v) * 100).toFixed(1) + "%";
  if (typeof v === "number") return v.toLocaleString("ja-JP", { maximumFractionDigits: 2 });
  return v;
}

async function load() {
  const token = getToken();
  const statusEl = document.getElementById("status");
  if (!token) { statusEl.textContent = "トークンを入力してください"; return; }
  statusEl.textContent = "読み込み中...";
  try {
    const r = await fetch("/api/orders", { headers: { Authorization: "Bearer " + token } });
    if (!r.ok) { statusEl.textContent = "エラー: " + r.status + "(トークンを確認してください)"; return; }
    const data = await r.json();
    allRows = data.rows;
    render();
    statusEl.textContent = "";
  } catch (e) {
    statusEl.textContent = "通信エラー: " + e.message;
  }
}

function render() {
  const thead = document.getElementById("thead");
  thead.innerHTML = HEADERS_META.map(h => '<th class="' + (NUM_COLS.includes(h) ? "num" : "") + '">' + h + '</th>').join("");
  const q = document.getElementById("q").value.trim().toLowerCase();
  const tbody = document.getElementById("tbody");
  tbody.innerHTML = "";
  let shown = 0;
  allRows.forEach(row => {
    const orderNo = row[0];
    const searchable = ((row[0]||"") + " " + (row[3]||"")).toLowerCase();
    if (q && searchable.indexOf(q) === -1) return;
    shown++;
    const tr = document.createElement("tr");
    HEADERS_META.forEach((h, i) => {
      const td = document.createElement("td");
      if (EDITABLE.includes(h)) {
        td.className = "num";
        const inp = document.createElement("input");
        inp.type = "number"; inp.step = "any";
        inp.value = row[i] === null || row[i] === undefined ? "" : row[i];
        inp.addEventListener("change", () => saveField(tr, orderNo, h, inp.value));
        td.appendChild(inp);
      } else {
        td.className = NUM_COLS.includes(h) ? "num" : "";
        td.textContent = fmt(row[i], h);
      }
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  document.getElementById("count").textContent = shown.toLocaleString("ja-JP") + " / " + allRows.length.toLocaleString("ja-JP") + " 件";
}

const FIELD_KEY = { "仕入原価(円)": "仕入原価円", "送料(円)": "送料円", "梱包費(円)": "梱包費円" };

async function saveField(tr, orderNo, header, value) {
  tr.className = "saving";
  const token = getToken();
  const body = { 注文番号: orderNo };
  body[FIELD_KEY[header]] = value === "" ? "" : Number(value);
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

document.getElementById("q").addEventListener("input", render);
if (getToken()) load();
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
  if (req.method === "GET" && req.url === "/orders") {
    return sendHtml(res, ORDERS_PAGE);
  }

  const protectedRoutes = ["/api/orders", "/download/売上管理表.xlsx", "/api/import"];
  if (protectedRoutes.includes(req.url) && !isAuthorized(req)) {
    return sendJson(res, 401, { error: "認証に失敗しました(トークンを確認してください)" });
  }

  try {
    if (req.method === "POST" && req.url === "/api/orders") return await handleAddOrder(req, res);
    if (req.method === "PATCH" && req.url === "/api/orders") return await handlePatchOrder(req, res);
    if (req.method === "GET" && req.url === "/api/orders") return await handleListOrders(req, res);
    if (req.method === "GET" && req.url === "/download/売上管理表.xlsx") return await handleDownload(req, res);
    if (req.method === "POST" && req.url === "/api/import") return await handleImport(req, res);
  } catch (e) {
    console.error(e);
    return sendJson(res, 500, { error: "サーバー内部でエラーが発生しました" });
  }

  sendJson(res, 404, { error: "not found" });
});

server.listen(PORT, "127.0.0.1", () => {
  console.log(`agate-tool-server listening on 127.0.0.1:${PORT}`);
});
