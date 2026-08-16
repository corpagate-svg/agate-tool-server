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
    --good:        #006300;
    --series-rev:  #2a78d6;
    --series-cost: #eb6834;
    --accent-wash: rgba(42,120,214,0.10);
  }
  @media (prefers-color-scheme: dark) {
    :root:not([data-theme="light"]) {
      color-scheme: dark;
      --page: #0d0d0d; --surface: #1a1a19; --surface-2: #232322;
      --ink: #ffffff; --ink-2: #c3c2b7; --ink-muted: #898781;
      --border: rgba(255,255,255,0.10); --good: #0ca30c;
      --series-rev: #3987e5; --series-cost: #d95926;
      --accent-wash: rgba(57,135,229,0.14);
    }
  }
  * { box-sizing: border-box; }
  body {
    margin: 0; background: var(--page); color: var(--ink);
    font-family: -apple-system, BlinkMacSystemFont, "Hiragino Sans", "Hiragino Kaku Gothic ProN", "Yu Gothic UI", "Segoe UI", Roboto, sans-serif;
    -webkit-font-smoothing: antialiased;
  }
  .wrap { max-width: 1180px; margin: 0 auto; padding: 32px 24px 60px; display: flex; flex-direction: column; gap: 20px; }
  .hdr { display: flex; justify-content: space-between; align-items: flex-end; flex-wrap: wrap; gap: 12px; border-bottom: 1px solid var(--border); padding-bottom: 18px; }
  .hdr h1 { margin: 0 0 6px; font-size: 24px; font-weight: 700; letter-spacing: -0.01em; }
  .hdr .sub { margin: 0; color: var(--ink-2); font-size: 13px; }

  .auth-row { display: flex; gap: 8px; align-items: center; flex-wrap: wrap; }
  .auth-row input[type=password] { padding: 8px 10px; border-radius: 7px; border: 1px solid var(--border); background: var(--surface-2); color: var(--ink); font-size: 13px; min-width: 260px; }
  .btn { font: inherit; font-size: 12.5px; font-weight: 600; color: var(--ink); background: var(--surface-2); border: 1px solid var(--border); border-radius: 7px; padding: 7px 14px; cursor: pointer; }
  .btn:hover { background: var(--accent-wash); }
  #status { font-size: 12.5px; color: var(--ink-muted); }

  .panel { background: var(--surface); border: 1px solid var(--border); border-radius: 12px; padding: 22px 24px 20px; }
  .panel h2 { margin: 0; font-size: 15px; font-weight: 700; }
  .panel .desc { margin: 3px 0 0; font-size: 12.5px; color: var(--ink-muted); }
  .panel-body { margin-top: 16px; }

  .browser-toolbar { display: flex; align-items: center; gap: 12px; flex-wrap: wrap; }
  .search-input { font: inherit; font-size: 13px; color: var(--ink); background: var(--surface-2); border: 1px solid var(--border); border-radius: 8px; padding: 9px 12px; flex: 1; min-width: 220px; }
  .search-input:focus { outline: 2px solid var(--series-rev); outline-offset: 1px; background: var(--surface); }
  .result-count { font-size: 12.5px; color: var(--ink-muted); white-space: nowrap; }

  .table-scroll { overflow-x: auto; min-width: 0; max-height: 640px; overflow-y: auto; margin-top: 14px; }
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

  td input { font: inherit; font-variant-numeric: tabular-nums; font-size: 12.5px; color: var(--ink); background: var(--surface-2); border: 1px solid var(--border); border-radius: 5px; padding: 5px 6px; width: 78px; text-align: right; }
  td input:focus { outline: 2px solid var(--series-rev); outline-offset: 1px; background: var(--surface); }
  tr.saving td { background: #fff7e0 !important; }
  tr.saved td { background: #e9f7ec !important; }
  tr.error td { background: #fde8e8 !important; }
</style></head>
<body>
<div class="wrap">
  <div class="hdr">
    <div>
      <h1>物販事業ツール(サーバー版)</h1>
      <p class="sub">株式会社アゲイト — 注文一覧。仕入原価・送料・梱包費を書き換えると、その場でサーバーに保存されます</p>
    </div>
    <div class="auth-row">
      <input id="token" type="password" placeholder="アクセストークン">
      <button class="btn" id="saveToken">トークンを記憶</button>
      <span id="status"></span>
    </div>
  </div>

  <div class="panel">
    <h2>注文一覧</h2>
    <p class="desc">注文番号・商品メモ・サイトで検索できます。<b>収益USD・ドル円レート・仕入原価・送料・梱包費は直接書き換えられます</b>(最終利益はその場で再計算されます)。</p>
    <div class="panel-body">
      <div class="browser-toolbar">
        <input type="text" class="search-input" id="q" placeholder="注文番号・商品メモ・サイトで検索…">
        <span class="result-count" id="count"></span>
      </div>
      <div class="table-scroll">
        <table>
          <thead><tr id="thead"></tr></thead>
          <tbody id="tbody"></tbody>
        </table>
      </div>
    </div>
  </div>
</div>
<script>
const HEADERS_META = ["注文番号","日付","サイト","商品メモ","商品ID","収益USD","ドル円レート","収益円","手数料(円)","仕入原価(円)","送料(円)","梱包費(円)","最終利益(円)","利益率"];
const EDITABLE = ["収益USD","ドル円レート","仕入原価(円)","送料(円)","梱包費(円)"];
const NUM_COLS = ["収益USD","ドル円レート","収益円","手数料(円)","仕入原価(円)","送料(円)","梱包費(円)","最終利益(円)","利益率"];
const FIELD_KEY = { "収益USD": "収益USD", "ドル円レート": "ドル円レート", "仕入原価(円)": "仕入原価円", "送料(円)": "送料円", "梱包費(円)": "梱包費円" };
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
  thead.innerHTML = HEADERS_META.map((h, i) => '<th class="' + (NUM_COLS.includes(h) ? "num" : "") + (i === 0 ? " sticky-col" : "") + '">' + h + '</th>').join("");
  const q = document.getElementById("q").value.trim().toLowerCase();
  const tbody = document.getElementById("tbody");
  tbody.innerHTML = "";
  let shown = 0;
  allRows.forEach(row => {
    const orderNo = row[0];
    const searchable = ((row[0]||"") + " " + (row[2]||"") + " " + (row[3]||"")).toLowerCase();
    if (q && searchable.indexOf(q) === -1) return;
    shown++;
    const tr = document.createElement("tr");
    HEADERS_META.forEach((h, i) => {
      const td = document.createElement("td");
      if (i === 0) td.className = "sticky-col";
      if (EDITABLE.includes(h)) {
        td.className = (td.className ? td.className + " " : "") + "num";
        const inp = document.createElement("input");
        inp.type = "number"; inp.step = "any";
        inp.value = row[i] === null || row[i] === undefined ? "" : row[i];
        inp.addEventListener("change", () => saveField(tr, orderNo, h, inp.value));
        td.appendChild(inp);
      } else if (h === "サイト") {
        const chip = document.createElement("span"); chip.className = "site-chip"; chip.textContent = row[i] || "不明"; td.appendChild(chip);
      } else if (h === "最終利益(円)") {
        td.className = (td.className ? td.className + " " : "") + "num profit-cell" + (Number(row[i]) < 0 ? " bad" : "");
        td.textContent = fmt(row[i], h);
      } else {
        td.className = (td.className ? td.className + " " : "") + (NUM_COLS.includes(h) ? "num" : "");
        td.textContent = fmt(row[i], h);
      }
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  document.getElementById("count").textContent = shown.toLocaleString("ja-JP") + " / " + allRows.length.toLocaleString("ja-JP") + " 件";
}

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
