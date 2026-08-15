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

if (!API_TOKEN) {
  console.error("API_TOKEN が設定されていません(.env を確認してください)。起動を中止します。");
  process.exit(1);
}

fs.mkdirSync(DATA_DIR, { recursive: true });

async function loadWorkbook() {
  const wb = new ExcelJS.Workbook();
  if (fs.existsSync(LEDGER_PATH)) {
    await wb.xlsx.readFile(LEDGER_PATH);
    return wb;
  }
  const ws = wb.addWorksheet("売上管理表");
  ws.addRow(HEADERS);
  ws.getRow(1).font = { bold: true };
  await wb.xlsx.writeFile(LEDGER_PATH);
  return wb;
}

function getSheet(wb) {
  return wb.getWorksheet("売上管理表") || wb.worksheets[0];
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

function readBody(req, maxBytes) {
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
    req.on("end", () => resolve(Buffer.concat(chunks).toString("utf8")));
    req.on("error", reject);
  });
}

async function handleAddOrder(req, res) {
  let body;
  try {
    body = JSON.parse(await readBody(req, 1024 * 1024));
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

  const hasNum = (v) => v !== undefined && v !== null && v !== "" && Number.isFinite(Number(v));
  const cost = hasNum(body["仕入原価円"]) ? Number(body["仕入原価円"]) : null;
  const shipping = hasNum(body["送料円"]) ? Number(body["送料円"]) : null;
  const packing = hasNum(body["梱包費円"]) ? Number(body["梱包費円"]) : null;

  let profit = "";
  let margin = "";
  if (cost !== null && shipping !== null && packing !== null) {
    profit = Math.round(revenueJpy - fee - cost - shipping - packing);
    margin = revenueJpy !== 0 ? Number((profit / revenueJpy).toFixed(4)) : "";
  }

  const wb = await loadWorkbook();
  const ws = getSheet(wb);
  ws.addRow([
    body["注文番号"], body["日付"], body["サイト"], body["商品メモ"], body["商品ID"] || "",
    usd, rate, Math.round(revenueJpy), fee,
    cost === null ? "" : cost, shipping === null ? "" : shipping, packing === null ? "" : packing,
    profit, margin,
  ]);
  await wb.xlsx.writeFile(LEDGER_PATH);

  sendJson(res, 200, { status: "ok", 収益円: Math.round(revenueJpy), 手数料円: fee, 最終利益円: profit, 利益率: margin });
}

async function handleListOrders(req, res) {
  const wb = await loadWorkbook();
  const ws = getSheet(wb);
  const rows = [];
  ws.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    rows.push(row.values.slice(1));
  });
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

const server = http.createServer(async (req, res) => {
  if (req.method === "OPTIONS") {
    res.writeHead(204, {
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type, Authorization",
    });
    return res.end();
  }

  if (req.method === "GET" && req.url === "/") {
    res.writeHead(200, { "Content-Type": "text/plain; charset=utf-8" });
    return res.end("agate-tool-server: OK");
  }

  if (req.url === "/api/orders" && !isAuthorized(req)) {
    return sendJson(res, 401, { error: "認証に失敗しました(トークンを確認してください)" });
  }
  if (req.url === "/download/売上管理表.xlsx" && !isAuthorized(req)) {
    return sendJson(res, 401, { error: "認証に失敗しました(トークンを確認してください)" });
  }

  try {
    if (req.method === "POST" && req.url === "/api/orders") {
      return await handleAddOrder(req, res);
    }
    if (req.method === "GET" && req.url === "/api/orders") {
      return await handleListOrders(req, res);
    }
    if (req.method === "GET" && req.url === "/download/売上管理表.xlsx") {
      return await handleDownload(req, res);
    }
  } catch (e) {
    console.error(e);
    return sendJson(res, 500, { error: "サーバー内部でエラーが発生しました" });
  }

  sendJson(res, 404, { error: "not found" });
});

server.listen(PORT, "127.0.0.1", () => {
  console.log(`agate-tool-server listening on 127.0.0.1:${PORT}`);
});
