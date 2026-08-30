const fs = require("fs");
const path = require("path");
const { sendJson, sendHtml } = require("../httpUtil");

// eBay OAuth連携(server.jsから移設。ロジックは変更していません)
const DATA_DIR = path.join(__dirname, "..", "..", "data");
const EBAY_APP_ID = process.env.EBAY_APP_ID || "";
const EBAY_CERT_ID = process.env.EBAY_CERT_ID || "";
const EBAY_RUNAME = process.env.EBAY_RUNAME || "";
const EBAY_SCOPE = "https://api.ebay.com/oauth/api_scope/sell.inventory.readonly";
const EBAY_TOKEN_PATH = path.join(DATA_DIR, "ebay_token.json");

function loadEbayToken() {
  try {
    return JSON.parse(fs.readFileSync(EBAY_TOKEN_PATH, "utf8"));
  } catch (e) {
    return null;
  }
}

function saveEbayToken(data) {
  fs.writeFileSync(EBAY_TOKEN_PATH, JSON.stringify(data, null, 2));
}

let ebayAccessTokenCache = { token: null, expiresAt: 0 };

async function ebayTokenRequest(params) {
  const basic = Buffer.from(`${EBAY_APP_ID}:${EBAY_CERT_ID}`).toString("base64");
  const res = await fetch("https://api.ebay.com/identity/v1/oauth2/token", {
    method: "POST",
    headers: {
      "Authorization": `Basic ${basic}`,
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body: new URLSearchParams(params).toString(),
  });
  const text = await res.text();
  let data;
  try {
    data = JSON.parse(text);
  } catch (e) {
    throw new Error(`eBayトークン取得に失敗しました(${res.status}): ${text.slice(0, 200)}`);
  }
  if (!res.ok) {
    throw new Error(data.error_description || data.error || `eBayトークン取得に失敗しました(${res.status})`);
  }
  return data;
}

async function ebayExchangeCode(code) {
  return ebayTokenRequest({ grant_type: "authorization_code", code, redirect_uri: EBAY_RUNAME });
}

async function ebayRefreshAccessToken(refreshToken) {
  return ebayTokenRequest({ grant_type: "refresh_token", refresh_token: refreshToken, scope: EBAY_SCOPE });
}

async function getEbayAccessToken() {
  const now = Date.now();
  if (ebayAccessTokenCache.token && ebayAccessTokenCache.expiresAt > now + 60 * 1000) {
    return ebayAccessTokenCache.token;
  }
  const stored = loadEbayToken();
  if (!stored || !stored.refresh_token) {
    throw new Error("eBayとの連携がまだ完了していません(/ebay/connect から認可を行ってください)");
  }
  const data = await ebayRefreshAccessToken(stored.refresh_token);
  ebayAccessTokenCache = { token: data.access_token, expiresAt: now + data.expires_in * 1000 };
  return data.access_token;
}

async function handleEbayConnect(req, res) {
  if (!EBAY_APP_ID || !EBAY_RUNAME) {
    return sendJson(res, 500, { error: "EBAY_APP_ID / EBAY_RUNAME が設定されていません(.envを確認してください)" });
  }
  const params = new URLSearchParams({
    client_id: EBAY_APP_ID,
    redirect_uri: EBAY_RUNAME,
    response_type: "code",
    scope: EBAY_SCOPE,
  });
  res.writeHead(302, { Location: `https://auth.ebay.com/oauth2/authorize?${params.toString()}` });
  res.end();
}

async function handleEbayCallback(req, res) {
  const urlObj = new URL(req.url, "http://localhost");
  const code = urlObj.searchParams.get("code");
  if (!code) {
    return sendHtml(res, "<p>認可コードが見つかりませんでした。もう一度 /ebay/connect からやり直してください。</p>");
  }
  try {
    const data = await ebayExchangeCode(code);
    saveEbayToken({
      refresh_token: data.refresh_token,
      refresh_token_expires_in: data.refresh_token_expires_in,
      connected_at: new Date().toISOString(),
    });
    ebayAccessTokenCache = { token: data.access_token, expiresAt: Date.now() + data.expires_in * 1000 };
    return sendHtml(res, "<p>eBayとの連携が完了しました。このタブは閉じて構いません。</p>");
  } catch (e) {
    return sendHtml(res, "<p>連携に失敗しました: " + String((e && e.message) || e) + "</p>");
  }
}

module.exports = {
  getEbayAccessToken,
  handleEbayConnect,
  handleEbayCallback,
};
