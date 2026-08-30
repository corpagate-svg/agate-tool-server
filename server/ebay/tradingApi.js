// eBay Trading API(XML/SOAP形式)の共通呼び出し処理。読み取り専用の呼び出しのみを想定。
const TRADING_API_ENDPOINT = "https://api.ebay.com/ws/api.dll";
const TRADING_API_COMPATIBILITY_LEVEL = "1219";
const TRADING_API_SITE_ID = "0"; // 0 = eBay.com (US)

function decodeXmlEntities(str) {
  return String(str)
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

function extractAck(xmlText) {
  const m = /<Ack>([^<]*)<\/Ack>/.exec(xmlText);
  return m ? m[1] : null;
}

function extractErrorSummary(xmlText) {
  const blocks = xmlText.match(/<Errors>[\s\S]*?<\/Errors>/g) || [];
  return blocks
    .map((b) => {
      const code = /<ErrorCode>([^<]*)<\/ErrorCode>/.exec(b);
      const short = /<ShortMessage>([^<]*)<\/ShortMessage>/.exec(b);
      return `[${code ? code[1] : "?"}] ${short ? decodeXmlEntities(short[1]) : ""}`;
    })
    .join("; ");
}

async function callTradingApi(callName, xmlBody, accessToken) {
  const res = await fetch(TRADING_API_ENDPOINT, {
    method: "POST",
    headers: {
      "X-EBAY-API-COMPATIBILITY-LEVEL": TRADING_API_COMPATIBILITY_LEVEL,
      "X-EBAY-API-CALL-NAME": callName,
      "X-EBAY-API-SITEID": TRADING_API_SITE_ID,
      "X-EBAY-API-IAF-TOKEN": accessToken,
      "Content-Type": "text/xml",
    },
    body: xmlBody,
  });
  const text = await res.text();
  const ack = extractAck(text);
  if (!res.ok || ack === "Failure" || ack === null) {
    const errorSummary = extractErrorSummary(text) || text.slice(0, 300);
    throw new Error(`Trading API ${callName} エラー(HTTP ${res.status}, Ack=${ack || "unknown"}): ${errorSummary}`);
  }
  return text;
}

module.exports = { callTradingApi, decodeXmlEntities };
