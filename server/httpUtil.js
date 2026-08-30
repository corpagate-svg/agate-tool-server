function sendJson(res, status, obj) {
  const body = JSON.stringify(obj);
  res.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Access-Control-Allow-Origin": "*",
  });
  res.end(body);
}

function sendHtml(res, html) {
  res.writeHead(200, { "Content-Type": "text/html; charset=utf-8" });
  res.end(html);
}

// 日本語などASCII外の文字を含むファイル名をContent-Dispositionヘッダーに安全に載せる。
// HTTPヘッダー値にはASCII外の文字(日本語等)をそのまま入れられない(Nodeが例外を投げる)ため、
// ASCII向けの簡易フォールバック名(filename=)と、RFC 5987準拠のUTF-8名(filename*=)を両方指定する。
function contentDispositionAttachment(filename) {
  const asciiFallback = String(filename).replace(/[^\x20-\x7e]/g, "_");
  const encoded = encodeURIComponent(filename);
  return `attachment; filename="${asciiFallback}"; filename*=UTF-8''${encoded}`;
}

module.exports = { sendJson, sendHtml, contentDispositionAttachment };
