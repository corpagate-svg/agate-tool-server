const { parseOrderText } = require("./orderParser");

function resolveManualInventoryLink(body) {
  const requestedPid = String(body && body["商品ID"] || "").trim();
  const hasPastedOrderText = body && Object.prototype.hasOwnProperty.call(body, "注文貼付テキスト");

  // 原文を再解析できない場合は単一明細と証明できないため、売上登録だけを許可し、
  // 手動Pxxxxによる在庫連動は行わない。
  if (!hasPastedOrderText) {
    return {
      pid: "",
      blocked: Boolean(requestedPid),
      reason: requestedPid ? "注文貼り付け原文を確認できないため、リアル在庫連動を行いませんでした" : "",
      parsed: null,
    };
  }

  const parsed = parseOrderText(body["注文貼付テキスト"]);
  const allowed = parsed.parseStatus === "OK" && parsed.itemCount === 1;
  if (allowed) return { pid: requestedPid, blocked: false, reason: "", parsed };

  const reason = parsed.itemCount > 1
    ? "複数商品注文のため、単一Pxxxxによるリアル在庫連動を行いませんでした"
    : "解析状態が要確認のため、リアル在庫連動を行いませんでした";
  return { pid: "", blocked: Boolean(requestedPid), reason, parsed };
}

function canEditOrderProductId(oldPid, newPid) {
  const oldValue = String(oldPid || "").trim();
  const newValue = String(newPid || "").trim();
  return Boolean(oldValue) || !newValue;
}

module.exports = { resolveManualInventoryLink, canEditOrderProductId };
