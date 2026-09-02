(function initOrderParser(root, factory) {
  const api = factory();
  if (typeof module === "object" && module.exports) module.exports = api;
  else root.AgateOrderParser = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function createOrderParser() {
  const MONTH_MAP = { Jan: 1, Feb: 2, Mar: 3, Apr: 4, May: 5, Jun: 6, Jul: 7, Aug: 8, Sep: 9, Oct: 10, Nov: 11, Dec: 12 };

  function previousNonEmptyIndex(lines, start, lowerBound) {
    for (let i = start; i >= lowerBound; i--) {
      if (lines[i]) return i;
    }
    return -1;
  }

  function nextNonEmptyIndex(lines, start, upperBound) {
    for (let i = start; i < upperBound; i++) {
      if (lines[i]) return i;
    }
    return -1;
  }

  function findSubtotalQuantity(lines) {
    for (const line of lines) {
      const match = /^小計[（(]\s*(\d+)\s*点\s*[）)]$/.exec(line);
      if (match) return Number(match[1]);
    }
    return null;
  }

  function isSectionHeading(line, heading) {
    const escaped = heading.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    return new RegExp("^(?:#{1,6}\\s*)?" + escaped + "\\s*$").test(line);
  }

  function findProductSection(lines) {
    for (let start = 0; start < lines.length; start++) {
      if (!isSectionHeading(lines[start], "商品")) continue;
      for (let end = start + 1; end < lines.length; end++) {
        if (isSectionHeading(lines[end], "注文")) return { lines: lines.slice(start + 1, end), found: true };
      }
    }
    return { lines: [], found: false };
  }

  function detectSite(lines) {
    const revenueIndex = lines.indexOf("お客様の収益");
    const upperBound = revenueIndex === -1 ? lines.length : revenueIndex;
    const orderAmountLabels = new Set(["商品価格", "商品合計", "小計", "注文の合計金額"]);
    let hasGbPounds = false;

    for (let i = 0; i < upperBound; i++) {
      if (!/^GB\s*£\s*[-+]?\s*[\d,.]+$/.test(lines[i])) continue;
      const labelIndex = previousNonEmptyIndex(lines, i - 1, 0);
      if (labelIndex !== -1) {
        const normalizedLabel = lines[labelIndex].replace(/[（(]\s*\d+\s*点\s*[）)]$/, "");
        if (orderAmountLabels.has(normalizedLabel)) hasGbPounds = true;
      }
    }

    // US/AUは実サンプル確認前のため推測しない。GBポンドだけをUKへ正規化する。
    return hasGbPounds ? "UK" : "要確認";
  }

  function parseItems(lines) {
    const anchors = [];
    for (let i = 0; i < lines.length; i++) {
      const match = /^商品ID\s*[:：]\s*(\d+)\s*$/.exec(lines[i]);
      if (match) anchors.push({ index: i, ebayItemId: match[1] });
    }

    return anchors.map((anchor, itemIndex) => {
      const lowerBound = itemIndex === 0 ? 0 : anchors[itemIndex - 1].index + 1;
      const upperBound = itemIndex + 1 < anchors.length ? anchors[itemIndex + 1].index : lines.length;
      const errors = [];

      let skuIndex = -1;
      let sku = "";
      for (let i = anchor.index - 1; i >= lowerBound; i--) {
        const skuMatch = /^独自のラベル（SKU）\s*[:：]\s*(.*)$/.exec(lines[i]);
        if (skuMatch) {
          skuIndex = i;
          sku = skuMatch[1].trim();
          break;
        }
      }

      let title = "";
      if (skuIndex === -1) {
        errors.push("SKUラベルがありません");
      } else {
        const titleIndex = previousNonEmptyIndex(lines, skuIndex - 1, lowerBound);
        if (titleIndex !== -1) title = lines[titleIndex];
      }
      if (!title) errors.push("商品タイトルがありません");
      if (!sku) errors.push("SKUがありません");

      const quantityLabels = [];
      for (let i = anchor.index + 1; i < upperBound; i++) {
        if (lines[i] === "数量") quantityLabels.push(i);
      }

      let quantity = null;
      if (quantityLabels.length !== 1) {
        errors.push(quantityLabels.length === 0 ? "数量がありません" : "数量ラベルが複数あります");
      } else {
        const quantityIndex = nextNonEmptyIndex(lines, quantityLabels[0] + 1, upperBound);
        const quantityMatch = quantityIndex === -1 ? null : /^(\d+)$/.exec(lines[quantityIndex]);
        if (!quantityMatch || Number(quantityMatch[1]) < 1) errors.push("販売数量を取得できません");
        else quantity = Number(quantityMatch[1]);
      }

      if (!lines.slice(anchor.index + 1, upperBound).includes("商品価格")) errors.push("商品価格ラベルがありません");
      if (!lines.slice(anchor.index + 1, upperBound).includes("商品合計")) errors.push("商品合計ラベルがありません");

      return {
        title,
        ebayItemId: String(anchor.ebayItemId),
        quantity,
        sku,
        parseStatus: errors.length ? "要確認" : "OK",
        errors,
      };
    });
  }

  function parseOrderText(text) {
    const lines = String(text || "").split(/\r?\n/).map((line) => line.trim());
    const result = {
      orderNo: "", date: "", site: "要確認", usd: "", note: "", qty: "",
      items: [], itemCount: 0, quantityTotal: 0, subtotalQuantity: null,
      parseStatus: "要確認", parseErrors: [],
    };

    for (const line of lines) {
      const match = /^(\d{1,3}-\d{4,7}-\d{4,7})$/.exec(line);
      if (match) { result.orderNo = match[1]; break; }
    }

    const saleIndex = lines.indexOf("販売");
    if (saleIndex !== -1) {
      for (let i = saleIndex + 1; i < lines.length && i < saleIndex + 4; i++) {
        const dateMatch = /^([A-Za-z]{3})[a-z]*\s+(\d{1,2}),\s*(\d{4})/.exec(lines[i]);
        if (!dateMatch) continue;
        const monthKey = dateMatch[1][0].toUpperCase() + dateMatch[1].slice(1, 3).toLowerCase();
        const month = MONTH_MAP[monthKey];
        if (month) {
          result.date = dateMatch[3] + "-" + String(month).padStart(2, "0") + "-" + String(dateMatch[2]).padStart(2, "0");
          break;
        }
      }
    }

    result.site = detectSite(lines);

    const revenueIndex = lines.indexOf("注文の収益");
    if (revenueIndex !== -1) {
      for (let i = revenueIndex + 1; i < lines.length && i < revenueIndex + 3; i++) {
        const usdMatch = /US\s*\$\s*([\d.]+)/.exec(lines[i]);
        if (usdMatch) { result.usd = usdMatch[1]; break; }
      }
    }

    const productSection = findProductSection(lines);
    result.items = productSection.found ? parseItems(productSection.lines) : [];
    result.itemCount = result.items.length;
    result.quantityTotal = result.items.reduce((sum, item) => sum + (Number.isInteger(item.quantity) ? item.quantity : 0), 0);
    result.subtotalQuantity = findSubtotalQuantity(lines);
    result.note = result.items.map((item) => item.title).filter(Boolean).join(" / ");
    result.qty = result.items.length ? String(result.quantityTotal) : "";

    if (!productSection.found) result.parseErrors.push("商品セクションの範囲を特定できません");
    if (!result.items.length) result.parseErrors.push("注文明細がありません");
    if (result.items.some((item) => item.parseStatus !== "OK")) result.parseErrors.push("必須値を取得できない明細があります");
    if (result.subtotalQuantity === null) {
      result.parseErrors.push("小計点数を取得できません");
    } else if (result.quantityTotal !== result.subtotalQuantity) {
      result.parseErrors.push("明細の数量合計と小計が一致しません");
    }
    if (result.site === "要確認") result.parseErrors.push("販売サイトを安全に判定できません");
    result.parseStatus = result.parseErrors.length ? "要確認" : "OK";

    return result;
  }

  return { parseOrderText };
});
