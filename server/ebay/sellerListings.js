// GetMyeBaySelling(Trading API)で現在の出品中商品(ActiveList)を取得する
const { getEbayAccessToken } = require("./auth");
const { callTradingApi, decodeXmlEntities } = require("./tradingApi");

const ENTRIES_PER_PAGE = 200; // GetMyeBaySellingの1ページあたり最大件数

function buildGetMyeBaySellingXml(pageNumber) {
  return `<?xml version="1.0" encoding="utf-8"?>
<GetMyeBaySellingRequest xmlns="urn:ebay:apis:eBLBaseComponents">
  <ErrorLanguage>en_US</ErrorLanguage>
  <WarningLevel>High</WarningLevel>
  <ActiveList>
    <Sort>TimeLeft</Sort>
    <Pagination>
      <EntriesPerPage>${ENTRIES_PER_PAGE}</EntriesPerPage>
      <PageNumber>${pageNumber}</PageNumber>
    </Pagination>
  </ActiveList>
</GetMyeBaySellingRequest>`;
}

function getTag(block, tag) {
  const m = new RegExp(`<${tag}(?:\\s[^>]*)?>([^<]*)</${tag}>`).exec(block);
  return m ? decodeXmlEntities(m[1]) : null;
}

function parseActiveItems(xmlText) {
  const activeListMatch = /<ActiveList>[\s\S]*?<\/ActiveList>/.exec(xmlText);
  if (!activeListMatch) return [];
  const blocks = activeListMatch[0].split("<Item>").slice(1).map((s) => s.split("</Item>")[0]);
  return blocks.map((block) => {
    const quantityRaw = getTag(block, "Quantity");
    const soldRaw = getTag(block, "QuantitySold");
    const availableRaw = getTag(block, "QuantityAvailable");
    const quantity = quantityRaw !== null ? Number(quantityRaw) : null;
    const quantitySold = soldRaw !== null ? Number(soldRaw) : 0;
    // QuantityAvailableが無い場合は 出品数-売却数 で補完(オークション形式など)
    const quantityAvailable = availableRaw !== null
      ? Number(availableRaw)
      : (quantity !== null ? quantity - quantitySold : null);
    const priceMatch = /<CurrentPrice currencyID="([^"]*)">([^<]*)<\/CurrentPrice>/.exec(block);
    return {
      itemId: getTag(block, "ItemID"),
      title: getTag(block, "Title"),
      sku: getTag(block, "SKU"),
      active: true, // ActiveListのみを取得しているため常にtrue
      quantity,
      quantitySold,
      quantityAvailable,
      price: priceMatch ? Number(priceMatch[2]) : null,
      currency: priceMatch ? priceMatch[1] : null,
      viewItemUrl: getTag(block, "ViewItemURL"),
      startTime: getTag(block, "StartTime"),
      hasVariations: block.includes("<Variations>"), // バリエーション(色・サイズ違い等)出品かどうか
    };
  });
}

function extractTotalPages(xmlText) {
  const m = /<PaginationResult>[\s\S]*?<TotalNumberOfPages>(\d+)<\/TotalNumberOfPages>[\s\S]*?<\/PaginationResult>/.exec(xmlText);
  return m ? Number(m[1]) : 1;
}

async function fetchAllActiveListings() {
  const accessToken = await getEbayAccessToken();
  const all = [];
  let page = 1;
  let totalPages = 1;
  do {
    const xml = buildGetMyeBaySellingXml(page);
    const responseText = await callTradingApi("GetMyeBaySelling", xml, accessToken);
    all.push(...parseActiveItems(responseText));
    totalPages = extractTotalPages(responseText);
    page += 1;
  } while (page <= totalPages);
  return all;
}

module.exports = { fetchAllActiveListings };
