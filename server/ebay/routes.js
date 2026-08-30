// server.jsから呼び出されるeBay関連ルートハンドラの入口
const { handleEbayConnect, handleEbayCallback } = require("./auth");
const { fetchAllActiveListings } = require("./sellerListings");
const { rebuildInventoryFromEbay } = require("./inventorySync");
const { sendJson } = require("../httpUtil");

async function handleEbayInventory(req, res) {
  try {
    const items = await fetchAllActiveListings();
    sendJson(res, 200, { count: items.length, items });
  } catch (e) {
    console.error("[ebay] GetMyeBaySelling 取得エラー:", e);
    sendJson(res, 502, { error: String((e && e.message) || e) });
  }
}

// server.js側の在庫管理表アクセス(INV_HEADERS/INVENTORY_PATH/loadInventoryWorkbook)を
// 実行時に注入して使う。server.jsとebay/配下を相互require(循環参照)にしないための構成。
function createInventoryRebuildHandler(inventoryDeps) {
  return async function handleEbayInventoryRebuild(req, res) {
    try {
      const summary = await rebuildInventoryFromEbay(inventoryDeps);
      sendJson(res, 200, { status: "ok", ...summary });
    } catch (e) {
      console.error("[ebay] 在庫管理表への反映エラー:", e);
      sendJson(res, 502, { error: String((e && e.message) || e) });
    }
  };
}

module.exports = {
  handleEbayConnect,
  handleEbayCallback,
  handleEbayInventory,
  createInventoryRebuildHandler,
};
