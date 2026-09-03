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
// afterRebuild(省略可): 在庫反映成功後に実行する追加処理(前月為替の自動確定など)。
// 独立したtry/catchで囲み、ここで例外が起きても在庫反映自体の成功レスポンスには影響しない
// (為替側の問題でeBay同期そのものが失敗扱いにならないようにするための設計)。
function createInventoryRebuildHandler(inventoryDeps, afterRebuild) {
  return async function handleEbayInventoryRebuild(req, res) {
    let summary;
    try {
      summary = await rebuildInventoryFromEbay(inventoryDeps);
    } catch (e) {
      console.error("[ebay] 在庫管理表への反映エラー:", e);
      return sendJson(res, 502, { error: String((e && e.message) || e) });
    }
    let fx = null;
    if (afterRebuild) {
      try {
        fx = await afterRebuild();
      } catch (e) {
        console.error("[ebay] 月次為替確定処理でエラー(在庫反映自体は成功):", e);
        fx = { action: "error", error: String((e && e.message) || e) };
      }
    }
    sendJson(res, 200, { status: "ok", ...summary, fx });
  };
}

module.exports = {
  handleEbayConnect,
  handleEbayCallback,
  handleEbayInventory,
  createInventoryRebuildHandler,
};
