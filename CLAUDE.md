# agate-tool-server

株式会社アゲイト・物販事業向けの自社ツール(Node.js製、KAGOYA VPS上でセルフホスト)。
売上分析・在庫・注文の3タブを持つダッシュボードを `server/server.js` で提供する。

## eBay API利用制限(重要・厳守)

今後eBay関連の実装(Shopee連携含む)を行う際は、必ず以下を守ること。

- 使用してよいのは **Inventory API(読み取り専用)のみ**
- 取得してよいのは **自分自身の出品情報**(タイトル・価格・在庫数・SKU・商品説明)だけ
- **Fulfillment API(注文情報API)は、今後の明示的な指示がない限り絶対に呼び出さない**
- **購入者の氏名・住所・連絡先など、他のeBayユーザーの個人情報を取得・保存するコードは一切書かない**

### 理由

eBayの「Marketplace Account Deletion」規約対応との整合性のため。eBay側には「Inventory API(自分の出品データ)のみを取得し、他ユーザーの個人情報は一切収集・保存しない」という内容で適用除外(exemption)を申請済みで、OAuth Scopesも `sell.inventory.readonly` のみに設定済み。実装がこの申告内容と異なると規約違反のリスクがあるため、厳密に守ること。

Shopee側は読み取り・書き込み両方を行ってよい(eBayから抽出した商品・出品データをShopeeに出品するため)。この非対称性(eBay=読み取り専用/Shopee=読み書き可)は意図的な方針であり、変更しない。
