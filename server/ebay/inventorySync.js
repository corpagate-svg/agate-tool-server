// eBay Trading API(GetMyeBaySelling)の取得結果を、既存の在庫管理表フォーマットへ反映する。
//
// 【US起点方式】
// USサイトの出品を基準リストとし、US 1出品 = メインシート1行を必ず維持する。
// UK/AUの出品は、同じ商品と思われるUS出品に名寄せ(タイトル一致)で紐付けるだけで、
// 単独の行としては絶対に追加しない。どのUS出品にも紐付かなかったUK/AU出品は、
// 削除せず「UK_AU保管(US未紐付け)」シートに保管し、次回実行時に改めて評価し直す
// (毎回eBayから取得した最新の生データだけを元に保管シートを作り直すため、
//  紐付いた/削除されたUK・AU出品は自然に保管シートから消える)。
//
// 【タイトル照合方針(2026-09-05〜)】
// 今後タイトルへ送料訴求文([Extra Items Ship FREE]等)を追加しない運用のため、
// 照合は「前後空白の除去+連続空白の1個化(titleKey)」だけを行い、完全一致した
// US/UK/AU出品だけを紐付ける。fuzzy matching・編集距離・類似度判定・prefix一致・
// 「...」「…」からの推測は一切行わない。一致しないものは推測せず「UK_AU保管」
// (未紐付け)へ残す。メインシートの「商品名」列にはUSのタイトルを一切加工せず
// (送料文言の除去も行わず)そのまま保存する(raw title)。UK/AUタイトルは紐付け
// 判定にのみ一時的に使用し、紐付け済みの行には保存しない(在庫管理表の列構成を
// 変えないため。未紐付けの場合のみ既存の「UK_AU保管」シートに生タイトルのまま残る)。
const ExcelJS = require("exceljs");
const crypto = require("crypto");
const { fetchAllActiveListings } = require("./sellerListings");
const { withInventoryLock, atomicWriteWorkbook } = require("../inventoryLock");
const { copyProtectedSheets, addReplenishmentCandidate, REPLENISHMENT_STATUS } = require("../inventoryProtectedSheets");
const {
  normText,
  INV_HEADER_FILL,
  INV_HEADER_FONT,
  INV_FLAG_FILL,
} = require("../inventoryRebuild");

function siteFromViewItemUrl(url) {
  try {
    const host = new URL(url).hostname;
    if (host.includes("ebay.com.au")) return "AU";
    if (host.includes("ebay.co.uk")) return "UK";
    if (host.includes("ebay.com")) return "US";
  } catch (e) {
    // 無効なURLは判定不能として扱う
  }
  return null;
}

// US/UK/AUの照合キー。今後タイトルへ送料訴求文を追加しない運用のため、
// 送料文言の除去(stripShippingNote)には依存せず、前後空白の除去と連続空白の
// 1個化だけを行う。それ以外は元タイトルを一切変更しない(大文字小文字も区別する)。
// 意図的にfuzzy matching・prefix一致・類似度判定は行わない。
function titleKey(title) {
  return normText(title);
}

function startTimeMs(item) {
  if (!item || !item.startTime) return null;
  const t = Date.parse(item.startTime);
  return Number.isFinite(t) ? t : null;
}

function inventoryQuantity(value) {
  if (value === null || value === undefined || value === "") return null;
  const quantity = Number(value);
  return Number.isSafeInteger(quantity) && quantity >= 0 ? quantity : null;
}

// 複数候補がある場合、USの出品開始日時に最も近いものを選ぶ。
// eBay側にUS-UK/AUを直接結びつけるコード(リンクID)が無いため、
// 現状これが唯一使える実用的な手がかり。
function pickBestMatch(candidates, anchorStartTime) {
  if (candidates.length === 0) return { match: null, ambiguous: false };
  if (candidates.length === 1) return { match: candidates[0], ambiguous: false };
  const anchorMs = anchorStartTime ? Date.parse(anchorStartTime) : null;
  let best = candidates[0];
  let bestDiff = Infinity;
  for (const c of candidates) {
    const t = startTimeMs(c);
    const diff = (Number.isFinite(anchorMs) && t !== null) ? Math.abs(t - anchorMs) : Infinity;
    if (diff < bestDiff) { bestDiff = diff; best = c; }
  }
  return { match: best, ambiguous: true };
}

function buildPool(items) {
  const pool = new Map();
  for (const it of items) {
    const k = titleKey(it.title);
    if (!pool.has(k)) pool.set(k, []);
    pool.get(k).push(it);
  }
  return pool;
}

function takeMatch(pool, key, anchorStartTime) {
  const list = pool.get(key);
  if (!list || list.length === 0) return { match: null, ambiguous: false };
  const { match, ambiguous } = pickBestMatch(list, anchorStartTime);
  list.splice(list.indexOf(match), 1);
  return { match, ambiguous };
}

// eBay APIへの通信(数秒〜数十秒かかり得る)はロックの外で行い、他の保存操作を
// 不必要に待たせない。ロックを取るのは「在庫管理表.xlsxの読み込み→マージ→書き込み」
// の間だけにし、しかもロック取得後に改めて最新のxlsxを読み込む(eBay通信中に行われた
// 棚卸数量・仕入価格・日本語商品名・リアル在庫の変更を、古い状態で上書きして消さないため)。
async function rebuildInventoryFromEbay({ INV_HEADERS, INVENTORY_PATH, loadInventoryWorkbook, fetchListings = fetchAllActiveListings }) {
  const allItems = await fetchListings();

  const usRawAll = allItems.filter((it) => siteFromViewItemUrl(it.viewItemUrl) === "US");
  const ukRawAll = allItems.filter((it) => siteFromViewItemUrl(it.viewItemUrl) === "UK");
  const auRawAll = allItems.filter((it) => siteFromViewItemUrl(it.viewItemUrl) === "AU");
  const unknownSiteCount = allItems.length - usRawAll.length - ukRawAll.length - auRawAll.length;

  // バリエーション(色・サイズ違い等)出品は対象外
  const usVariationCount = usRawAll.filter((it) => it.hasVariations).length;
  const ukVariationCount = ukRawAll.filter((it) => it.hasVariations).length;
  const auVariationCount = auRawAll.filter((it) => it.hasVariations).length;

  const usItemsRaw = usRawAll.filter((it) => !it.hasVariations);
  const ukItems = ukRawAll.filter((it) => !it.hasVariations);
  const auItems = auRawAll.filter((it) => !it.hasVariations);
  const usTargetCount = usItemsRaw.length;

  // US出品を開始日時の早い順に処理する(割り当てを決定的にするため)
  const usItems = [...usItemsRaw].sort((a, b) => (startTimeMs(a) || 0) - (startTimeMs(b) || 0));

  const ukPool = buildPool(ukItems);
  const auPool = buildPool(auItems);

  // ---- ここから在庫管理表.xlsxへの読み込み→マージ→書き込み。ここだけ排他ロックする ----
  return withInventoryLock(async () => {
  // 既存シートを読み込み、US_出品ID(最優先)またはタイトル(予備)で
  // 商品ID・仕入情報を引き継ぐ。ロック取得後に改めて読み込むことで、eBay通信中に
  // 保存された最新の変更を土台にできる。
  const oldWb = await loadInventoryWorkbook();
  const oldWs = oldWb.getWorksheet("在庫管理表") || oldWb.worksheets[0];
  const savedByUsId = new Map();
  const savedUsIdCounts = new Map();
  const savedByTitle = new Map();
  let maxId = 0;
  oldWs.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const pid = row.getCell(1).value;
    if (!pid) return;
    const idNum = parseInt(String(pid).slice(1), 10);
    if (Number.isFinite(idNum)) maxId = Math.max(maxId, idNum);
    const fullRow = [];
    for (let c = 1; c <= INV_HEADERS.length; c++) fullRow.push(row.getCell(c).value);
    const savedRow = {
      商品ID: pid,
      fullRow,
      wasRemoved: String(fullRow[18] || "").includes("見当たりません"),
    };
    const usId = fullRow[6];
    if (usId !== null && usId !== undefined && usId !== "") {
      const usIdKey = String(usId);
      savedByUsId.set(usIdKey, savedRow);
      savedUsIdCounts.set(usIdKey, (savedUsIdCounts.get(usIdKey) || 0) + 1);
    }
    const tKey = titleKey(fullRow[1]);
    if (!savedByTitle.has(tKey)) savedByTitle.set(tKey, []);
    savedByTitle.get(tKey).push(savedRow);
  });

  const claimedSaved = new Set();
  const today = new Date().toISOString().slice(0, 10);
  const rowsOut = [];
  const newProducts = [];
  let nextId = maxId + 1;
  let ambiguousCount = 0;
  let replenishmentCandidatesSkipped = 0;
  const replenishmentCandidates = [];
  const syncExecutionId = crypto.randomUUID();
  const detectedAt = new Date().toISOString();

  for (const usItem of usItems) {
    const key = titleKey(usItem.title);
    const ukResult = takeMatch(ukPool, key, usItem.startTime);
    const auResult = takeMatch(auPool, key, usItem.startTime);
    if (ukResult.ambiguous || auResult.ambiguous) ambiguousCount++;

    const siteCount = 1 + (ukResult.match ? 1 : 0) + (auResult.match ? 1 : 0);
    const qtys = [usItem.quantityAvailable, ukResult.match && ukResult.match.quantityAvailable, auResult.match && auResult.match.quantityAvailable]
      .filter((v) => v !== null && v !== undefined);
    const qtyMismatch = new Set(qtys).size > 1;
    const flag = (ukResult.ambiguous || auResult.ambiguous || qtyMismatch || siteCount < 3) ? "要確認" : "";

    const usItemId = String(usItem.itemId);
    const exactSavedRow = savedByUsId.get(usItemId);
    const hasUniqueExactMatch = Boolean(exactSavedRow) && savedUsIdCounts.get(usItemId) === 1;
    let savedRow = exactSavedRow;
    if (!savedRow) {
      const candidates = (savedByTitle.get(key) || []).filter((s) => !claimedSaved.has(s));
      savedRow = candidates[0];
    }

    let pid;
    let purchasePrice = null, purchaseDate = null, purchaseFrom = null, note = "";
    let realStock = null, realStockConfirmedAt = null, stocktakeQty = null, stocktakeAt = null;
    let jaName = null, stocktakeChecked = null;
    if (savedRow) {
      claimedSaved.add(savedRow);
      pid = savedRow.商品ID;
      purchasePrice = savedRow.fullRow[15];
      purchaseDate = savedRow.fullRow[16];
      purchaseFrom = savedRow.fullRow[17];
      note = savedRow.wasRemoved ? "" : (savedRow.fullRow[18] || "");
      realStock = savedRow.fullRow[21] ?? null;
      realStockConfirmedAt = savedRow.fullRow[22] ?? null;
      stocktakeQty = savedRow.fullRow[23] ?? null;
      stocktakeAt = savedRow.fullRow[24] ?? null;
      jaName = savedRow.fullRow[26] ?? null;
      stocktakeChecked = savedRow.fullRow[28] ?? null;
    } else {
      pid = "P" + String(nextId).padStart(4, "0");
      nextId++;
      newProducts.push([pid, usItem.title]);
    }
    if (!purchaseDate && usItem.startTime) purchaseDate = usItem.startTime.slice(0, 10);
    // リアル在庫は当社独自管理の値。eBay APIでの再取込では上書きしない。
    // 未設定(導入前・移行直後)の場合のみ、初期値としてUSの在庫数をコピーする。
    if (realStock === null || realStock === undefined) realStock = usItem.quantityAvailable;

    // 補充候補は、同じUS Item IDで一意に特定できる既存商品のUS在庫が増えた場合だけ作る。
    // タイトルによる引継ぎ・新規商品・不正数量は誤紐付けを避けるため候補化しない。
    if (hasUniqueExactMatch && savedRow === exactSavedRow) {
      const oldUsQty = inventoryQuantity(savedRow.fullRow[3]);
      const newUsQty = inventoryQuantity(usItem.quantityAvailable);
      if (!/^P\d+$/.test(String(pid))) {
        replenishmentCandidatesSkipped++;
      } else if (oldUsQty !== null && newUsQty !== null) {
        if (newUsQty > oldUsQty) {
          replenishmentCandidates.push({
            "商品ID": String(pid),
            "US出品ID": usItemId,
            "同期前US在庫": oldUsQty,
            "同期後US在庫": newUsQty,
            "補充候補数量": newUsQty - oldUsQty,
            "検知日時": detectedAt,
            "検知元": `eBay同期:${syncExecutionId}`,
            "状態": REPLENISHMENT_STATUS.PENDING,
          });
        }
      } else {
        replenishmentCandidatesSkipped++;
      }
    } else if (savedRow) {
      // 既存行をタイトルだけで引き継いだ場合やUS Item IDが重複する場合は、安全に商品を
      // 特定できないため候補を作らない。同期結果のsummaryへ件数だけ残す。
      const oldUsQty = inventoryQuantity(savedRow.fullRow[3]);
      const newUsQty = inventoryQuantity(usItem.quantityAvailable);
      if (oldUsQty !== null && newUsQty !== null && newUsQty > oldUsQty) {
        replenishmentCandidatesSkipped++;
      }
    }

    rowsOut.push({
      row: [
        pid,
        // 商品名はeBayから取得したUSタイトルをそのまま保存する(raw title)。
        // 送料文言の除去や表示用の省略は行わず、照合(titleKey)専用に別途正規化する。
        usItem.title,
        null,
        usItem.quantityAvailable,
        siteCount,
        flag,
        Number(usItem.itemId),
        usItem.price,
        usItem.quantitySold,
        ukResult.match ? Number(ukResult.match.itemId) : null,
        ukResult.match ? ukResult.match.price : null,
        ukResult.match ? ukResult.match.quantitySold : null,
        auResult.match ? Number(auResult.match.itemId) : null,
        auResult.match ? auResult.match.price : null,
        auResult.match ? auResult.match.quantitySold : null,
        purchasePrice, purchaseDate, purchaseFrom, note,
        ukResult.match ? ukResult.match.quantityAvailable : null,
        auResult.match ? auResult.match.quantityAvailable : null,
        realStock, realStockConfirmedAt, stocktakeQty, stocktakeAt,
        usItem.galleryUrl || null, jaName, usItem.sku || null, stocktakeChecked,
      ],
      flag: Boolean(flag),
    });
  }

  // US起点方式での削除判定: メインシートに残すのは「今回のUS Active Listings
  // (バリエーション除外後)に実際に存在するItemID」の商品だけ。
  // 旧行のUS_出品IDが今回の一覧に見当たらない場合は、出品終了と判断してメインシートから削除する
  // (UK/AUの有無・在庫数0は削除判定に一切使わない。US ItemIDの有無のみで判定)。
  const allSavedRows = new Set([...savedByUsId.values(), ...[...savedByTitle.values()].flat()]);
  let deletedCount = 0;
  for (const s of allSavedRows) {
    if (claimedSaved.has(s)) continue;
    deletedCount++;
  }

  rowsOut.sort((a, b) => String(a.row[0]).localeCompare(String(b.row[0])));

  // どのUSにも紐付かなかったUK/AU出品(=保管対象)
  const orphanRows = [];
  for (const [, list] of ukPool) for (const it of list) orphanRows.push({ site: "UK", ...it });
  for (const [, list] of auPool) for (const it of list) orphanRows.push({ site: "AU", ...it });

  // ---- xlsx書き出し ----
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("在庫管理表");
  ws.addRow(INV_HEADERS);
  for (let c = 1; c <= INV_HEADERS.length; c++) {
    const cell = ws.getRow(1).getCell(c);
    cell.fill = INV_HEADER_FILL;
    cell.font = INV_HEADER_FONT;
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  }
  ws.getRow(1).height = 30;
  ws.views = [{ state: "frozen", ySplit: 1 }];

  rowsOut.forEach((r) => {
    const excelRow = ws.addRow(r.row);
    if (r.flag) {
      for (let c = 1; c <= INV_HEADERS.length; c++) excelRow.getCell(c).fill = INV_FLAG_FILL;
    }
  });
  const widths = [10, 40, 30, 12, 10, 12, 14, 12, 12, 14, 12, 12, 14, 12, 12, 14, 12, 16, 30, 12, 12, 12, 16, 12, 18, 40, 30, 16, 12];
  widths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });

  const ws2 = wb.addWorksheet("UK_AU保管(US未紐付け)");
  ws2.addRow(["サイト", "出品ID", "商品名", "在庫数", "累計売却数", "価格", "通貨", "出品開始日時"]);
  for (let c = 1; c <= 8; c++) {
    ws2.getRow(1).getCell(c).fill = INV_HEADER_FILL;
    ws2.getRow(1).getCell(c).font = INV_HEADER_FONT;
  }
  ws2.views = [{ state: "frozen", ySplit: 1 }];
  orphanRows.forEach((it) => {
    ws2.addRow([it.site, Number(it.itemId) || it.itemId, it.title, it.quantityAvailable, it.quantitySold, it.price, it.currency, it.startTime]);
  });
  [8, 14, 45, 10, 12, 10, 10, 24].forEach((w, i) => { ws2.getColumn(i + 1).width = w; });

  const ws3 = wb.addWorksheet("この表について");
  ws3.getCell("A1").value = "在庫管理表について";
  ws3.getCell("A1").font = { bold: true, size: 13 };
  const notes = [
    `・eBay Trading API(GetMyeBaySelling / ActiveList)の取得結果(${today}時点)を元に、USサイトの出品を基準として自動更新しました。`,
    "・メインシートの行数は「USサイトの出品数(バリエーション出品を除く)」と必ず一致します。UK/AUの出品は対応するUS商品の行に付随情報として紐づけるだけで、単独の行としては追加していません。",
    "・商品名(タイトル)はeBayから取得した文字列をそのまま保存しています(前後・連続する空白の違いだけは無視して一致と判定しますが、それ以外は完全一致した場合のみ紐付けます。あいまいな類似判定は行いません)。",
    "・黄色でハイライトした行は「UK/AUの対応付けが複数候補から推定で選ばれた」「在庫数が国ごとに違う」「UK/AUが3カ国揃っていない」のいずれかに該当します。",
    "・今回のUS Active Listingsに存在しないUS出品ID(=出品終了・売り切れ後の自動終了などでeBayから消えたもの)は、メインシートから削除しています(UK/AUの有無や在庫数0は削除理由にしていません)。",
    "・UK/AUが複数候補ある場合は、US出品の出品開始日時に最も近いものを自動選択しています(確実な保証はできないため、該当行は黄色でハイライトしています)。",
    "・UK/AUの出品のうち、どのUS出品にも対応付けられなかったものは「UK_AU保管(US未紐付け)」シートに保管しています。次回の取り込み時に対応するUS出品が見つかれば自動的にメインシートへ統合され、eBay上で削除・売り切れが確認された場合は保管領域からも自動的に取り除かれます。",
    "・商品IDは、USの出品ID(Item Number)が前回と同じ場合はそのまま引き継ぎます。出品IDが変わった場合でも商品名が一致すれば引き継ぎます。仕入価格・仕入日・仕入先・備考は前回入力済みの内容をそのまま引き継ぎます。",
    "・「リアル在庫」「リアル在庫確認日」「棚卸入力数量」「棚卸入力日時」は当社独自管理の値のため、このeBay同期では一切上書きしません(前回の値をそのまま引き継ぎます)。",
    "・「UK在庫数」「AU在庫数」は、eBay自己申告の在庫数をそのまま反映したものです(リアル在庫とは別物です)。",
    "・注意: US出品が削除・終了してメインシートから消えた商品は、その時点のリアル在庫の記録もこの表からは消えます(削除自体は行われず単に一覧から除外されるため、現物在庫が残っている場合は「相違」ページや在庫変更履歴で確認してください)。",
    "・「画像URL」「SKU」はeBayの最新値をそのまま反映したものです(毎回更新されます)。「日本語商品名」は当社独自入力のため、このeBay同期では一切上書きしません。",
    "・「棚卸チェック」は棚卸で実際に数えたかどうかを管理する当社独自のフラグです(このeBay同期では一切上書きしません。解除は棚卸画面の「棚卸チェックをすべて解除」ボタンでのみ行います)。",
  ];
  notes.forEach((n, i) => {
    const cell = ws3.getCell(`A${i + 3}`);
    cell.value = n;
    cell.font = { italic: true, color: { argb: "FF808080" }, size: 9 };
  });
  ws3.getColumn(1).width = 110;

  copyProtectedSheets(oldWb, wb);
  for (const candidate of replenishmentCandidates) addReplenishmentCandidate(wb, candidate);
  await atomicWriteWorkbook(wb, INVENTORY_PATH);

  return {
    usActiveListingsTotal: usRawAll.length,
    usVariationsExcluded: usVariationCount,
    usTargetCount,
    mainSheetCurrentRows: rowsOut.length,
    total: rowsOut.length,
    new: newProducts.length,
    newProducts,
    deletedCount,
    ambiguousMatches: ambiguousCount,
    ukAuOrphanCount: orphanRows.length,
    variationsExcludedTotal: usVariationCount + ukVariationCount + auVariationCount,
    unknownSiteSkipped: unknownSiteCount,
    replenishmentCandidatesCreated: replenishmentCandidates.length,
    replenishmentCandidatesSkipped,
  };
  }); // withInventoryLock ここまで
}

module.exports = { rebuildInventoryFromEbay, siteFromViewItemUrl };
