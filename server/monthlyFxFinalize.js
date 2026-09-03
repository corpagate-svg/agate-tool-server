// 「eBay最新情報を取り込む」実行時に、終了済みの前月為替を自動確定し、その月の注文を
// 確定レートへ揃える恒久機能。eBay同期そのものとは責務を分離し、依存はすべて呼び出し元
// (server.js)から注入で受け取る(既存のロック・Excel書き込み・レート計算ロジックを
// 重複実装せずそのまま再利用するための構成。exchangeRateFetcher.js/exchangeRateStore.js
// の既存責務・既存の確定保護(confirmRate)は一切変更しない)。

// 「前月が安全に確定できる」条件: 前月分のBOJ日次データが1件以上あり、かつ「翌月」
// (= 呼び出し時点の実際の現在月)のBOJデータも既に1件以上存在すること。
// BOJの日次系列は時系列順に公表されるため、「翌月分のデータが存在する」という事実自体が
// 「前月分の公表はもう追いつき切っている」ことの直接的な証拠になる。月末の実際の営業日が
// 何日かを知る必要がなく、祝日・連休の並びに関わらず常に正しく判定できる
// (「月末から何日以内」のような暦日ベースの閾値は、連休が月末にかかる場合に前月の最終
// 営業日がその閾値より前になり得るため、永久にスキップし続める恐れがあり採用しない)。
async function isPreviousMonthSafeToConfirm({ targetYearMonth, currentYearMonth, fetchMonthlyAverageRate }) {
  const targetResult = await fetchMonthlyAverageRate(targetYearMonth);
  if (!targetResult.ok) return { safe: false, reason: "target_month_fetch_failed", targetResult };
  const nextMonthResult = await fetchMonthlyAverageRate(currentYearMonth);
  if (!nextMonthResult.ok) return { safe: false, reason: "next_month_not_started", targetResult };
  return { safe: true, targetResult };
}

// 前月の為替を必要なら確定し、その月の注文を確定レートへ揃える。冪等(何度呼んでも安全)。
// 為替確定と注文更新の両方が完了して初めて「完了」とみなす設計ではなく、両者を独立した
// 事実として毎回チェックし直す(為替は確定済みだが注文更新が未完了/失敗した場合でも、
// 次回呼び出し時に「確定済みレコードを読み、レートが違う注文だけ更新する」という同じ経路を
// 通るため、途中失敗からの再開が自然に成立する。専用の完了フラグ列を新設する必要がない)。
async function finalizePreviousMonthFx(deps) {
  const {
    EXCHANGE_RATE_PATH, LEDGER_PATH,
    currentYearMonth, previousYearMonthOf,
    fetchMonthlyAverageRate, BOJ_SOURCE_LABEL,
    getExchangeRate, confirmExchangeRate,
    withSalesLock, loadSalesWorkbookLocked, atomicWriteWorkbook,
    monthSheetName, COL, numOrNull, computeOrderFinancials,
  } = deps;

  const nowYearMonth = currentYearMonth();
  const targetYearMonth = previousYearMonthOf(nowYearMonth);

  let record = await getExchangeRate({ RATE_PATH: EXCHANGE_RATE_PATH, yearMonth: targetYearMonth });
  let justConfirmed = false;

  if (!record || record["状態"] !== "確定") {
    const safety = await isPreviousMonthSafeToConfirm({ targetYearMonth, currentYearMonth: nowYearMonth, fetchMonthlyAverageRate });
    if (!safety.safe) {
      return { action: "skipped", reason: safety.reason, targetYearMonth };
    }
    const result = safety.targetResult;
    const saveResult = await confirmExchangeRate({
      RATE_PATH: EXCHANGE_RATE_PATH, yearMonth: targetYearMonth, rate: result.average, count: result.count,
      startDate: result.startDate, endDate: result.endDate, seriesCode: result.seriesCode,
      source: BOJ_SOURCE_LABEL, inputMethod: "自動(月次確定)",
    });
    if (saveResult.ok) {
      justConfirmed = true;
      record = await getExchangeRate({ RATE_PATH: EXCHANGE_RATE_PATH, yearMonth: targetYearMonth });
    } else {
      // 既に確定済み(409)等: 他経路で確定済みだった可能性があるため、現在の記録で続行を試みる。
      record = await getExchangeRate({ RATE_PATH: EXCHANGE_RATE_PATH, yearMonth: targetYearMonth });
      if (!record || record["状態"] !== "確定") {
        return { action: "skipped", reason: "confirm_failed", targetYearMonth, error: saveResult.error };
      }
    }
  }

  const rate = record["レート"];
  const sheetName = monthSheetName(targetYearMonth);

  const syncResult = await withSalesLock(async () => {
    const wb = await loadSalesWorkbookLocked();
    const ws = sheetName ? wb.getWorksheet(sheetName) : null;
    if (!ws) return { updatedCount: 0, targetCount: 0 };

    let updatedCount = 0;
    let targetCount = 0;
    for (let r = 2; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      if (!row.getCell(COL.注文番号).value) continue;
      targetCount++;
      const currentRate = numOrNull(row.getCell(COL.ドル円レート).value);
      if (currentRate === rate) continue; // 既に確定レートと一致 -> 冪等スキップ(書き換えない)

      const usd = numOrNull(row.getCell(COL.収益USD).value);
      const cost = numOrNull(row.getCell(COL.仕入原価).value);
      const shipping = numOrNull(row.getCell(COL.送料).value);
      const packing = numOrNull(row.getCell(COL.梱包費).value);
      const existingRevenueJpy = Number(row.getCell(COL.収益円).value) || 0;
      const { revenueJpy, fee, profit, margin } = computeOrderFinancials(usd, rate, existingRevenueJpy, cost, shipping, packing);

      row.getCell(COL.ドル円レート).value = rate;
      row.getCell(COL.収益円).value = Math.round(revenueJpy);
      row.getCell(COL.手数料).value = fee;
      row.getCell(COL.最終利益).value = profit;
      row.getCell(COL.利益率).value = margin;
      row.commit();
      updatedCount++;
    }
    if (updatedCount > 0) {
      await atomicWriteWorkbook(wb, LEDGER_PATH);
    }
    return { updatedCount, targetCount };
  });

  const action = justConfirmed
    ? "confirmed_and_updated"
    : (syncResult.updatedCount > 0 ? "orders_synced" : "already_up_to_date");

  return { action, targetYearMonth, rate, ...syncResult };
}

module.exports = { finalizePreviousMonthFx, isPreviousMonthSafeToConfirm };
