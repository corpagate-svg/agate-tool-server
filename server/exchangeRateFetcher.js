// 日本銀行 時系列統計データ検索サイトAPIから、東京市場ドル・円スポット中心相場(FXERD05)の
// 日次データを取得し、月次平均を計算する。Excelの読み書きは一切行わない(exchangeRateStore.jsの責務)。
// ネットワーク呼び出しはExcelロックの外側で行う前提のモジュールのため、ここではロックを持たない。
// テスト用: 実際の日銀APIへアクセスせずにHTTP統合テストを行うためのオーバーライド。
// 本番運用ではAGATE_BOJ_API_BASE_OVERRIDEを設定しないため、常に日銀の正式エンドポイントを使う。
const BOJ_API_BASE = process.env.AGATE_BOJ_API_BASE_OVERRIDE || "https://www.stat-search.boj.or.jp/api/v1/getDataCode";
const BOJ_DB = "FM08";
const BOJ_SERIES_CODE = "FXERD05";
const BOJ_SOURCE_LABEL = "日本銀行";
const DEFAULT_TIMEOUT_MS = 10000;

function toBojYearMonth(yearMonth) {
  return String(yearMonth || "").replace("-", "");
}

function buildRequestUrl(yearMonth) {
  const ym = toBojYearMonth(yearMonth);
  const params = new URLSearchParams({
    format: "json", lang: "jp", db: BOJ_DB, code: BOJ_SERIES_CODE,
    startDate: ym, endDate: ym,
  });
  return `${BOJ_API_BASE}?${params.toString()}`;
}

// YYYYMMDD(数値または文字列) -> YYYY-MM-DD。形式が想定外ならnull。
function normalizeDate(raw) {
  const s = String(raw);
  if (!/^\d{8}$/.test(s)) return null;
  return `${s.slice(0, 4)}-${s.slice(4, 6)}-${s.slice(6, 8)}`;
}

async function fetchWithTimeout(url, { timeoutMs = DEFAULT_TIMEOUT_MS, fetchImpl = fetch } = {}) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    return await fetchImpl(url, { signal: controller.signal });
  } finally {
    clearTimeout(timer);
  }
}

// 日銀APIから対象年月の日次データを取得し、正規化・重複排除・異常値除外まで行う。
// 戻り値: { ok: true, dailyRates: [{date,value}], excludedCount, seriesCode, dbName }
//      または { ok: false, error, errorType }
async function fetchDailyRates(yearMonth, { fetchImpl = fetch, timeoutMs = DEFAULT_TIMEOUT_MS } = {}) {
  const url = buildRequestUrl(yearMonth);
  let res;
  try {
    res = await fetchWithTimeout(url, { timeoutMs, fetchImpl });
  } catch (e) {
    if (e && e.name === "AbortError") {
      return { ok: false, error: "日銀APIへの接続がタイムアウトしました", errorType: "timeout" };
    }
    return { ok: false, error: `日銀APIへの接続に失敗しました: ${e.message}`, errorType: "network" };
  }

  let text;
  try {
    text = await res.text();
  } catch (e) {
    return { ok: false, error: "日銀APIの応答を読み取れませんでした", errorType: "network" };
  }

  if (!res.ok) {
    return { ok: false, error: `日銀APIがHTTPエラーを返しました(${res.status})`, errorType: "http", httpStatus: res.status };
  }

  let json;
  try {
    json = JSON.parse(text);
  } catch (e) {
    return { ok: false, error: "日銀APIの応答形式が想定と異なります(JSON解析に失敗しました)", errorType: "format" };
  }
  if (!json || typeof json !== "object") {
    return { ok: false, error: "日銀APIの応答形式が想定と異なります", errorType: "format" };
  }
  if (json.STATUS !== 200) {
    return {
      ok: false,
      error: `日銀APIがエラーを返しました: ${json.MESSAGE || "STATUS=" + json.STATUS}`,
      errorType: "boj_error",
      bojStatus: json.STATUS,
    };
  }

  const resultSet = Array.isArray(json.RESULTSET) ? json.RESULTSET : [];
  const series = resultSet.find((s) => s && s.SERIES_CODE === BOJ_SERIES_CODE) || resultSet[0];
  if (!series || !series.VALUES) {
    // 「正常に終了しましたが、該当データはありませんでした」(M181030I)等、系列は空でも200になる場合がある。
    return { ok: true, dailyRates: [], excludedCount: 0, seriesCode: BOJ_SERIES_CODE, dbName: BOJ_DB };
  }

  const dates = Array.isArray(series.VALUES.SURVEY_DATES) ? series.VALUES.SURVEY_DATES : [];
  const values = Array.isArray(series.VALUES.VALUES) ? series.VALUES.VALUES : [];

  // 同一日付の重複は1件として扱う(後に出現した値で上書き。通常のAPI応答では発生しない想定)。
  const byDate = new Map();
  let excludedCount = 0;
  for (let i = 0; i < dates.length; i++) {
    const dateStr = normalizeDate(dates[i]);
    const rawValue = values[i];
    if (!dateStr) { excludedCount++; continue; }
    const num = typeof rawValue === "number" ? rawValue : null;
    // nullは欠測値(BOJ側の仕様どおり)。数値でも0以下・非有限は不正値として除外する。
    if (num === null || !Number.isFinite(num) || num <= 0) { excludedCount++; continue; }
    byDate.set(dateStr, num);
  }

  const dailyRates = Array.from(byDate.entries())
    .map(([date, value]) => ({ date, value }))
    .sort((a, b) => a.date.localeCompare(b.date));

  return { ok: true, dailyRates, excludedCount, seriesCode: BOJ_SERIES_CODE, dbName: BOJ_DB };
}

// 日次データの単純算術平均。空の場合は average=null。
// 日次値は一切丸めずに合計・平均を計算し、最終的な月次平均だけを小数第2位に丸めて
// 正式な適用レートとする(実日銀APIの確認結果を踏まえた仕様。丸め位置を誤ると
// 収益円・Payoneer手数料・最終利益の計算結果がずれるため、丸めるのはここ1箇所だけにする)。
function computeMonthlyAverage(dailyRates) {
  if (!dailyRates || !dailyRates.length) return { average: null, count: 0, startDate: null, endDate: null };
  const sum = dailyRates.reduce((acc, r) => acc + r.value, 0);
  const rawAverage = sum / dailyRates.length;
  const average = Math.round(rawAverage * 100) / 100;
  return {
    average, count: dailyRates.length,
    startDate: dailyRates[0].date, endDate: dailyRates[dailyRates.length - 1].date,
  };
}

// fetchDailyRates + computeMonthlyAverage をまとめた便利関数。0件の場合はok:falseにする
// (呼び出し側が「平均を保存できない」ケースと「平均が正常に0件だった」ケースを区別しやすくするため)。
async function fetchMonthlyAverageRate(yearMonth, opts = {}) {
  const result = await fetchDailyRates(yearMonth, opts);
  if (!result.ok) return result;
  const { average, count, startDate, endDate } = computeMonthlyAverage(result.dailyRates);
  if (count === 0) {
    return { ok: false, error: `${yearMonth}の日銀公表データが取得できませんでした(0件)`, errorType: "no_data" };
  }
  return {
    ok: true,
    average, count, startDate, endDate,
    seriesCode: result.seriesCode, dbName: result.dbName, excludedCount: result.excludedCount,
    dailyRates: result.dailyRates,
  };
}

module.exports = {
  BOJ_API_BASE, BOJ_DB, BOJ_SERIES_CODE, BOJ_SOURCE_LABEL,
  buildRequestUrl, normalizeDate,
  fetchDailyRates, computeMonthlyAverage, fetchMonthlyAverageRate,
};
