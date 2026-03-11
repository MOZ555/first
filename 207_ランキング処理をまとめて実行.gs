/*************************************************
 * 101_Yahooランキング取得.gs
 *************************************************/

/**
 * Yahoo売買代金ランキングを取得して、扱いやすい配列で返す
 * 返り値:
 * [
 *   {
 *     rank: 1,
 *     code: "285A",
 *     name: "キオクシアホールディングス(株)",
 *     market: "東証PRM",
 *     price: 18055,
 *     changeRate: -9.61,
 *     tradingValue: 54179500500,
 *     yahooUpdateTime: "2026/03/09 09:36"
 *   },
 *   ...
 * ]
 */
function Yahoo売買代金ランキングデータを取得する() {
  const url = 'https://finance.yahoo.co.jp/stocks/ranking/tradingValueHigh?market=all';

  const response = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    headers: {
      'User-Agent': 'Mozilla/5.0'
    }
  });

  const code = response.getResponseCode();
  const html = response.getContentText();

  Logger.log('HTTP code: ' + code);
  Logger.log('HTML length: ' + html.length);

  if (code !== 200) {
    throw new Error('Yahooページ取得失敗 HTTP code=' + code);
  }

  const startKey = 'window.__PRELOADED_STATE__ = ';
  const start = html.indexOf(startKey);

  if (start === -1) {
    throw new Error('PRELOADED_STATE開始位置が見つかりません');
  }

  const end = html.indexOf('</script>', start);
  if (end === -1) {
    throw new Error('PRELOADED_STATE終了位置が見つかりません');
  }

  const jsonText = html
    .substring(start + startKey.length, end)
    .trim()
    .replace(/;$/, '');

  const json = JSON.parse(jsonText);

  const ranking = json.mainRankingList?.results || [];

  Logger.log('Yahoo取得件数: ' + ranking.length);

  return ranking.map((r, i) => ({
    rank: Number(r.rank || (i + 1)),
    code: r.stockCode || '',
    name: r.stockName || '',
    market: r.marketName || '',
    price: toNumber_(r.savePrice),
    changeRate: toNumber_(r.rankingResult?.tradingValue?.changePriceRate),
    tradingValue: toNumber_(r.rankingResult?.tradingValue?.tradingValue),
    yahooUpdateTime: normalizeYahooDateTime_(r.rankingResult?.tradingValue?.updateDateTime || '')
  }));
}

/**
 * 互換用ラッパー
 */
function Yahoo売買代金ランキングを取得する() {
  return Yahoo売買代金ランキングデータを取得する();
}

/**
 * "18,055" や "-9.61" を数値化
 */
function toNumber_(value) {
  if (value === null || value === undefined || value === '') return '';
  const num = Number(String(value).replace(/,/g, '').trim());
  return isNaN(num) ? '' : num;
}

/**
 * Yahooの更新日時を正規化
 * 例:
 * "2026/03/09 09:36" -> "2026/03/09 09:36:00"
 * "03/09 09:36"      -> "2026/03/09 09:36:00"
 */
function normalizeYahooDateTime_(value) {
  if (!value) return '';

  let s = String(value).trim();

  // すでに yyyy/MM/dd HH:mm:ss
  if (/^\d{4}\/\d{2}\/\d{2}\s+\d{2}:\d{2}:\d{2}$/.test(s)) {
    return s;
  }

  // yyyy/MM/dd HH:mm
  if (/^\d{4}\/\d{2}\/\d{2}\s+\d{2}:\d{2}$/.test(s)) {
    return s + ':00';
  }

  // MM/dd HH:mm → 当年補完
  if (/^\d{2}\/\d{2}\s+\d{2}:\d{2}$/.test(s)) {
    const year = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy');
    return year + '/' + s + ':00';
  }

  return s;
}


/*************************************************
 * 103_YahooランキングDB保存.gs
 *************************************************/

/**
 * 取得済みデータをランキングDBへ追記
 * 重複判定キー:
 * Yahoo更新日時 + 順位 + 銘柄コード
 */
function Yahoo売買代金ランキングをDBに追記する(ランキングデータ) {
  if (!ランキングデータ || ランキングデータ.length === 0) {
    return { 件数: 0 };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('ランキングDB');
  if (!sheet) {
    throw new Error('シート「ランキングDB」が見つかりません');
  }

  const now = new Date();
  const tz = Session.getScriptTimeZone();
  const batchId = Utilities.formatDate(now, tz, 'yyyyMMdd_HHmmss');
  const 取得元 = 'yahoo';

  const lastRow = sheet.getLastRow();

  // 既存キー読み込み
  const existingKeys = new Set();
  if (lastRow > 1) {
    const existingData = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    // A:時刻 B:バッチID C:Yahoo更新日時 D:順位 E:銘柄コード F:銘柄名
    existingData.forEach(row => {
      const yahooUpdateTime = row[2];
      const rank = row[3];
      const code = row[4];
      const key = [String(yahooUpdateTime), String(rank), String(code)].join('|');
      existingKeys.add(key);
    });
  }

  const rows = [];
  for (const r of ランキングデータ) {
    const key = [String(r.yahooUpdateTime), String(r.rank), String(r.code)].join('|');
    if (existingKeys.has(key)) continue;

    rows.push([
      now,                 // 時刻
      batchId,             // バッチID
      r.yahooUpdateTime,   // Yahoo更新日時
      r.rank,              // 順位
      r.code,              // 銘柄コード
      r.name,              // 銘柄名
      r.market,            // 市場
      r.price,             // 株価
      r.changeRate,        // 前日比率
      r.tradingValue,      // 売買代金
      取得元               // 取得元
    ]);
  }

  if (rows.length === 0) {
    return { 件数: 0 };
  }

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);

  return { 件数: rows.length };
}


/*************************************************
 * 206_ランキング処理をまとめて実行.gs
 *************************************************/

function ランキング処理をまとめて実行する() {
  if (!市場時間内か判定する_()) {
    Logger.log('市場時間外のため処理スキップ');
    return {
      開始時刻: new Date(),
      Yahoo取得件数: 0,
      DB追記件数: 0,
      当日集計件数: 0,
      累積集計件数: 0,
      累積候補件数: 0,
      当日候補件数: 0,
      成功: true,
      スキップ: true,
      エラー: ''
    };
  }

  const 開始時刻 = new Date();
  const タイムゾーン = Session.getScriptTimeZone();

  Logger.log('=== ランキング処理 開始 ===');
  Logger.log('開始: ' + Utilities.formatDate(開始時刻, タイムゾーン, 'yyyy/MM/dd HH:mm:ss'));

  const 実行結果 = {
    開始時刻: 開始時刻,
    Yahoo取得件数: 0,
    DB追記件数: 0,
    当日集計件数: 0,
    累積集計件数: 0,
    累積候補件数: 0,
    当日候補件数: 0,
    成功: false,
    エラー: ''
  };

  try {
    const 取得データ = Yahoo売買代金ランキングを取得する();
    実行結果.Yahoo取得件数 = 取得データ.length;
    Logger.log('Yahoo取得件数: ' + 取得データ.length);

    if (!取得データ || 取得データ.length === 0) {
      throw new Error('Yahoo売買代金ランキングの取得件数が0件でした');
    }

    // 取得済みデータをそのまま渡す
    const DB結果 = Yahoo売買代金ランキングをDBに追記する(取得データ);
    実行結果.DB追記件数 = DB結果.件数 || 0;
    Logger.log('DB追記件数: ' + 実行結果.DB追記件数);

    実行結果.当日集計件数 = 当日ランキング滞在集計を作成して件数を返す_();
    Logger.log('当日集計件数: ' + 実行結果.当日集計件数);

    実行結果.累積集計件数 = 累積ランキング滞在集計を作成して件数を返す_();
    Logger.log('累積集計件数: ' + 実行結果.累積集計件数);

    実行結果.累積候補件数 = 累積候補抽出を作成して件数を返す_();
    Logger.log('累積候補件数: ' + 実行結果.累積候補件数);

    実行結果.当日候補件数 = 当日候補抽出を作成して件数を返す_();
    Logger.log('当日候補件数: ' + 実行結果.当日候補件数);

    実行結果.成功 = true;

    const 終了時刻 = new Date();
    Logger.log('終了: ' + Utilities.formatDate(終了時刻, タイムゾーン, 'yyyy/MM/dd HH:mm:ss'));
    Logger.log('=== ランキング処理 正常終了 ===');

  } catch (e) {
    実行結果.エラー = String(e);
    Logger.log('エラー: ' + String(e));
    Logger.log('=== ランキング処理 異常終了 ===');
    throw e;
  }

  return 実行結果;
}

function 当日ランキング滞在集計を作成して件数を返す_() {
  当日ランキング滞在集計を作成する();
  return シートデータ件数を返す_('ランキング滞在集計_当日');
}

function 累積ランキング滞在集計を作成して件数を返す_() {
  累積ランキング滞在集計を作成する();
  return シートデータ件数を返す_('ランキング滞在集計_累積');
}

function 累積候補抽出を作成して件数を返す_() {
  ランキング候補抽出を作成する();
  return シートデータ件数を返す_('ランキング候補抽出');
}

function 当日候補抽出を作成して件数を返す_() {
  当日ランキング候補抽出を作成する();
  return シートデータ件数を返す_('ランキング候補抽出_当日');
}

function シートデータ件数を返す_(シート名) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const シート = ss.getSheetByName(シート名);
  if (!シート) return 0;
  return Math.max(0, シート.getLastRow() - 1);
}