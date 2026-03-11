function 当日ランキング滞在集計を作成する() {
  ランキング滞在集計を作成する_({
    DBシート名: 'ランキングDB',
    出力シート名: 'ランキング滞在集計_当日',
    集計種別: '当日'
  });
}

function 累積ランキング滞在集計を作成する() {
  ランキング滞在集計を作成する_({
    DBシート名: 'ランキングDB',
    出力シート名: 'ランキング滞在集計_累積',
    集計種別: '累積'
  });
}

function ランキング滞在集計を作成する_(設定値) {
  const DBシート名 = 設定値.DBシート名;
  const 出力シート名 = 設定値.出力シート名;
  const 集計種別 = 設定値.集計種別; // '当日' or '累積'

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const DBシート = ss.getSheetByName(DBシート名);
  if (!DBシート) throw new Error(DBシート名 + ' シートがありません');

  let 出力シート = ss.getSheetByName(出力シート名);
  if (!出力シート) {
    出力シート = ss.insertSheet(出力シート名);
  }

  const 最終行 = DBシート.getLastRow();
  const 最終列 = DBシート.getLastColumn();

  出力シート.clearContents();
  出力シート.getRange(1, 1, 1, 9).setValues([[
    '銘柄コード',
    '出現回数',
    '最高順位',
    '平均順位',
    '最新順位',
    '最新売買代金',
    '滞在スコア',
    '初回時刻',
    '最終時刻'
  ]]);

  if (最終行 < 2) {
    Logger.log(DBシート名 + ' にデータがありません');
    return;
  }

  const ヘッダー行 = DBシート.getRange(1, 1, 1, 最終列).getValues()[0];
  const データ一覧 = DBシート.getRange(2, 1, 最終行 - 1, 最終列).getValues();

  const 時刻列 = 列番号を取得する_(ヘッダー行, ['時刻', 'time', '取得時刻']);
  const 順位列 = 列番号を取得する_(ヘッダー行, ['順位', 'rank']);
  const 銘柄コード列 = 列番号を取得する_(ヘッダー行, ['銘柄コード', 'code']);
  const 売買代金列 = 列番号を取得する_(ヘッダー行, ['売買代金', 'tradingValue']);

  if (時刻列 === -1 || 順位列 === -1 || 銘柄コード列 === -1 || 売買代金列 === -1) {
    throw new Error(
      DBシート名 +
      ' の必要列が見つかりません。必要: 時刻/time, 順位/rank, 銘柄コード/code, 売買代金/tradingValue'
    );
  }

  const 今日キー = 今日の日付キーを取得する_();
  const 集計マップ = {};

  for (let i = 0; i < データ一覧.length; i++) {
    const 行 = データ一覧[i];

    const 時刻 = 行[時刻列];
    const 順位 = 数値に変換する_またはnull_(行[順位列]);
    const 銘柄コード = 集計用に銘柄コードを正規化する_(行[銘柄コード列]);
    const 売買代金 = 数値に変換する_またはnull_(行[売買代金列]);

    if (!銘柄コード) continue;
    if (順位 === null) continue;

    if (集計種別 === '当日') {
      const 行日付キー = 値から日付キーを作る_(時刻);
      if (行日付キー !== 今日キー) continue;
    }

    if (!集計マップ[銘柄コード]) {
      集計マップ[銘柄コード] = {
        銘柄コード: 銘柄コード,
        出現回数: 0,
        順位合計: 0,
        最高順位: 順位,
        最新順位: 順位,
        最新売買代金: 売買代金,
        初回時刻: 時刻,
        最終時刻: 時刻
      };
    }

    const 集計 = 集計マップ[銘柄コード];

    集計.出現回数 += 1;
    集計.順位合計 += 順位;

    if (順位 < 集計.最高順位) {
      集計.最高順位 = 順位;
    }

    if (後の時刻か判定する_(時刻, 集計.最終時刻)) {
      集計.最終時刻 = 時刻;
      集計.最新順位 = 順位;
      集計.最新売買代金 = 売買代金;
    }

    if (前の時刻か判定する_(時刻, 集計.初回時刻)) {
      集計.初回時刻 = 時刻;
    }
  }

  const 出力データ = [];

  Object.keys(集計マップ).forEach(function(銘柄コード) {
    const 集計 = 集計マップ[銘柄コード];
    const 平均順位 = 集計.順位合計 / 集計.出現回数;
    const 滞在スコア = 集計.出現回数 * (51 - 平均順位);

    出力データ.push([
      集計.銘柄コード,
      集計.出現回数,
      集計.最高順位,
      平均順位,
      集計.最新順位,
      集計.最新売買代金,
      滞在スコア,
      集計.初回時刻,
      集計.最終時刻
    ]);
  });

  出力データ.sort(function(a, b) {
    if (b[6] !== a[6]) return b[6] - a[6];
    if (b[1] !== a[1]) return b[1] - a[1];
    return a[3] - b[3];
  });

  if (出力データ.length > 0) {
    出力シート.getRange(2, 1, 出力データ.length, 9).setValues(出力データ);

    出力シート.getRange(2, 2, 出力データ.length, 1).setNumberFormat('0');
    出力シート.getRange(2, 3, 出力データ.length, 3).setNumberFormat('0.00');
    出力シート.getRange(2, 6, 出力データ.length, 1).setNumberFormat('#,##0');
    出力シート.getRange(2, 7, 出力データ.length, 1).setNumberFormat('0.00');
    出力シート.getRange(2, 8, 出力データ.length, 2).setNumberFormat('yyyy/mm/dd hh:mm:ss');
  }

  Logger.log(出力シート名 + ' 作成完了: ' + 出力データ.length + '銘柄');
}

function 今日の日付キーを取得する_() {
  const 現在 = new Date();
  return Utilities.formatDate(現在, Session.getScriptTimeZone(), 'yyyyMMdd');
}

function 値から日付キーを作る_(値) {
  if (!値) return '';

  if (値 instanceof Date) {
    return Utilities.formatDate(値, Session.getScriptTimeZone(), 'yyyyMMdd');
  }

  const 日付 = new Date(値);
  if (isNaN(日付.getTime())) return '';

  return Utilities.formatDate(日付, Session.getScriptTimeZone(), 'yyyyMMdd');
}

function 列番号を取得する_(ヘッダー行, 候補名一覧) {
  for (let i = 0; i < 候補名一覧.length; i++) {
    const 列番号 = ヘッダー行.indexOf(候補名一覧[i]);
    if (列番号 !== -1) return 列番号;
  }
  return -1;
}

function 数値に変換する_またはnull_(値) {
  if (値 === null || 値 === undefined || 値 === '') return null;

  if (typeof 値 === 'number') {
    return isFinite(値) ? 値 : null;
  }

  const 文字列 = String(値).replace(/,/g, '').trim();
  if (文字列 === '') return null;

  const 数値 = Number(文字列);
  return isFinite(数値) ? 数値 : null;
}

function 集計用に銘柄コードを正規化する_(値) {
  if (値 === null || 値 === undefined || 値 === '') return '';

  let 文字列 = String(値).trim();
  文字列 = 文字列.replace(/,/g, '');

  if (/^\d+(\.0+)?$/.test(文字列)) {
    文字列 = String(parseInt(文字列, 10));
  }

  return 文字列;
}

function 後の時刻か判定する_(a, b) {
  const 時刻A = 安全に時刻値へ変換する_(a);
  const 時刻B = 安全に時刻値へ変換する_(b);
  return 時刻A > 時刻B;
}

function 前の時刻か判定する_(a, b) {
  const 時刻A = 安全に時刻値へ変換する_(a);
  const 時刻B = 安全に時刻値へ変換する_(b);
  return 時刻A < 時刻B;
}

function 安全に時刻値へ変換する_(値) {
  if (値 instanceof Date) return 値.getTime();

  const 日付 = new Date(値);
  const 時刻値 = 日付.getTime();
  return isNaN(時刻値) ? 0 : 時刻値;
}