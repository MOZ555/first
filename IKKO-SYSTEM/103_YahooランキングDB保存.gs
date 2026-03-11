function Yahoo売買代金ランキングをDBに追記する() {
  const 取得結果 = Yahoo売買代金ランキングデータを取得する();
  const Yahoo更新日時 = 取得結果.Yahoo更新日時 || '';
  const 行一覧 = 取得結果.ランキング一覧 || [];

  if (行一覧.length === 0) {
    throw new Error('Yahoo売買代金ランキングが0件でした');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const DBシート = ss.getSheetByName(設定.ランキングDBシート名) || ss.insertSheet(設定.ランキングDBシート名);

  const 必須ヘッダー = [
    '時刻',
    'バッチID',
    'Yahoo更新日時',
    '順位',
    '銘柄コード',
    '銘柄名',
    '市場',
    '株価',
    '前日比率',
    '売買代金',
    '取得元'
  ];

  DBシートのヘッダーを整える_(DBシート, 必須ヘッダー);

  const ヘッダー = DBシート.getRange(1, 1, 1, DBシート.getLastColumn()).getValues()[0];
  const Yahoo更新日時列 = ヘッダー.indexOf('Yahoo更新日時');

  const 最新Yahoo更新日時 = 最新Yahoo更新日時を取得する_(DBシート, Yahoo更新日時列);

  if (Yahoo更新日時 && 最新Yahoo更新日時 === Yahoo更新日時) {
    Logger.log('Yahoo更新日時が前回と同じためDB追記スキップ: ' + Yahoo更新日時);

    return {
      成功: true,
      スキップ: true,
      件数: 0,
      バッチID: '',
      Yahoo更新日時: Yahoo更新日時
    };
  }

  const 現在 = new Date();
  const タイムゾーン = Session.getScriptTimeZone();
  const 取得時刻 = Utilities.formatDate(現在, タイムゾーン, 'yyyy/MM/dd HH:mm:ss');
  const バッチID = Utilities.formatDate(現在, タイムゾーン, 'yyyyMMdd_HHmmss');

  const 書き込みデータ = 行一覧.map(function(行) {
    const 行オブジェクト = {
      時刻: 取得時刻,
      バッチID: バッチID,
      Yahoo更新日時: Yahoo更新日時,
      順位: 行.順位,
      銘柄コード: 行.銘柄コード,
      銘柄名: 行.銘柄名,
      市場: 行.市場,
      株価: 行.株価,
      前日比率: 行.前日比率,
      売買代金: 行.売買代金,
      取得元: 行.取得元
    };

    return ヘッダー順の配列に変換する_(ヘッダー, 行オブジェクト);
  });

  DBシート
    .getRange(DBシート.getLastRow() + 1, 1, 書き込みデータ.length, 書き込みデータ[0].length)
    .setValues(書き込みデータ);

  Logger.log(
    'Yahoo売買代金ランキングをDBに追記しました: ' +
    書き込みデータ.length +
    '件 / バッチID=' + バッチID +
    ' / Yahoo更新日時=' + Yahoo更新日時
  );

  return {
    成功: true,
    スキップ: false,
    件数: 書き込みデータ.length,
    バッチID: バッチID,
    Yahoo更新日時: Yahoo更新日時
  };
}

function DBシートのヘッダーを整える_(シート, 必須ヘッダー) {
  const 最終列 = Math.max(シート.getLastColumn(), 1);
  const 最終行 = シート.getLastRow();

  if (最終行 === 0) {
    シート.getRange(1, 1, 1, 必須ヘッダー.length).setValues([必須ヘッダー]);
    return;
  }

  const 現在ヘッダー = シート.getRange(1, 1, 1, 最終列).getValues()[0];

  for (let i = 0; i < 必須ヘッダー.length; i++) {
    const ヘッダー名 = 必須ヘッダー[i];
    if (現在ヘッダー.indexOf(ヘッダー名) === -1) {
      const 追加列 = シート.getLastColumn() + 1;
      シート.getRange(1, 追加列).setValue(ヘッダー名);
      現在ヘッダー.push(ヘッダー名);
    }
  }
}

function 最新Yahoo更新日時を取得する_(シート, Yahoo更新日時列) {
  if (Yahoo更新日時列 === -1) return '';

  const 最終行 = シート.getLastRow();
  if (最終行 < 2) return '';

  const 値一覧 = シート
    .getRange(2, Yahoo更新日時列 + 1, 最終行 - 1, 1)
    .getValues();

  for (let i = 値一覧.length - 1; i >= 0; i--) {
    const 値 = String(値一覧[i][0] || '').trim();
    if (値) {
      return 値;
    }
  }

  return '';
}

function ヘッダー順の配列に変換する_(ヘッダー, 行オブジェクト) {
  return ヘッダー.map(function(ヘッダー名) {
    return Object.prototype.hasOwnProperty.call(行オブジェクト, ヘッダー名)
      ? 行オブジェクト[ヘッダー名]
      : '';
  });
}