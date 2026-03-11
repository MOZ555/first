function 日次終了処理を実行する() {
  const 開始時刻 = new Date();

  try {
    Logger.log('=== 日次終了処理 開始 ===');

    // 1. 最終監視シートを最新状態で再作成
    最終監視シートを作成する();

    // 2. 最終監視シートを履歴保存
    最終監視シートを履歴保存する();

    // 3. 実行ログ保存
    日次終了処理ログを保存する_({
      ステータス: '成功',
      メッセージ: '最終監視シート再作成・履歴保存 完了',
      開始時刻: 開始時刻,
      終了時刻: new Date()
    });

    Logger.log('=== 日次終了処理 正常終了 ===');
  } catch (e) {
    日次終了処理ログを保存する_({
      ステータス: '失敗',
      メッセージ: e.message || String(e),
      開始時刻: 開始時刻,
      終了時刻: new Date()
    });

    Logger.log('=== 日次終了処理 失敗 === ' + (e.message || e));
    throw e;
  }
}

function 最終監視シートを履歴保存する() {
  const SRC_SHEET = '最終監視シート';
  const DST_SHEET = '最終監視シート_履歴';

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const srcSheet = ss.getSheetByName(SRC_SHEET);
  if (!srcSheet) throw new Error('最終監視シートがありません');

  let dstSheet = ss.getSheetByName(DST_SHEET);
  if (!dstSheet) {
    dstSheet = ss.insertSheet(DST_SHEET);
  }

  const lastRow = srcSheet.getLastRow();
  const lastCol = srcSheet.getLastColumn();

  if (lastRow < 2) {
    Logger.log('保存するデータなし');
    return;
  }

  const header = srcSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const data = srcSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const col監視順位 = header.indexOf('監視順位');
  const col銘柄コード = header.indexOf('銘柄コード');
  const col銘柄名 = header.indexOf('銘柄名');
  const col市場 = header.indexOf('市場');
  const colテーマ = header.indexOf('テーマ');
  const col出所 = header.indexOf('出所');
  const col候補区分 = header.indexOf('候補区分');
  const col前日比 = header.indexOf('前日比');
  const col除外 = header.indexOf('除外');
  const col最終スコア = header.indexOf('最終スコア');
  const col資金流入変化 = header.indexOf('資金流入変化');
  const col累積スコア = header.indexOf('累積スコア');
  const col当日スコア = header.indexOf('当日スコア');
  const col累積メモ = header.indexOf('累積メモ');
  const col当日メモ = header.indexOf('当日メモ');

  if (
    col監視順位 === -1 ||
    col銘柄コード === -1 ||
    col銘柄名 === -1 ||
    col市場 === -1 ||
    colテーマ === -1 ||
    col出所 === -1 ||
    col候補区分 === -1 ||
    col前日比 === -1 ||
    col除外 === -1 ||
    col最終スコア === -1 ||
    col資金流入変化 === -1 ||
    col累積スコア === -1 ||
    col当日スコア === -1 ||
    col累積メモ === -1 ||
    col当日メモ === -1
  ) {
    throw new Error('最終監視シートの必要列が見つかりません');
  }

  const today = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyy/MM/dd'
  );

  const output = [];

  for (let i = 0; i < data.length; i++) {
    const row = data[i];

    output.push([
      today,
      row[col監視順位],
      row[col銘柄コード],
      row[col銘柄名],
      row[col市場],
      row[colテーマ],
      row[col出所],
      row[col候補区分],
      row[col前日比],
      row[col除外],
      row[col最終スコア],
      row[col資金流入変化],
      row[col累積スコア],
      row[col当日スコア],
      row[col累積メモ],
      row[col当日メモ]
    ]);
  }

  let dstLastRow = dstSheet.getLastRow();

  if (dstLastRow === 0) {
    dstSheet.getRange(1, 1, 1, 16).setValues([[
      '日付',
      '監視順位',
      '銘柄コード',
      '銘柄名',
      '市場',
      'テーマ',
      '出所',
      '候補区分',
      '前日比',
      '除外',
      '最終スコア',
      '資金流入変化',
      '累積スコア',
      '当日スコア',
      '累積メモ',
      '当日メモ'
    ]]);
    dstLastRow = 1;
  }

  dstSheet.getRange(dstLastRow + 1, 1, output.length, 16).setValues(output);

  dstSheet.getRange(dstLastRow + 1, 2, output.length, 1).setNumberFormat('0');     // 監視順位
  dstSheet.getRange(dstLastRow + 1, 9, output.length, 1).setNumberFormat('0.00');   // 前日比
  dstSheet.getRange(dstLastRow + 1, 11, output.length, 1).setNumberFormat('0.00');  // 最終スコア
  dstSheet.getRange(dstLastRow + 1, 13, output.length, 2).setNumberFormat('0.00');  // 累積/当日スコア

  dstSheet.setFrozenRows(1);

  Logger.log('最終監視履歴 保存: ' + output.length + '件');
}

function 日次終了処理ログを保存する_(params) {
  const LOG_SHEET = '日次終了処理ログ';

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let logSheet = ss.getSheetByName(LOG_SHEET);
  if (!logSheet) {
    logSheet = ss.insertSheet(LOG_SHEET);
    logSheet.getRange(1, 1, 1, 6).setValues([[
      '実行日',
      '開始時刻',
      '終了時刻',
      '処理秒数',
      'ステータス',
      'メッセージ'
    ]]);
    logSheet.setFrozenRows(1);
  }

  const 開始時刻 = params.開始時刻 || new Date();
  const 終了時刻 = params.終了時刻 || new Date();
  const ステータス = params.ステータス || '';
  const メッセージ = params.メッセージ || '';

  const 実行日 = Utilities.formatDate(
    開始時刻,
    Session.getScriptTimeZone(),
    'yyyy/MM/dd'
  );

  const 開始時刻文字列 = Utilities.formatDate(
    開始時刻,
    Session.getScriptTimeZone(),
    'yyyy/MM/dd HH:mm:ss'
  );

  const 終了時刻文字列 = Utilities.formatDate(
    終了時刻,
    Session.getScriptTimeZone(),
    'yyyy/MM/dd HH:mm:ss'
  );

  const 処理秒数 = ((終了時刻.getTime() - 開始時刻.getTime()) / 1000);

  logSheet.appendRow([
    実行日,
    開始時刻文字列,
    終了時刻文字列,
    処理秒数,
    ステータス,
    メッセージ
  ]);
}