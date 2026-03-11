function ランキング候補抽出を作成する() {
  候補抽出シートを作成する_({
    集計元シート名: 'ランキング滞在集計_累積',
    DBシート名: 'ランキングDB',
    出力シート名: 'ランキング候補抽出'
  });
}

function 候補抽出シートを作成する_(設定値) {
  const 集計元シート名 = 設定値.集計元シート名;
  const DBシート名 = 設定値.DBシート名;
  const 出力シート名 = 設定値.出力シート名;

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const 集計元シート = ss.getSheetByName(集計元シート名);
  if (!集計元シート) throw new Error(集計元シート名 + ' シートがありません');

  const DBシート = ss.getSheetByName(DBシート名);
  if (!DBシート) throw new Error(DBシート名 + ' シートがありません');

  let 出力シート = ss.getSheetByName(出力シート名);
  if (!出力シート) {
    出力シート = ss.insertSheet(出力シート名);
  }

  const 集計元最終行 = 集計元シート.getLastRow();
  const 集計元最終列 = 集計元シート.getLastColumn();

  出力シート.clearContents();
  出力シート.getRange(1, 1, 1, 13).setValues([[
    '候補順位',
    '銘柄コード',
    '候補区分',
    'バッチ数',
    '出現回数',
    '出現率',
    '最高順位',
    '平均順位',
    '最新順位',
    '最新売買代金',
    '滞在スコア',
    '候補スコア',
    'メモ'
  ]]);

  if (集計元最終行 < 2) {
    Logger.log(集計元シート名 + ' にデータがありません');
    return;
  }

  const ヘッダー = 集計元シート.getRange(1, 1, 1, 集計元最終列).getValues()[0];
  const 値一覧 = 集計元シート.getRange(2, 1, 集計元最終行 - 1, 集計元最終列).getValues();

  const 銘柄コード列 = 列番号を取得する_(ヘッダー, ['銘柄コード', 'code']);
  const 出現回数列 = 列番号を取得する_(ヘッダー, ['出現回数', 'count']);
  const 最高順位列 = 列番号を取得する_(ヘッダー, ['最高順位', 'bestRank']);
  const 平均順位列 = 列番号を取得する_(ヘッダー, ['平均順位', 'avgRank']);
  const 最新順位列 = 列番号を取得する_(ヘッダー, ['最新順位', 'latestRank']);
  const 最新売買代金列 = 列番号を取得する_(ヘッダー, ['最新売買代金', 'latestTradingValue']);
  const 滞在スコア列 = 列番号を取得する_(ヘッダー, ['滞在スコア', 'stayScore']);

  if (
    銘柄コード列 === -1 ||
    出現回数列 === -1 ||
    最高順位列 === -1 ||
    平均順位列 === -1 ||
    最新順位列 === -1 ||
    最新売買代金列 === -1 ||
    滞在スコア列 === -1
  ) {
    throw new Error(集計元シート名 + ' の必要列が見つかりません');
  }

  const バッチ数 = 一意なバッチ数を数える_(DBシート);
  if (バッチ数 === 0) {
    Logger.log(DBシート名 + ' にバッチIDがありません');
    return;
  }

  const 出力データ = [];

  for (let i = 0; i < 値一覧.length; i++) {
    const 行 = 値一覧[i];

    const 銘柄コード = 候補抽出用に銘柄コードを正規化する_(行[銘柄コード列]);
    const 出現回数 = 数値に変換する_またはnull_(行[出現回数列]);
    const 最高順位 = 数値に変換する_またはnull_(行[最高順位列]);
    const 平均順位 = 数値に変換する_またはnull_(行[平均順位列]);
    const 最新順位 = 数値に変換する_またはnull_(行[最新順位列]);
    const 最新売買代金 = 数値に変換する_またはnull_(行[最新売買代金列]);
    const 滞在スコア = 数値に変換する_またはnull_(行[滞在スコア列]);

    if (!銘柄コード) continue;
    if (出現回数 === null || 最高順位 === null || 平均順位 === null || 最新順位 === null) continue;

    const 出現率 = 出現回数 / バッチ数;

    let 加点 = 0;
    if (最高順位 <= 10) 加点 += 5;
    if (最新順位 <= 10) 加点 += 5;
    if (最新順位 < 平均順位) 加点 += 3;
    if (出現率 >= 0.8) 加点 += 5;
    else if (出現率 >= 0.6) 加点 += 3;

    const 候補スコア =
      (出現率 * 100) +
      (51 - 平均順位) +
      (51 - 最新順位) +
      加点;

    const 候補区分 = 候補区分を判定する_(出現率, 平均順位, 最新順位);
    if (!候補区分) continue;

    const メモ = 候補メモを作成する_(候補区分, 出現率, 平均順位, 最新順位);

    出力データ.push([
      '',
      銘柄コード,
      候補区分,
      バッチ数,
      出現回数,
      出現率,
      最高順位,
      平均順位,
      最新順位,
      最新売買代金,
      滞在スコア,
      候補スコア,
      メモ
    ]);
  }

  出力データ.sort(function(a, b) {
    if (b[11] !== a[11]) return b[11] - a[11];
    if (b[5] !== a[5]) return b[5] - a[5];
    return a[7] - b[7];
  });

  for (let i = 0; i < 出力データ.length; i++) {
    出力データ[i][0] = i + 1;
  }

  if (出力データ.length > 0) {
    出力シート.getRange(2, 1, 出力データ.length, 13).setValues(出力データ);

    出力シート.getRange(2, 1, 出力データ.length, 1).setNumberFormat('0');
    出力シート.getRange(2, 4, 出力データ.length, 2).setNumberFormat('0');
    出力シート.getRange(2, 6, 出力データ.length, 1).setNumberFormat('0.00%');
    出力シート.getRange(2, 7, 出力データ.length, 3).setNumberFormat('0.00');
    出力シート.getRange(2, 10, 出力データ.length, 1).setNumberFormat('#,##0');
    出力シート.getRange(2, 11, 出力データ.length, 2).setNumberFormat('0.00');
  }

  Logger.log(出力シート名 + ' 作成完了: ' + 出力データ.length + '件');
}

function 一意なバッチ数を数える_(シート) {
  const 最終行 = シート.getLastRow();
  const 最終列 = シート.getLastColumn();
  if (最終行 < 2) return 0;

  const ヘッダー = シート.getRange(1, 1, 1, 最終列).getValues()[0];
  const バッチID列 = 列番号を取得する_(ヘッダー, ['バッチID', 'batchId']);
  if (バッチID列 === -1) return 0;

  const 値一覧 = シート.getRange(2, バッチID列 + 1, 最終行 - 1, 1).getValues();
  const 見たバッチ = {};

  for (let i = 0; i < 値一覧.length; i++) {
    const 値 = String(値一覧[i][0] || '').trim();
    if (!値) continue;
    見たバッチ[値] = true;
  }

  return Object.keys(見たバッチ).length;
}

function 候補区分を判定する_(出現率, 平均順位, 最新順位) {
  if (出現率 >= 0.6 && 平均順位 <= 15 && 最新順位 <= 15) return 'A';
  if (出現率 >= 0.4 && 平均順位 <= 25 && 最新順位 <= 25) return 'B';
  if (出現率 >= 0.2 && 最新順位 <= 15) return 'C';
  return '';
}

function 候補メモを作成する_(候補区分, 出現率, 平均順位, 最新順位) {
  if (候補区分 === 'A') return '高滞在率。上位継続';
  if (候補区分 === 'B') return '中滞在率。監視候補';
  if (候補区分 === 'C') return '単発〜中滞在。様子見';
  return '';
}

function 候補抽出用に銘柄コードを正規化する_(値) {
  if (値 === null || 値 === undefined || 値 === '') return '';

  let 文字列 = String(値).trim();
  文字列 = 文字列.replace(/,/g, '');

  if (/^\d+(\.0+)?$/.test(文字列)) {
    文字列 = String(parseInt(文字列, 10));
  }

  return 文字列;
}