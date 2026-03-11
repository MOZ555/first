/*************************************************
 * 204_最終監視シート作成.gs
 *************************************************/

/**
 * 最終監視シートを作成する
 *
 * 入力:
 * - ランキング候補抽出
 * - ランキング候補抽出_当日
 * - ランキングDB
 * - テーママスタ
 *
 * 出力:
 * - 最終監視シート
 */
function 最終監視シートを作成する() {
  const 累積候補シート名 = 'ランキング候補抽出';
  const 当日候補シート名 = 'ランキング候補抽出_当日';
  const ランキングDBシート名 = 'ランキングDB';
  const テーママスタシート名 = 'テーママスタ';
  const 出力シート名 = '最終監視シート';

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const 累積候補シート = ss.getSheetByName(累積候補シート名);
  if (!累積候補シート) {
    throw new Error('シート「' + 累積候補シート名 + '」が見つかりません');
  }

  const 当日候補シート = ss.getSheetByName(当日候補シート名); // 無くてもOK

  const ランキングDBシート = ss.getSheetByName(ランキングDBシート名);
  if (!ランキングDBシート) {
    throw new Error('シート「' + ランキングDBシート名 + '」が見つかりません');
  }

  const テーママスタシート = ss.getSheetByName(テーママスタシート名); // 無くてもOK

  let 出力シート = ss.getSheetByName(出力シート名);
  if (!出力シート) {
    出力シート = ss.insertSheet(出力シート名);
  }

  const 出力列数 = 16;

  // ===== 2行目以降だけ消す（条件付き書式は消さない）=====
  const 最終行 = 出力シート.getLastRow();
  const 最終列 = Math.max(出力シート.getLastColumn(), 出力列数);

  if (最終行 >= 2) {
    出力シート.getRange(2, 1, 最終行 - 1, 最終列).clearContent();
  }

  // ===== ヘッダー =====
  出力シート.getRange(1, 1, 1, 出力列数).setValues([[
    '監視順位',
    '注目',
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

  const 累積候補マップ = 候補シートをコード別マップに変換する_(累積候補シート);
  const 当日候補マップ = 当日候補シート
    ? 候補シートをコード別マップに変換する_(当日候補シート)
    : {};

  const ランキング最新情報マップ = ランキングDBをコード別最新情報マップに変換する_(ランキングDBシート);
  const テーママップ = テーママスタシート
    ? テーママスタをコード別マップに変換する_(テーママスタシート)
    : {};

  const 全銘柄コード一覧オブジェクト = {};

  Object.keys(累積候補マップ).forEach(function(銘柄コード) {
    全銘柄コード一覧オブジェクト[銘柄コード] = true;
  });

  Object.keys(当日候補マップ).forEach(function(銘柄コード) {
    全銘柄コード一覧オブジェクト[銘柄コード] = true;
  });

  const 全銘柄コード一覧 = Object.keys(全銘柄コード一覧オブジェクト);
  const 出力データ = [];

  for (let i = 0; i < 全銘柄コード一覧.length; i++) {
    const 銘柄コード = 全銘柄コード一覧[i];
    const 累積候補 = 累積候補マップ[銘柄コード] || null;
    const 当日候補 = 当日候補マップ[銘柄コード] || null;
    const 最新情報 = ランキング最新情報マップ[銘柄コード] || {};
    const テーマ情報 = テーママップ[銘柄コード] || {};

    const 銘柄名 = 最新情報.銘柄名 || '';
    const 市場 = 最新情報.市場 || '';
    const 前日比 = 最新情報.前日比;
    const テーマ = テーマ情報.テーマ || '';

    const 出所 = 最終出所を判定する_(累積候補, 当日候補);
    const 資金流入変化 = 資金流入変化を判定する_(累積候補, 当日候補);
    const 候補区分 = 最終候補区分を判定する_(累積候補, 当日候補);
    const 除外 = 除外判定を返す_(前日比);
    const 最終スコア = 最終スコアを計算する_(累積候補, 当日候補, 前日比);

    const 累積スコア = 累積候補 ? 累積候補.候補スコア : '';
    const 当日スコア = 当日候補 ? 当日候補.候補スコア : '';
    const 累積メモ = 累積候補 ? 累積候補.メモ : '';
    const 当日メモ = 当日候補 ? 当日候補.メモ : '';

    const 注目 = 注目記号を返す_(累積候補, 当日候補, 前日比, 除外);

    出力データ.push([
      '', // 監視順位
      注目,
      銘柄コード,
      銘柄名,
      市場,
      テーマ,
      出所,
      候補区分,
      前日比 === null || 前日比 === undefined ? '' : 前日比,
      除外,
      最終スコア,
      資金流入変化,
      累積スコア,
      当日スコア,
      累積メモ,
      当日メモ
    ]);
  }

  出力データ.sort(function(a, b) {
    // 除外は一番下へ
    const 除外A = String(a[9] || '');
    const 除外B = String(b[9] || '');
    if (除外A !== 除外B) {
      if (除外A === '除外') return 1;
      if (除外B === '除外') return -1;
    }

    // 最終スコア高い順
    const スコアA = 数値に変換する_またはゼロ_(a[10]);
    const スコアB = 数値に変換する_またはゼロ_(b[10]);
    if (スコアB !== スコアA) return スコアB - スコアA;

    // 候補区分 A > B > C
    const 候補区分重みA = 候補区分の重みを返す_(a[7]);
    const 候補区分重みB = 候補区分の重みを返す_(b[7]);
    if (候補区分重みB !== 候補区分重みA) return 候補区分重みB - 候補区分重みA;

    // 最後はコード順
    return String(a[2]).localeCompare(String(b[2]));
  });

  for (let i = 0; i < 出力データ.length; i++) {
    出力データ[i][0] = i + 1;
  }

  if (出力データ.length > 0) {
    出力シート.getRange(2, 1, 出力データ.length, 出力列数).setValues(出力データ);
  
    最終監視シートの銘柄コードにリンクを設定する_(出力シート, 2, 出力データ.length);

    出力シート.getRange(2, 1, 出力データ.length, 1).setNumberFormat('0');      // 監視順位
    出力シート.getRange(2, 9, 出力データ.length, 1).setNumberFormat('0.00');     // 前日比
    出力シート.getRange(2, 11, 出力データ.length, 1).setNumberFormat('0.00');    // 最終スコア
    出力シート.getRange(2, 13, 出力データ.length, 2).setNumberFormat('0.00');    // 累積/当日スコア
  }

  出力シート.setFrozenRows(1);
  出力シート.autoResizeColumns(1, 出力列数);

  最終監視シートの条件付き書式を設定する_();

  Logger.log('最終監視シート 作成完了: ' + 出力データ.length + '件');
}

/**
 * 候補シートをコード別マップに変換
 * 必要列:
 * - 銘柄コード
 * - 候補区分
 * - 候補スコア
 * - メモ
 */
function 候補シートをコード別マップに変換する_(シート) {
  const 最終行 = シート.getLastRow();
  const 最終列 = シート.getLastColumn();

  const 結果マップ = {};
  if (最終行 < 2) return 結果マップ;

  const ヘッダー = シート.getRange(1, 1, 1, 最終列).getValues()[0];
  const 値一覧 = シート.getRange(2, 1, 最終行 - 1, 最終列).getValues();

  const 銘柄コード列 = 列番号を取得する_(ヘッダー, ['銘柄コード', 'code']);
  const 候補区分列 = 列番号を取得する_(ヘッダー, ['候補区分', 'candidateClass']);
  const 候補スコア列 = 列番号を取得する_(ヘッダー, ['候補スコア', 'candidateScore']);
  const メモ列 = 列番号を取得する_(ヘッダー, ['メモ', 'memo']);

  if (
    銘柄コード列 === -1 ||
    候補区分列 === -1 ||
    候補スコア列 === -1 ||
    メモ列 === -1
  ) {
    throw new Error(
      シート.getName() + ' の必要列（銘柄コード / 候補区分 / 候補スコア / メモ）が見つかりません'
    );
  }

  for (let i = 0; i < 値一覧.length; i++) {
    const 行 = 値一覧[i];

    const 銘柄コード = 最終監視用に銘柄コードを正規化する_(行[銘柄コード列]);
    if (!銘柄コード) continue;

    結果マップ[銘柄コード] = {
      銘柄コード: 銘柄コード,
      候補区分: String(行[候補区分列] || '').trim(),
      候補スコア: 数値に変換する_またはnull_(行[候補スコア列]),
      メモ: String(行[メモ列] || '').trim()
    };
  }

  return 結果マップ;
}

/**
 * ランキングDBをコード別最新情報マップに変換
 * 必要列:
 * - 銘柄コード
 * - 銘柄名
 * - 市場
 * - 前日比率
 */
function ランキングDBをコード別最新情報マップに変換する_(シート) {
  const 最終行 = シート.getLastRow();
  const 最終列 = シート.getLastColumn();

  const 結果マップ = {};
  if (最終行 < 2) return 結果マップ;

  const ヘッダー = シート.getRange(1, 1, 1, 最終列).getValues()[0];
  const 値一覧 = シート.getRange(2, 1, 最終行 - 1, 最終列).getValues();

  const 銘柄コード列 = 列番号を取得する_(ヘッダー, ['銘柄コード', 'code']);
  const 銘柄名列 = 列番号を取得する_(ヘッダー, ['銘柄名', 'name']);
  const 市場列 = 列番号を取得する_(ヘッダー, ['市場', '市場区分', 'market']);
  const 前日比列 = 列番号を取得する_(ヘッダー, ['前日比率', '前日比', 'changePct']);

  if (
    銘柄コード列 === -1 ||
    銘柄名列 === -1 ||
    市場列 === -1 ||
    前日比列 === -1
  ) {
    throw new Error(
      シート.getName() + ' の必要列（銘柄コード / 銘柄名 / 市場 / 前日比率）が見つかりません'
    );
  }

  for (let i = 0; i < 値一覧.length; i++) {
    const 行 = 値一覧[i];
    const 銘柄コード = 最終監視用に銘柄コードを正規化する_(行[銘柄コード列]);
    if (!銘柄コード) continue;

    // 下にある行ほど新しい前提
    結果マップ[銘柄コード] = {
      銘柄コード: 銘柄コード,
      銘柄名: String(行[銘柄名列] || '').trim(),
      市場: String(行[市場列] || '').trim(),
      前日比: 数値に変換する_またはnull_(行[前日比列])
    };
  }

  return 結果マップ;
}

/**
 * テーママスタをコード別マップに変換
 * 必要列:
 * - コード or 銘柄コード
 * - テーマ
 */
function テーママスタをコード別マップに変換する_(シート) {
  const 最終行 = シート.getLastRow();
  const 最終列 = シート.getLastColumn();

  const 結果マップ = {};
  if (最終行 < 2) return 結果マップ;

  const ヘッダー = シート.getRange(1, 1, 1, 最終列).getValues()[0];
  const 値一覧 = シート.getRange(2, 1, 最終行 - 1, 最終列).getValues();

  const 銘柄コード列 = 列番号を取得する_(ヘッダー, ['コード', '銘柄コード', 'code']);
  const テーマ列 = 列番号を取得する_(ヘッダー, ['テーマ', 'theme']);

  if (銘柄コード列 === -1 || テーマ列 === -1) {
    throw new Error(
      シート.getName() + ' の必要列（コード / テーマ）が見つかりません'
    );
  }

  for (let i = 0; i < 値一覧.length; i++) {
    const 行 = 値一覧[i];
    const 銘柄コード = 最終監視用に銘柄コードを正規化する_(行[銘柄コード列]);
    if (!銘柄コード) continue;

    結果マップ[銘柄コード] = {
      銘柄コード: 銘柄コード,
      テーマ: String(行[テーマ列] || '').trim()
    };
  }

  return 結果マップ;
}

function 最終出所を判定する_(累積候補, 当日候補) {
  if (累積候補 && 当日候補) return '両方';
  if (累積候補) return '累積';
  if (当日候補) return '当日';
  return '';
}

function 資金流入変化を判定する_(累積候補, 当日候補) {
  if (累積候補 && 当日候補) return '継続強い';
  if (!累積候補 && 当日候補) return '新規流入';
  if (累積候補 && !当日候補) return '鈍化';
  return '';
}

function 最終候補区分を判定する_(累積候補, 当日候補) {
  const 累積区分 = 累積候補 ? 累積候補.候補区分 : '';
  const 当地区分 = 当日候補 ? 当日候補.候補区分 : '';

  const 累積重み = 候補区分の重みを返す_(累積区分);
  const 当日重み = 候補区分の重みを返す_(当地区分);

  return 累積重み >= 当日重み ? 累積区分 : 当地区分;
}

function 最終スコアを計算する_(累積候補, 当日候補, 前日比) {
  const 累積スコア = 累積候補 ? 数値に変換する_またはゼロ_(累積候補.候補スコア) : 0;
  const 当日スコア = 当日候補 ? 数値に変換する_またはゼロ_(当日候補.候補スコア) : 0;

  let ベーススコア = 0;

  // 累積候補を土台
  // 当日候補もある場合だけ上乗せ
  if (累積候補 && 当日候補) {
    ベーススコア = 累積スコア + (当日スコア * 0.3) + 10;
  } else if (累積候補) {
    ベーススコア = 累積スコア;
  } else if (当日候補) {
    ベーススコア = 当日スコア;
  }

  const 前日比補正点 = 前日比補正点を返す_(前日比);

  return ベーススコア + 前日比補正点;
}

function 前日比補正点を返す_(前日比) {
  const 値 = 数値に変換する_またはnull_(前日比);
  if (値 === null) return 0;

  if (値 >= 5) return 20;
  if (値 >= 3) return 12;
  if (値 >= 1) return 6;
  if (値 > -1) return 0;
  if (値 <= -8) return -30;
  if (値 <= -5) return -20;
  if (値 <= -3) return -12;

  return 0;
}

function 除外判定を返す_(前日比) {
  const 値 = 数値に変換する_またはnull_(前日比);
  if (値 === null) return '';
  if (値 <= -7) return '除外';
  return '';
}

function 候補区分の重みを返す_(候補区分) {
  if (候補区分 === 'A') return 3;
  if (候補区分 === 'B') return 2;
  if (候補区分 === 'C') return 1;
  return 0;
}

function 注目記号を返す_(累積候補, 当日候補, 前日比, 除外) {
  if (除外 === '除外') return '×';

  const 前日比数値 = 数値に変換する_またはnull_(前日比);

  if (累積候補 && 当日候補 && 前日比数値 !== null && 前日比数値 >= 0) {
    return '◎';
  }

  if (当日候補) {
    return '○';
  }

  if (累積候補) {
    return '△';
  }

  return '';
}

/**
 * 最終監視シートの条件付き書式を設定
 * 毎回ルールを入れ直す
 */
function 最終監視シートの条件付き書式を設定する_() {
  const シート名 = '最終監視シート';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const シート = ss.getSheetByName(シート名);
  if (!シート) throw new Error('シート「' + シート名 + '」が見つかりません');

  const 最終行 = Math.max(シート.getMaxRows(), 2);
  const ルール一覧 = [];

  // ===== 注目（B列）=====
  ルール一覧.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('◎')
      .setBackground('#b6d7a8')
      .setBold(true)
      .setRanges([シート.getRange('B2:B' + 最終行)])
      .build()
  );

  ルール一覧.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('○')
      .setBackground('#d9ead3')
      .setRanges([シート.getRange('B2:B' + 最終行)])
      .build()
  );

  ルール一覧.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('△')
      .setBackground('#fff2cc')
      .setRanges([シート.getRange('B2:B' + 最終行)])
      .build()
  );

  ルール一覧.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('×')
      .setBackground('#f4cccc')
      .setBold(true)
      .setRanges([シート.getRange('B2:B' + 最終行)])
      .build()
  );

  // ===== 候補区分（H列）=====
  ルール一覧.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('A')
      .setBackground('#cfe2f3')
      .setRanges([シート.getRange('H2:H' + 最終行)])
      .build()
  );

  ルール一覧.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('B')
      .setBackground('#d9ead3')
      .setRanges([シート.getRange('H2:H' + 最終行)])
      .build()
  );

  ルール一覧.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('C')
      .setBackground('#fff2cc')
      .setRanges([シート.getRange('H2:H' + 最終行)])
      .build()
  );

  // ===== 前日比（I列）=====
  ルール一覧.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(5)
      .setBackground('#b6d7a8')
      .setRanges([シート.getRange('I2:I' + 最終行)])
      .build()
  );

  ルール一覧.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(0, 4.9999)
      .setBackground('#d9ead3')
      .setRanges([シート.getRange('I2:I' + 最終行)])
      .build()
  );

  ルール一覧.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThanOrEqualTo(-3)
      .setBackground('#f4cccc')
      .setRanges([シート.getRange('I2:I' + 最終行)])
      .build()
  );

  ルール一覧.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThanOrEqualTo(-7)
      .setBackground('#ea9999')
      .setRanges([シート.getRange('I2:I' + 最終行)])
      .build()
  );

  // ===== 除外（J列）=====
  ルール一覧.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('除外')
      .setBackground('#ea9999')
      .setBold(true)
      .setRanges([シート.getRange('J2:J' + 最終行)])
      .build()
  );

  // ===== 最終スコア（K列）=====
  ルール一覧.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(80)
      .setBackground('#b6d7a8')
      .setRanges([シート.getRange('K2:K' + 最終行)])
      .build()
  );

  ルール一覧.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(60, 79.9999)
      .setBackground('#d9ead3')
      .setRanges([シート.getRange('K2:K' + 最終行)])
      .build()
  );

  // ===== 資金流入変化（L列）=====
  ルール一覧.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('継続強い')
      .setBackground('#b6d7a8')
      .setRanges([シート.getRange('L2:L' + 最終行)])
      .build()
  );

  ルール一覧.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('新規流入')
      .setBackground('#d9eaf7')
      .setRanges([シート.getRange('L2:L' + 最終行)])
      .build()
  );

  ルール一覧.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('鈍化')
      .setBackground('#fce5cd')
      .setRanges([シート.getRange('L2:L' + 最終行)])
      .build()
  );

  シート.setConditionalFormatRules(ルール一覧);
}

function 最終監視用に銘柄コードを正規化する_(値) {
  if (値 === null || 値 === undefined || 値 === '') return '';

  let 文字列 = String(値).trim();
  文字列 = 文字列.replace(/,/g, '');

  if (/^\d+(\.0+)?$/.test(文字列)) {
    文字列 = String(parseInt(文字列, 10));
  }

  return 文字列;
}

function 数値に変換する_またはゼロ_(値) {
  const 数値 = 数値に変換する_またはnull_(値);
  return 数値 === null ? 0 : 数値;
}

function 数値に変換する_またはnull_(値) {
  if (値 === null || 値 === undefined || 値 === '') return null;

  if (typeof 値 === 'number') {
    return isNaN(値) ? null : 値;
  }

  let 文字列 = String(値).trim();
  if (文字列 === '') return null;

  文字列 = 文字列.replace(/,/g, '');

  const 数値 = parseFloat(文字列);
  return isNaN(数値) ? null : 数値;
}

function 列番号を取得する_(ヘッダー, 候補名一覧) {
  for (let i = 0; i < 候補名一覧.length; i++) {
    const 候補名 = 候補名一覧[i];
    const 列番号 = ヘッダー.indexOf(候補名);
    if (列番号 !== -1) return 列番号;
  }
  return -1;
}

/**
 * まとめ更新用
 */
function 監視シートまでまとめて更新する() {
  ランキング処理をまとめて実行する();
  最終監視シートを作成する();
  Logger.log('取得・集計・候補抽出・最終監視シート 更新完了');
}

function 最終監視シートの銘柄コードにリンクを設定する_(シート, データ開始行, データ件数) {
  if (!シート || データ件数 <= 0) return;

  const 銘柄コード列 = 3; // C列
  const range = シート.getRange(データ開始行, 銘柄コード列, データ件数, 1);
  const values = range.getValues();

  const richTextValues = values.map(function(row) {
    const code = String(row[0] || '').trim();
    if (!code) {
      return [SpreadsheetApp.newRichTextValue().setText('').build()];
    }

    const url = 'https://finance.yahoo.co.jp/quote/' + code + '.T';

    return [
      SpreadsheetApp.newRichTextValue()
        .setText(code)
        .setLinkUrl(url)
        .build()
    ];
  });

  range.setRichTextValues(richTextValues);
}