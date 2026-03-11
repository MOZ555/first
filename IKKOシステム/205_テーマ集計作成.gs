/*************************************************
 * 205_テーマ集計作成.gs
 *************************************************/

/**
 * 実行入口
 * 1. 最終監視シートからテーマ集計を作成
 * 2. テーマ集計履歴へ追記
 * 3. テーマ推移グラフ用シートを作成
 * 4. 線グラフを作成 / 更新
 */
function テーマ集計を更新する() {
  const 集計結果 = テーマ集計を作成する_();
  テーマ集計履歴に追記する_(集計結果);
  テーマ推移グラフ用シートを作成する_();
  テーマ推移グラフを作成または更新する_();

  Logger.log('テーマ集計 更新完了');
}

/**
 * 最終監視シートからテーマ集計を作成し、
 * 「テーマ集計」シートへ出力する
 *
 * 返り値:
 * [
 *   {
 *     日付: '2026/03/10',
 *     テーマ: '半導体',
 *     銘柄数: 4,
 *     平均最終スコア: 78.25,
 *     当日候補数: 3,
 *     累積候補数: 2,
 *     strongest銘柄: '6723 ルネサス',
 *     主力銘柄一覧: '6723 ルネサス / 7735 SCREEN / 7741 HOYA'
 *   },
 *   ...
 * ]
 */
function テーマ集計を作成する_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const 元シート名 = '最終監視シート';
  const 出力シート名 = 'テーマ集計';

  const 元シート = ss.getSheetByName(元シート名);
  if (!元シート) {
    throw new Error('シート「' + 元シート名 + '」が見つかりません');
  }

  let 出力シート = ss.getSheetByName(出力シート名);
  if (!出力シート) {
    出力シート = ss.insertSheet(出力シート名);
  }

  const 最終行 = 元シート.getLastRow();
  const 最終列 = 元シート.getLastColumn();

  // 出力側クリア（条件付き書式は使わない想定なので全消しでもOK）
  出力シート.clearContents();

  const 出力ヘッダー = [[
    'テーマ順位',
    '日付',
    'テーマ',
    '銘柄数',
    '平均最終スコア',
    '当日候補数',
    '累積候補数',
    'strongest銘柄',
    '主力銘柄一覧'
  ]];
  出力シート.getRange(1, 1, 1, 出力ヘッダー[0].length).setValues(出力ヘッダー);

  if (最終行 < 2) {
    Logger.log('最終監視シートにデータがありません');
    return [];
  }

  const ヘッダー = 元シート.getRange(1, 1, 1, 最終列).getValues()[0];
  const 値一覧 = 元シート.getRange(2, 1, 最終行 - 1, 最終列).getValues();

  const 銘柄コード列 = 列番号を取得する_(ヘッダー, ['銘柄コード', 'code']);
  const 銘柄名列 = 列番号を取得する_(ヘッダー, ['銘柄名', 'name']);
  const テーマ列 = 列番号を取得する_(ヘッダー, ['テーマ', 'theme']);
  const 出所列 = 列番号を取得する_(ヘッダー, ['出所', 'source']);
  const 除外列 = 列番号を取得する_(ヘッダー, ['除外', 'exclude']);
  const 最終スコア列 = 列番号を取得する_(ヘッダー, ['最終スコア', 'finalScore']);

  if (
    銘柄コード列 === -1 ||
    銘柄名列 === -1 ||
    テーマ列 === -1 ||
    出所列 === -1 ||
    除外列 === -1 ||
    最終スコア列 === -1
  ) {
    throw new Error('最終監視シート の必要列（銘柄コード / 銘柄名 / テーマ / 出所 / 除外 / 最終スコア）が見つかりません');
  }

  const テーマ集計マップ = {};
  const 今日 = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd');

  for (let i = 0; i < 値一覧.length; i++) {
    const 行 = 値一覧[i];

    const 銘柄コード = 文字列にする_(行[銘柄コード列]);
    const 銘柄名 = 文字列にする_(行[銘柄名列]);
    const テーマ = 文字列にする_(行[テーマ列]) || 'テーマ未設定';
    const 出所 = 文字列にする_(行[出所列]);
    const 除外 = 文字列にする_(行[除外列]);
    const 最終スコア = 数値に変換する_またはゼロ_(行[最終スコア列]);

    if (!銘柄コード) continue;
    if (除外 === '除外') continue; // 除外銘柄はテーマ集計から外す

    if (!テーマ集計マップ[テーマ]) {
      テーマ集計マップ[テーマ] = {
        日付: 今日,
        テーマ: テーマ,
        銘柄数: 0,
        最終スコア合計: 0,
        当日候補数: 0,
        累積候補数: 0,
        strongest銘柄: '',
        strongestスコア: -999999,
        銘柄一覧: []
      };
    }

    const 集計 = テーマ集計マップ[テーマ];

    集計.銘柄数 += 1;
    集計.最終スコア合計 += 最終スコア;
    集計.銘柄一覧.push({
      銘柄コード: 銘柄コード,
      銘柄名: 銘柄名,
      最終スコア: 最終スコア
    });

    if (出所 === '当日') {
      集計.当日候補数 += 1;
    } else if (出所 === '累積') {
      集計.累積候補数 += 1;
    } else if (出所 === '両方') {
      集計.当日候補数 += 1;
      集計.累積候補数 += 1;
    }

    if (最終スコア > 集計.strongestスコア) {
      集計.strongestスコア = 最終スコア;
      集計.strongest銘柄 = 銘柄コード + ' ' + 銘柄名;
    }
  }

  const 集計結果一覧 = Object.keys(テーマ集計マップ).map(function(テーマ) {
    const 集計 = テーマ集計マップ[テーマ];

    集計.銘柄一覧.sort(function(a, b) {
      return b.最終スコア - a.最終スコア;
    });

    const 主力銘柄一覧 = 集計.銘柄一覧
      .slice(0, 3)
      .map(function(x) {
        return x.銘柄コード + ' ' + x.銘柄名;
      })
      .join(' / ');

    return {
      日付: 集計.日付,
      テーマ: 集計.テーマ,
      銘柄数: 集計.銘柄数,
      平均最終スコア: 集計.銘柄数 > 0 ? 集計.最終スコア合計 / 集計.銘柄数 : 0,
      当日候補数: 集計.当日候補数,
      累積候補数: 集計.累積候補数,
      strongest銘柄: 集計.strongest銘柄,
      主力銘柄一覧: 主力銘柄一覧
    };
  });

  集計結果一覧.sort(function(a, b) {
    if (b.銘柄数 !== a.銘柄数) return b.銘柄数 - a.銘柄数;
    if (b.平均最終スコア !== a.平均最終スコア) return b.平均最終スコア - a.平均最終スコア;
    return String(a.テーマ).localeCompare(String(b.テーマ));
  });

  const 出力データ = [];
  for (let i = 0; i < 集計結果一覧.length; i++) {
    const r = 集計結果一覧[i];
    出力データ.push([
      i + 1,
      r.日付,
      r.テーマ,
      r.銘柄数,
      r.平均最終スコア,
      r.当日候補数,
      r.累積候補数,
      r.strongest銘柄,
      r.主力銘柄一覧
    ]);
  }

  if (出力データ.length > 0) {
    出力シート.getRange(2, 1, 出力データ.length, 出力データ[0].length).setValues(出力データ);
    出力シート.getRange(2, 1, 出力データ.length, 1).setNumberFormat('0');
    出力シート.getRange(2, 4, 出力データ.length, 1).setNumberFormat('0');
    出力シート.getRange(2, 5, 出力データ.length, 1).setNumberFormat('0.00');
    出力シート.getRange(2, 6, 出力データ.length, 2).setNumberFormat('0');
  }

  出力シート.setFrozenRows(1);
  出力シート.autoResizeColumns(1, 9);

  return 集計結果一覧;
}

/**
 * テーマ集計結果を「テーマ集計履歴」へ追記
 * 同一日付×同一テーマが既にあれば上書き
 */
function テーマ集計履歴に追記する_(集計結果一覧) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const シート名 = 'テーマ集計履歴';

  let シート = ss.getSheetByName(シート名);
  if (!シート) {
    シート = ss.insertSheet(シート名);
  }

  const ヘッダー = [[
    '日付',
    'テーマ',
    '銘柄数',
    '平均最終スコア',
    '当日候補数',
    '累積候補数',
    'strongest銘柄',
    '主力銘柄一覧'
  ]];

  if (シート.getLastRow() === 0) {
    シート.getRange(1, 1, 1, ヘッダー[0].length).setValues(ヘッダー);
  }

  if (!集計結果一覧 || 集計結果一覧.length === 0) {
    return;
  }

  const 最終行 = シート.getLastRow();

  const 既存キー行マップ = {};
  if (最終行 >= 2) {
    const 既存データ = シート.getRange(2, 1, 最終行 - 1, 2).getValues();
    for (let i = 0; i < 既存データ.length; i++) {
      const 日付 = 文字列にする_(既存データ[i][0]);
      const テーマ = 文字列にする_(既存データ[i][1]);
      const key = 日付 + '|' + テーマ;
      既存キー行マップ[key] = i + 2;
    }
  }

  const appendRows = [];
  const updateTargets = [];

  for (let i = 0; i < 集計結果一覧.length; i++) {
    const r = 集計結果一覧[i];
    const row = [
      r.日付,
      r.テーマ,
      r.銘柄数,
      r.平均最終スコア,
      r.当日候補数,
      r.累積候補数,
      r.strongest銘柄,
      r.主力銘柄一覧
    ];

    const key = r.日付 + '|' + r.テーマ;

    if (既存キー行マップ[key]) {
      updateTargets.push({
        rowNumber: 既存キー行マップ[key],
        values: row
      });
    } else {
      appendRows.push(row);
    }
  }

  for (let i = 0; i < updateTargets.length; i++) {
    const t = updateTargets[i];
    シート.getRange(t.rowNumber, 1, 1, t.values.length).setValues([t.values]);
  }

  if (appendRows.length > 0) {
    シート.getRange(シート.getLastRow() + 1, 1, appendRows.length, appendRows[0].length).setValues(appendRows);
  }

  const 更新後最終行 = シート.getLastRow();
  if (更新後最終行 >= 2) {
    シート.getRange(2, 3, 更新後最終行 - 1, 1).setNumberFormat('0');
    シート.getRange(2, 4, 更新後最終行 - 1, 1).setNumberFormat('0.00');
    シート.getRange(2, 5, 更新後最終行 - 1, 2).setNumberFormat('0');
  }

  シート.setFrozenRows(1);
  シート.autoResizeColumns(1, 8);
}

/**
 * 「テーマ集計履歴」から
 * 「テーマ推移グラフ用」シートを作る
 *
 * 今回はまず
 * 「テーマ別 銘柄数推移」用の横持ちテーブルを作る
 */
function テーマ推移グラフ用シートを作成する_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const 履歴シート名 = 'テーマ集計履歴';
  const 出力シート名 = 'テーマ推移グラフ用';

  const 履歴シート = ss.getSheetByName(履歴シート名);
  if (!履歴シート) {
    throw new Error('シート「' + 履歴シート名 + '」が見つかりません');
  }

  let 出力シート = ss.getSheetByName(出力シート名);
  if (!出力シート) {
    出力シート = ss.insertSheet(出力シート名);
  }

  出力シート.clearContents();

  const 最終行 = 履歴シート.getLastRow();
  const 最終列 = 履歴シート.getLastColumn();

  if (最終行 < 2) {
    出力シート.getRange(1, 1).setValue('履歴データなし');
    return;
  }

  const ヘッダー = 履歴シート.getRange(1, 1, 1, 最終列).getValues()[0];
  const 値一覧 = 履歴シート.getRange(2, 1, 最終行 - 1, 最終列).getValues();

  const 日付列 = 列番号を取得する_(ヘッダー, ['日付', 'date']);
  const テーマ列 = 列番号を取得する_(ヘッダー, ['テーマ', 'theme']);
  const 銘柄数列 = 列番号を取得する_(ヘッダー, ['銘柄数', 'count']);

  if (日付列 === -1 || テーマ列 === -1 || 銘柄数列 === -1) {
    throw new Error('テーマ集計履歴 の必要列（日付 / テーマ / 銘柄数）が見つかりません');
  }

  const 日付一覧オブジェクト = {};
  const テーマ一覧オブジェクト = {};
  const データマップ = {};

  for (let i = 0; i < 値一覧.length; i++) {
    const 行 = 値一覧[i];
    const 日付 = 文字列にする_(行[日付列]);
    const テーマ = 文字列にする_(行[テーマ列]);
    const 銘柄数 = 数値に変換する_またはゼロ_(行[銘柄数列]);

    if (!日付 || !テーマ) continue;

    日付一覧オブジェクト[日付] = true;
    テーマ一覧オブジェクト[テーマ] = true;

    if (!データマップ[日付]) {
      データマップ[日付] = {};
    }
    データマップ[日付][テーマ] = 銘柄数;
  }

  const 日付一覧 = Object.keys(日付一覧オブジェクト).sort();
  let テーマ一覧 = Object.keys(テーマ一覧オブジェクト).sort();

  // 表示線が増えすぎると見にくいので、直近日の上位テーマを優先
  if (日付一覧.length > 0) {
    const 最新日付 = 日付一覧[日付一覧.length - 1];
    const 最新テーマスコア配列 = テーマ一覧.map(function(テーマ) {
      return {
        テーマ: テーマ,
        銘柄数: データマップ[最新日付] && データマップ[最新日付][テーマ]
          ? データマップ[最新日付][テーマ]
          : 0
      };
    });

    最新テーマスコア配列.sort(function(a, b) {
      if (b.銘柄数 !== a.銘柄数) return b.銘柄数 - a.銘柄数;
      return String(a.テーマ).localeCompare(String(b.テーマ));
    });

    テーマ一覧 = 最新テーマスコア配列.slice(0, 8).map(function(x) {
      return x.テーマ;
    });
  }

  const 出力データ = [];
  const 出力ヘッダー = ['日付'].concat(テーマ一覧);
  出力データ.push(出力ヘッダー);

  for (let i = 0; i < 日付一覧.length; i++) {
    const 日付 = 日付一覧[i];
    const row = [日付];

    for (let j = 0; j < テーマ一覧.length; j++) {
      const テーマ = テーマ一覧[j];
      const 値 = データマップ[日付] && データマップ[日付][テーマ]
        ? データマップ[日付][テーマ]
        : 0;
      row.push(値);
    }

    出力データ.push(row);
  }

  出力シート.getRange(1, 1, 出力データ.length, 出力データ[0].length).setValues(出力データ);

  if (出力データ.length >= 2 && 出力データ[0].length >= 2) {
    出力シート.getRange(2, 2, 出力データ.length - 1, 出力データ[0].length - 1).setNumberFormat('0');
  }

  出力シート.setFrozenRows(1);
  出力シート.autoResizeColumns(1, 出力データ[0].length);
}

/**
 * 「テーマ推移グラフ用」シートを元に
 * 線グラフを作成 / 更新
 */
function テーマ推移グラフを作成または更新する_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const シート名 = 'テーマ推移グラフ用';
  const シート = ss.getSheetByName(シート名);
  if (!シート) {
    throw new Error('シート「' + シート名 + '」が見つかりません');
  }

  const 最終行 = シート.getLastRow();
  const 最終列 = シート.getLastColumn();

  if (最終行 < 2 || 最終列 < 2) {
    Logger.log('グラフ用データ不足のためグラフ作成スキップ');
    return;
  }

  // 既存グラフ削除
  const charts = シート.getCharts();
  for (let i = 0; i < charts.length; i++) {
    シート.removeChart(charts[i]);
  }

  const range = シート.getRange(1, 1, 最終行, 最終列);

  const chart = シート.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(range)
    .setPosition(2, Math.max(最終列 + 2, 12), 0, 0)
    .setOption('title', 'テーマ別 銘柄数推移')
    .setOption('legend', { position: 'right' })
    .setOption('curveType', 'function')
    .setOption('pointSize', 5)
    .setOption('lineWidth', 2)
    .setOption('hAxis', { title: '日付' })
    .setOption('vAxis', { title: '銘柄数', minValue: 0 })
    .setNumHeaders(1)
    .build();

  シート.insertChart(chart);
}

/**
 * まとめ実行
 * 監視シート更新後にテーマ集計まで一気にやる用
 */
function 監視とテーマ集計までまとめて更新する() {
  監視シートまでまとめて更新する();
  テーマ集計を更新する();
  Logger.log('監視シート + テーマ集計 + グラフ 更新完了');
}

/*************************************************
 * 共通関数
 *************************************************/

function 数値に変換する_またはゼロ_(値) {
  const num = 数値に変換する_またはnull_(値);
  return num === null ? 0 : num;
}

function 数値に変換する_またはnull_(値) {
  if (値 === null || 値 === undefined || 値 === '') return null;

  if (typeof 値 === 'number') {
    return isNaN(値) ? null : 値;
  }

  let s = String(値).trim();
  if (s === '') return null;

  s = s.replace(/,/g, '');

  const num = parseFloat(s);
  return isNaN(num) ? null : num;
}

function 列番号を取得する_(ヘッダー, 候補名一覧) {
  for (let i = 0; i < 候補名一覧.length; i++) {
    const 候補名 = 候補名一覧[i];
    const idx = ヘッダー.indexOf(候補名);
    if (idx !== -1) return idx;
  }
  return -1;
}

function 文字列にする_(値) {
  if (値 === null || 値 === undefined) return '';
  return String(値).trim();
}