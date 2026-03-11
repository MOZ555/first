function ランキング滞在変化を作成する() {
  const DBシート名 = 'ランキングDB';
  const テーママスタシート名 = 'テーママスタ';
  const 出力シート名 = 'ランキング滞在変化';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const DBシート = ss.getSheetByName(DBシート名);
  if (!DBシート) throw new Error(DBシート名 + ' シートがありません');

  const テーママスタシート = ss.getSheetByName(テーママスタシート名); // 無くてもOK

  let 出力シート = ss.getSheetByName(出力シート名);
  if (!出力シート) {
    出力シート = ss.insertSheet(出力シート名);
  }

  const 出力列数 = 16;

  // 2行目以降だけクリア（条件付き書式を消さない）
  const 出力最終行 = 出力シート.getLastRow();
  const 出力最終列 = Math.max(出力シート.getLastColumn(), 出力列数);
  if (出力最終行 >= 2) {
    出力シート.getRange(2, 1, 出力最終行 - 1, 出力最終列).clearContent();
  }

  出力シート.getRange(1, 1, 1, 出力列数).setValues([[
    '変化順位',
    '銘柄コード',
    '銘柄名',
    '市場',
    'テーマ',
    '前日比',
    '前回日付',
    '当日日付',
    '前回出現回数',
    '当日出現回数',
    '出現回数差',
    '前回平均順位',
    '当日平均順位',
    '平均順位差',
    '判定',
    'メモ'
  ]]);

  const 最終行 = DBシート.getLastRow();
  const 最終列 = DBシート.getLastColumn();

  if (最終行 < 2) {
    Logger.log('ランキングDB にデータがありません');
    出力シート.getRange(2, 1).setValue('ランキングDB にデータがありません');
    return;
  }

  const ヘッダー = DBシート.getRange(1, 1, 1, 最終列).getValues()[0];
  const 値一覧 = DBシート.getRange(2, 1, 最終行 - 1, 最終列).getValues();

  const 時刻列 = 列番号を取得する_(ヘッダー, ['時刻', 'time', '取得時刻']);
  const バッチID列 = 列番号を取得する_(ヘッダー, ['バッチID', 'batchId']);
  const 順位列 = 列番号を取得する_(ヘッダー, ['順位', 'rank']);
  const 銘柄コード列 = 列番号を取得する_(ヘッダー, ['銘柄コード', 'code']);
  const 銘柄名列 = 列番号を取得する_(ヘッダー, ['銘柄名', 'name']);
  const 市場列 = 列番号を取得する_(ヘッダー, ['市場', '市場区分', 'market']);
  const 前日比列 = 列番号を取得する_(ヘッダー, ['前日比率', '前日比', 'changePct']);

  if (時刻列 === -1 || バッチID列 === -1 || 順位列 === -1 || 銘柄コード列 === -1) {
    throw new Error('ランキングDB の必要列（時刻 / バッチID / 順位 / 銘柄コード）が見つかりません');
  }

  const 日別集計 = {};
  const 日付一覧オブジェクト = {};
  const 最新情報マップ = {};

  for (let i = 0; i < 値一覧.length; i++) {
    const 行 = 値一覧[i];

    const 時刻 = 行[時刻列];
    const バッチID = String(行[バッチID列] || '').trim();
    const 順位 = 数値に変換する_またはnull_(行[順位列]);
    const 銘柄コード = 候補抽出用に銘柄コードを正規化する_(行[銘柄コード列]);

    if (!銘柄コード || !バッチID || 順位 === null) continue;

    const 日付キー = 値から日付キーを作る_(時刻);
    if (!日付キー) continue;

    日付一覧オブジェクト[日付キー] = true;

    if (!日別集計[日付キー]) {
      日別集計[日付キー] = {};
    }

    if (!日別集計[日付キー][銘柄コード]) {
      日別集計[日付キー][銘柄コード] = {
        銘柄コード: 銘柄コード,
        バッチID一覧: {},
        順位合計: 0,
        出現回数: 0
      };
    }

    const 集計 = 日別集計[日付キー][銘柄コード];

    // 同じバッチIDでの重複カウント防止
    if (!集計.バッチID一覧[バッチID]) {
      集計.バッチID一覧[バッチID] = true;
      集計.出現回数 += 1;
      集計.順位合計 += 順位;
    }

    // 下の行ほど新しい前提で上書き
    最新情報マップ[銘柄コード] = {
      銘柄名: 銘柄名列 !== -1 ? String(行[銘柄名列] || '').trim() : '',
      市場: 市場列 !== -1 ? String(行[市場列] || '').trim() : '',
      前日比: 前日比列 !== -1 ? 数値に変換する_またはnull_(行[前日比列]) : null
    };
  }

  const テーママップ = テーママスタシート
    ? テーママスタをコード別マップに変換する_(テーママスタシート)
    : {};

  const 日付一覧 = Object.keys(日付一覧オブジェクト).sort();
  if (日付一覧.length < 2) {
    Logger.log('比較対象の日付が2日分ありません');
    出力シート.getRange(2, 1).setValue('比較対象の日付が2日分ありません');
    return;
  }

  const 当日日付 = 日付一覧[日付一覧.length - 1];
  const 前回日付 = 日付一覧[日付一覧.length - 2];

  const 当日データ = 日別集計[当日日付] || {};
  const 前回データ = 日別集計[前回日付] || {};

  const 全銘柄コードオブジェクト = {};
  Object.keys(前回データ).forEach(function(銘柄コード) {
    全銘柄コードオブジェクト[銘柄コード] = true;
  });
  Object.keys(当日データ).forEach(function(銘柄コード) {
    全銘柄コードオブジェクト[銘柄コード] = true;
  });

  const 全銘柄コード一覧 = Object.keys(全銘柄コードオブジェクト);
  const 出力データ = [];

  for (let i = 0; i < 全銘柄コード一覧.length; i++) {
    const 銘柄コード = 全銘柄コード一覧[i];
    const 前回 = 前回データ[銘柄コード] || null;
    const 当日 = 当日データ[銘柄コード] || null;
    const 最新情報 = 最新情報マップ[銘柄コード] || {};
    const テーマ情報 = テーママップ[銘柄コード] || {};

    const 前回出現回数 = 前回 ? 前回.出現回数 : 0;
    const 当日出現回数 = 当日 ? 当日.出現回数 : 0;
    const 出現回数差 = 当日出現回数 - 前回出現回数;

    const 前回平均順位 = 前回 && 前回出現回数 > 0 ? 前回.順位合計 / 前回.出現回数 : null;
    const 当日平均順位 = 当日 && 当日出現回数 > 0 ? 当日.順位合計 / 当日.出現回数 : null;

    const 平均順位差 =
      前回平均順位 !== null && 当日平均順位 !== null
        ? 前回平均順位 - 当日平均順位
        : null;

    const 判定 = 滞在変化の判定をする_(前回出現回数, 当日出現回数, 前回平均順位, 当日平均順位);
    const メモ = 滞在変化メモを作成する_(前回出現回数, 当日出現回数, 前回平均順位, 当日平均順位);

    const 銘柄名 = 最新情報.銘柄名 || '';
    const 市場 = 最新情報.市場 || '';
    const 前日比 = 最新情報.前日比;
    const テーマ = テーマ情報.テーマ || '';

    出力データ.push([
      '', // 変化順位
      銘柄コード,
      銘柄名,
      市場,
      テーマ,
      前日比 === null || 前日比 === undefined ? '' : 前日比,
      日付キーを表示用に変換する_(前回日付),
      日付キーを表示用に変換する_(当日日付),
      前回出現回数,
      当日出現回数,
      出現回数差,
      前回平均順位,
      当日平均順位,
      平均順位差,
      判定,
      メモ
    ]);
  }

  出力データ.sort(function(a, b) {
    const 出現回数差A = 数値に変換する_またはゼロ_(a[10]);
    const 出現回数差B = 数値に変換する_またはゼロ_(b[10]);
    if (出現回数差B !== 出現回数差A) return 出現回数差B - 出現回数差A;

    const 平均順位差A = 数値に変換する_またはゼロ_(a[13]);
    const 平均順位差B = 数値に変換する_またはゼロ_(b[13]);
    if (平均順位差B !== 平均順位差A) return 平均順位差B - 平均順位差A;

    return String(a[1]).localeCompare(String(b[1]));
  });

  for (let i = 0; i < 出力データ.length; i++) {
    出力データ[i][0] = i + 1;
  }

  if (出力データ.length > 0) {
    出力シート.getRange(2, 1, 出力データ.length, 出力列数).setValues(出力データ);

    出力シート.getRange(2, 1, 出力データ.length, 1).setNumberFormat('0');      // 変化順位
    出力シート.getRange(2, 6, 出力データ.length, 1).setNumberFormat('0.00');     // 前日比
    出力シート.getRange(2, 9, 出力データ.length, 3).setNumberFormat('0');        // 出現回数系
    出力シート.getRange(2, 12, 出力データ.length, 3).setNumberFormat('0.00');    // 平均順位系
  }

  出力シート.setFrozenRows(1);
  出力シート.autoResizeColumns(1, 出力列数)

  ランキング滞在変化の条件付き書式を設定する_();

  Logger.log('ランキング滞在変化 作成完了: ' + 出力データ.length + '件 / 比較: ' + 前回日付 + ' → ' + 当日日付);
}

function 滞在変化の判定をする_(前回出現回数, 当日出現回数, 前回平均順位, 当日平均順位) {
  if (前回出現回数 === 0 && 当日出現回数 > 0) {
    return '新規流入';
  }

  if (当日出現回数 === 0 && 前回出現回数 > 0) {
    return '消失';
  }

  if (当日出現回数 > 前回出現回数) {
    if (前回平均順位 !== null && 当日平均順位 !== null && 当日平均順位 < 前回平均順位) {
      return '流入強化';
    }
    return '出現増加';
  }

  if (当日出現回数 < 前回出現回数) {
    return '鈍化';
  }

  if (前回平均順位 !== null && 当日平均順位 !== null) {
    if (当日平均順位 < 前回平均順位) {
      return '順位改善';
    }
    if (当日平均順位 > 前回平均順位) {
      return '順位悪化';
    }
  }

  return '横ばい';
}

function 滞在変化メモを作成する_(前回出現回数, 当日出現回数, 前回平均順位, 当日平均順位) {
  if (前回出現回数 === 0 && 当日出現回数 > 0) {
    return '前回なし → 当日登場';
  }

  if (当日出現回数 === 0 && 前回出現回数 > 0) {
    return '前回あり → 当日圏外';
  }

  const 出現差 = 当日出現回数 - 前回出現回数;

  if (出現差 > 0 && 前回平均順位 !== null && 当日平均順位 !== null && 当日平均順位 < 前回平均順位) {
    return '出現回数増 + 平均順位改善';
  }

  if (出現差 > 0) {
    return '出現回数が増加';
  }

  if (出現差 < 0) {
    return '出現回数が減少';
  }

  if (前回平均順位 !== null && 当日平均順位 !== null && 当日平均順位 < 前回平均順位) {
    return '出現回数同じで順位改善';
  }

  if (前回平均順位 !== null && 当日平均順位 !== null && 当日平均順位 > 前回平均順位) {
    return '出現回数同じで順位悪化';
  }

  return '大きな変化なし';
}

function 日付キーを表示用に変換する_(日付キー) {
  if (!日付キー) return '';

  const 文字列 = String(日付キー).trim();

  if (/^\d{8}$/.test(文字列)) {
    return 文字列.slice(0, 4) + '/' + 文字列.slice(4, 6) + '/' + 文字列.slice(6, 8);
  }

  return 文字列;
}

function 値から日付キーを作る_(値) {
  if (値 === null || 値 === undefined || 値 === '') return '';

  if (Object.prototype.toString.call(値) === '[object Date]' && !isNaN(値.getTime())) {
    return Utilities.formatDate(値, Session.getScriptTimeZone(), 'yyyyMMdd');
  }

  const 文字列 = String(値).trim();
  if (!文字列) return '';

  const 日付時刻 = new Date(文字列);
  if (!isNaN(日付時刻.getTime())) {
    return Utilities.formatDate(日付時刻, Session.getScriptTimeZone(), 'yyyyMMdd');
  }

  const 数字のみ = 文字列.replace(/[^\d]/g, '');
  if (/^\d{8}$/.test(数字のみ)) {
    return 数字のみ;
  }

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
    throw new Error(シート.getName() + ' の必要列（コード / テーマ）が見つかりません');
  }

  for (let i = 0; i < 値一覧.length; i++) {
    const 行 = 値一覧[i];
    const 銘柄コード = 候補抽出用に銘柄コードを正規化する_(行[銘柄コード列]);
    if (!銘柄コード) continue;

    結果マップ[銘柄コード] = {
      銘柄コード: 銘柄コード,
      テーマ: String(行[テーマ列] || '').trim()
    };
  }

  return 結果マップ;
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

function ランキング滞在変化の条件付き書式を設定する_() {
  const シート名 = 'ランキング滞在変化';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const シート = ss.getSheetByName(シート名);
  if (!シート) throw new Error(シート名 + ' シートがありません');

  const 最終行 = Math.max(シート.getMaxRows(), 2);

  const ルール一覧 = [];

  // ===== 前日比（F列）=====
  // 5%以上
  ルール一覧.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(5)
      .setBackground('#b6d7a8')
      .setRanges([シート.getRange('F2:F' + 最終行)])
      .build()
  );

  // 0%以上
  ルール一覧.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(0)
      .setBackground('#d9ead3')
      .setRanges([シート.getRange('F2:F' + 最終行)])
      .build()
  );

  // -3%以下
  ルール一覧.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThanOrEqualTo(-3)
      .setBackground('#f4cccc')
      .setRanges([シート.getRange('F2:F' + 最終行)])
      .build()
  );

  // -7%以下
  ルール一覧.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThanOrEqualTo(-7)
      .setBackground('#ea9999')
      .setRanges([シート.getRange('F2:F' + 最終行)])
      .build()
  );

  // ===== 判定（O列）=====
  ルール一覧.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('新規流入')
      .setBackground('#d9eaf7')
      .setRanges([シート.getRange('O2:O' + 最終行)])
      .build()
  );

  ルール一覧.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('流入強化')
      .setBackground('#d9ead3')
      .setRanges([シート.getRange('O2:O' + 最終行)])
      .build()
  );

  ルール一覧.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('消失')
      .setBackground('#dddddd')
      .setRanges([シート.getRange('O2:O' + 最終行)])
      .build()
  );

  ルール一覧.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('鈍化')
      .setBackground('#fce5cd')
      .setRanges([シート.getRange('O2:O' + 最終行)])
      .build()
  );

  // ===== 出現回数差（K列）=====
  ルール一覧.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(10)
      .setBackground('#d9ead3')
      .setRanges([シート.getRange('K2:K' + 最終行)])
      .build()
  );

  // ===== 平均順位差（N列）=====
  ルール一覧.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(5)
      .setBackground('#d9eaf7')
      .setRanges([シート.getRange('N2:N' + 最終行)])
      .build()
  );

  シート.setConditionalFormatRules(ルール一覧);
}