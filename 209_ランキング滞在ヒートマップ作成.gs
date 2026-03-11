function ランキング滞在ヒートマップを作成する() {
  const 元シート名 = 'ランキング滞在集計_累積';
  const 出力シート名 = 'ランキング滞在ヒートマップ';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const 元シート = ss.getSheetByName(元シート名);
  if (!元シート) throw new Error(元シート名 + ' シートがありません');

  let 出力シート = ss.getSheetByName(出力シート名);
  if (!出力シート) {
    出力シート = ss.insertSheet(出力シート名);
  }

  const 最終行 = 元シート.getLastRow();
  const 最終列 = 元シート.getLastColumn();

  出力シート.clearContents();
  出力シート.clearFormats();

  出力シート.getRange(1, 1, 1, 8).setValues([[
    '順位',
    '銘柄コード',
    '出現回数',
    '平均順位',
    '最新順位',
    '滞在スコア',
    'ヒート',
    '見た目メモ'
  ]]);

  if (最終行 < 2) {
    Logger.log('ランキング滞在集計_累積 にデータがありません');
    return;
  }

  const ヘッダー = 元シート.getRange(1, 1, 1, 最終列).getValues()[0];
  const 値一覧 = 元シート.getRange(2, 1, 最終行 - 1, 最終列).getValues();

  const 銘柄コード列 = 列番号を取得する_(ヘッダー, ['銘柄コード', 'code']);
  const 出現回数列 = 列番号を取得する_(ヘッダー, ['出現回数', 'count']);
  const 平均順位列 = 列番号を取得する_(ヘッダー, ['平均順位', 'avgRank']);
  const 最新順位列 = 列番号を取得する_(ヘッダー, ['最新順位', 'latestRank']);
  const 滞在スコア列 = 列番号を取得する_(ヘッダー, ['滞在スコア', 'stayScore']);

  if (
    銘柄コード列 === -1 ||
    出現回数列 === -1 ||
    平均順位列 === -1 ||
    最新順位列 === -1 ||
    滞在スコア列 === -1
  ) {
    throw new Error('ランキング滞在集計_累積 の必要列が見つかりません');
  }

  const データ = [];

  for (let i = 0; i < 値一覧.length; i++) {
    const 行 = 値一覧[i];

    const 銘柄コード = String(行[銘柄コード列] || '').trim();
    const 出現回数 = 数値に変換する_またはnull_(行[出現回数列]);
    const 平均順位 = 数値に変換する_またはnull_(行[平均順位列]);
    const 最新順位 = 数値に変換する_またはnull_(行[最新順位列]);
    const 滞在スコア = 数値に変換する_またはnull_(行[滞在スコア列]);

    if (!銘柄コード) continue;
    if (出現回数 === null || 平均順位 === null || 最新順位 === null || 滞在スコア === null) continue;

    データ.push({
      銘柄コード: 銘柄コード,
      出現回数: 出現回数,
      平均順位: 平均順位,
      最新順位: 最新順位,
      滞在スコア: 滞在スコア
    });
  }

  データ.sort(function(a, b) {
    if (b.滞在スコア !== a.滞在スコア) return b.滞在スコア - a.滞在スコア;
    if (b.出現回数 !== a.出現回数) return b.出現回数 - a.出現回数;
    return a.平均順位 - b.平均順位;
  });

  const 最大滞在スコア = データ.length > 0 ? データ[0].滞在スコア : 0;

  const 出力データ = データ.map(function(item, index) {
    const ヒート長さ = ヒート長さを計算する_(item.滞在スコア, 最大滞在スコア);
    const ヒート = '█'.repeat(ヒート長さ);
    const 見た目メモ = ヒートメモを作成する_(item.出現回数, item.平均順位, item.最新順位);

    return [
      index + 1,
      item.銘柄コード,
      item.出現回数,
      item.平均順位,
      item.最新順位,
      item.滞在スコア,
      ヒート,
      見た目メモ
    ];
  });

  if (出力データ.length > 0) {
    出力シート.getRange(2, 1, 出力データ.length, 8).setValues(出力データ);

    出力シート.getRange(2, 1, 出力データ.length, 1).setNumberFormat('0');
    出力シート.getRange(2, 3, 出力データ.length, 1).setNumberFormat('0');
    出力シート.getRange(2, 4, 出力データ.length, 3).setNumberFormat('0.00');
  }

  出力シート.setFrozenRows(1);
  出力シート.autoResizeColumns(1, 8);

  ヒート列に条件付き書式を設定する_(出力シート, 出力データ.length);

  Logger.log('ランキング滞在ヒートマップ 作成完了: ' + 出力データ.length + '件');
}

function ヒート長さを計算する_(滞在スコア, 最大滞在スコア) {
  if (!最大滞在スコア || 最大滞在スコア <= 0) return 1;

  const 比率 = 滞在スコア / 最大滞在スコア;

  if (比率 >= 0.9) return 10;
  if (比率 >= 0.8) return 9;
  if (比率 >= 0.7) return 8;
  if (比率 >= 0.6) return 7;
  if (比率 >= 0.5) return 6;
  if (比率 >= 0.4) return 5;
  if (比率 >= 0.3) return 4;
  if (比率 >= 0.2) return 3;
  if (比率 >= 0.1) return 2;
  return 1;
}

function ヒートメモを作成する_(出現回数, 平均順位, 最新順位) {
  if (出現回数 >= 8 && 平均順位 <= 15 && 最新順位 <= 15) {
    return 'かなり強い';
  }

  if (出現回数 >= 5 && 最新順位 <= 20) {
    return '継続監視';
  }

  if (最新順位 <= 15) {
    return '上位に残存';
  }

  return '様子見';
}

function ヒート列に条件付き書式を設定する_(シート, データ件数) {
  if (データ件数 <= 0) return;

  const ルール一覧 = シート.getConditionalFormatRules().filter(function(rule) {
    const 範囲一覧 = rule.getRanges();
    for (let i = 0; i < 範囲一覧.length; i++) {
      if (範囲一覧[i].getSheet().getName() === シート.getName()) {
        return false;
      }
    }
    return true;
  });

  const ヒート範囲 = シート.getRange(2, 7, データ件数, 1);

  const ルール = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('████████')
    .setBackground('#f4cccc')
    .setRanges([ヒート範囲])
    .build();

  const ルール2 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('█████')
    .setBackground('#fce5cd')
    .setRanges([ヒート範囲])
    .build();

  const ルール3 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('██')
    .setBackground('#fff2cc')
    .setRanges([ヒート範囲])
    .build();

  ルール一覧.push(ルール, ルール2, ルール3);
  シート.setConditionalFormatRules(ルール一覧);
}