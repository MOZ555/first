function Yahoo売買代金ランキングを確認する() {
  const 行一覧 = Yahoo売買代金ランキングを取得する();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const 確認シート = ss.getSheetByName(設定.比較確認シート名) || ss.insertSheet(設定.比較確認シート名);

  確認シート.clear();

  const 現在 = new Date();
  const タイムゾーン = Session.getScriptTimeZone();
  const 取得時刻 = Utilities.formatDate(現在, タイムゾーン, 'yyyy/MM/dd HH:mm:ss');

  確認シート.getRange(1, 1, 1, 2).setValues([[
    '取得日時',
    取得時刻
  ]]);

  確認シート.getRange(3, 1, 1, 9).setValues([[
    '順位',
    '銘柄コード',
    '銘柄名',
    '市場',
    '株価',
    '日付',
    '前日比値',
    '前日比率',
    '売買代金'
  ]]);

  if (行一覧.length > 0) {
    const 出力データ = 行一覧.map(function(行) {
      return [
        行.順位,
        行.銘柄コード,
        行.銘柄名,
        行.市場,
        行.株価,
        行.日付,
        行.前日比値,
        行.前日比率,
        行.売買代金
      ];
    });

    確認シート.getRange(4, 1, 出力データ.length, 出力データ[0].length).setValues(出力データ);
  }

  確認シート.autoResizeColumns(1, 9);

  Logger.log('Yahoo売買代金ランキング確認: ' + 行一覧.length + '件');
  Logger.log(JSON.stringify(行一覧.slice(0, 5), null, 2));

  return 行一覧;
}