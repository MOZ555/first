function Yahooランキング末尾確認をする() {
  const html = URLテキストを取得する_(設定.Yahoo売買代金ランキングURL);

  const text = HTMLをプレーンテキスト化する_(html)
    .replace(/\s+/g, ' ')
    .trim();

  const 開始文字 = '順位名称・コード・市場取引値前日比売買代金';
  const 開始位置 = text.indexOf(開始文字);
  if (開始位置 < 0) {
    Logger.log('開始文字が見つかりません');
    return;
  }

  const 本文 = text.substring(開始位置 + 開始文字.length);

  Logger.log('--- 本文末尾1000文字 ---');
  Logger.log(本文.slice(-1000));
}

function Yahooランキング49位50位確認() {
  const url = "https://finance.yahoo.co.jp/stocks/ranking/tradingValueHigh?market=all";

  const res = UrlFetchApp.fetch(url, {
    method: "get",
    muteHttpExceptions: true,
    headers: {
      "User-Agent": "Mozilla/5.0",
      "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
      "Cache-Control": "no-cache"
    }
  });

  Logger.log("HTTPコード: " + res.getResponseCode());

  const html = res.getContentText("UTF-8");

  const text = html
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<[^>]*>/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  const startMarker = "順位名称・コード・市場取引値前日比売買代金";
  const start = text.indexOf(startMarker);

  if (start < 0) {
    Logger.log("開始文字が見つからん");
    Logger.log(text.substring(0, 2000));
    return;
  }

  const body = text.substring(start + startMarker.length);

  const pos49 = body.indexOf("49");
  const pos50 = body.indexOf("50");

  Logger.log("49位置: " + pos49);
  Logger.log("50位置: " + pos50);

  if (pos49 > 0) {
    Logger.log("---49付近---");
    Logger.log(body.substring(Math.max(0, pos49 - 200), pos49 + 1200));
  }

  if (pos50 > 0) {
    Logger.log("---50付近---");
    Logger.log(body.substring(Math.max(0, pos50 - 200), pos50 + 1200));
  }
}

function Yahoo売買代金ランキングを確認する() {

  const url = "https://finance.yahoo.co.jp/stocks/ranking/tradingValueHigh?market=all";

  const res = UrlFetchApp.fetch(url, {
    method: "get",
    headers: {
      "User-Agent": "Mozilla/5.0",
      "Accept-Language": "ja,en-US;q=0.9,en;q=0.8"
    }
  });

  const html = res.getContentText("UTF-8");

  const rows = Yahoo売買代金ランキングを解析する_(html);

  Logger.log("取得件数: " + rows.length);

  rows.forEach(function(r){
    Logger.log(
      r.順位 + " | " +
      r.銘柄コード + " | " +
      r.銘柄名 + " | " +
      r.売買代金
    );
  });

}

function ランキングDBヘッダー確認() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('ランキングDB');
  if (!sheet) throw new Error('ランキングDB シートがありません');

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  Logger.log(JSON.stringify(headers));
}