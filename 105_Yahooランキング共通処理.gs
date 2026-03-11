const 設定 = {
  Yahoo売買代金ランキングURL: 'https://finance.yahoo.co.jp/stocks/ranking/tradingValueHigh?market=all',
  ランキングDBシート名: 'ランキングDB',
  比較確認シート名: 'Yahoo取得確認',
  ユーザーエージェント: 'Mozilla/5.0 (compatible; GoogleAppsScript; IKKO-system/1.0)',
};

function URLテキストを取得する_(url) {
  const レスポンス = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true,
    followRedirects: true,
    headers: {
      'User-Agent': 設定.ユーザーエージェント,
      'Accept-Language': 'ja,en-US;q=0.9,en;q=0.8',
      'Cache-Control': 'no-cache',
    },
  });

  const ステータスコード = レスポンス.getResponseCode();
  if (ステータスコード !== 200) {
    throw new Error('HTTP ' + ステータスコード + ': ' + url);
  }

  return レスポンス.getContentText('UTF-8');
}

function HTMLをプレーンテキスト化する_(html) {
  return html
    .replace(/<script[\s\S]*?<\/script>/gi, '\n')
    .replace(/<style[\s\S]*?<\/style>/gi, '\n')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n')
    .replace(/<\/div>/gi, '\n')
    .replace(/<\/li>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/\r/g, '')
    .replace(/\n{2,}/g, '\n')
    .trim();
}

function 数値化する_(value) {
  if (value === null || value === undefined || value === '') return null;
  const 数値 = Number(String(value).replace(/,/g, ''));
  return Number.isNaN(数値) ? null : 数値;
}

function パーセントを数値化する_(value) {
  if (!value) return null;
  const 数値 = Number(String(value).replace('%', ''));
  return Number.isNaN(数値) ? null : 数値;
}

function 市場時間内か判定する_() {
  const now = new Date();
  const timezone = Session.getScriptTimeZone();

  const weekday = Number(Utilities.formatDate(now, timezone, 'u')); // 1=月 ... 7=日

  // 土日
  if (weekday === 6 || weekday === 7) {
    return false;
  }

  // 日本の祝日
  if (日本の祝日か判定する_(now)) {
    return false;
  }

  const hhmm = Number(Utilities.formatDate(now, timezone, 'HHmm'));

  // 前場 09:00-11:30
  if (hhmm >= 900 && hhmm <= 1130) {
    return true;
  }

  // 後場 12:30-15:00
  if (hhmm >= 1230 && hhmm <= 1500) {
    return true;
  }

  return false;
}

function 日本の祝日か判定する_(date) {
  const calendarId = 'ja.japanese#holiday@group.v.calendar.google.com';
  const calendar = CalendarApp.getCalendarById(calendarId);

  const start = new Date(date);
  start.setHours(0, 0, 0, 0);

  const end = new Date(date);
  end.setHours(23, 59, 59, 999);

  const events = calendar.getEvents(start, end);

  return events.length > 0;
}