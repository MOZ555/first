function Yahoo売買代金ランキングデータを取得する() {
  const url = 'https://finance.yahoo.co.jp/stocks/ranking/tradingValueHigh?market=all';

  const response = UrlFetchApp.fetch(url, {
    headers: { 'User-Agent': 'Mozilla/5.0' }
  });

  const html = response.getContentText();

  Logger.log('HTTP code: ' + response.getResponseCode());
  Logger.log('HTML length: ' + html.length);

  const startKey = 'window.__PRELOADED_STATE__ = ';
  const start = html.indexOf(startKey);
  if (start === -1) {
    throw new Error('PRELOADED_STATE開始位置が見つかりません');
  }

  const end = html.indexOf('</script>', start);
  const jsonText = html
    .substring(start + startKey.length, end)
    .trim()
    .replace(/;$/, '');

  const json = JSON.parse(jsonText);

  const ranking = json.mainRankingList?.results || [];

  Logger.log('Yahoo取得件数: ' + ranking.length);

  return ranking.map((r, i) => ({
    rank: Number(r.rank || (i + 1)),
    code: r.stockCode || '',
    name: r.stockName || '',
    market: r.marketName || '',
    price: Number(String(r.savePrice || '').replace(/,/g, '')) || '',
    changeRate: Number(String(r.rankingResult?.tradingValue?.changePriceRate || '').replace(/,/g, '')) || '',
    tradingValue: Number(String(r.rankingResult?.tradingValue?.tradingValue || '').replace(/,/g, '')) || '',
    yahooUpdateTime: r.rankingResult?.tradingValue?.updateDateTime || ''
  }));
}
function Yahoo売買代金ランキングを取得する() {
  return Yahoo売買代金ランキングデータを取得する().ランキング一覧;
}