/**
 * Mutual Fund Valuation & Multi-Tenure Return Matrix.
 * Apps Script file. The HTML file in this project must be named Index.
 *
 * Follow these rules when changing this file:
 * - Read sheets scheme_codes and nav_data. Never invent a NAV.
 * - Join on scheme_code. scheme_codes supplies the fund name; nav_data supplies date and NAV.
 * - lookupNav uses the last NAV on or before the requested day.
 * - Tenure columns look back from Sale Date only. Do not clip them with From Date or To Date.
 * - Custom Range XIRR, SIP Corpus, Net Profit, and Absolute return use From Date through To Date only.
 * - Keep TENURE_CONFIG in this order, including 7Y. The same keys and order live in index.html TENURES.
 * - A tenure is null when the scheme has no NAV on or before the window start.
 *   Do not relabel a shorter history as 12Y. SIP XIRR is also null when the window contains no SIP date.
 * - 5D and 15D have no tenure XIRR. A monthly SIP does not fit those windows. Trailing CAGR still uses them.
 * - fund.risk is a 3-year window ending on Sale Date, from daily NAV. Cash rate is 6.5%.
 *   Capture, beta, alpha, information ratio, and R-squared use the other tracked funds in the same category, not Nifty.
 *   volatility is the annualized standard deviation.
 *   Also rolling 1Y return and hit rate, 95% one-day VaR, Treynor, correlation, and peer tracking error and difference.
 *   Tracking uses category peers, not an index. Expense ratio is not on the NAV sheet.
 *   Manager, tenure, and portfolio weights are not on the NAV sheet. getFundPublishedProfile loads them when the public profile responds.
 *   getFundOverlap loads the latest published equity portfolio for at most 20 schemes and returns pairwise overlap.
 *   Overlap is the sum of the smaller weight of each shared stock, divided by 100. Cash and debt are excluded.
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Mutual Fund Matrix')
    .addItem('📊 Open Return Matrix (Sidebar)', 'showSidebar')
    .addItem('🖥️ Open Return Matrix (Full Dialog)', 'showDialog')
    .addSeparator()
    .addItem('🌐 View Published Web App Link', 'showWebAppUrl')
    .addToUi();
}

function showSidebar() {
  var html = HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Mutual Fund Valuation & Return Matrix');
  SpreadsheetApp.getUi().showSidebar(html);
}

function showDialog() {
  var html = HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setWidth(1250)
    .setHeight(820);
  SpreadsheetApp.getUi().showModalDialog(html, 'Mutual Fund Valuation & Multi-Tenure Matrix');
}

function showWebAppUrl() {
  var url = ScriptApp.getService().getUrl();
  var ui = SpreadsheetApp.getUi();
  if (url) {
    ui.alert('Web App URL', 'Your published Web App URL is:\n\n' + url, ui.ButtonSet.OK);
  } else {
    ui.alert(
      'Web App Not Yet Deployed',
      'Please deploy the web app first via Deploy > New deployment > Web app to generate a standalone web link.',
      ui.ButtonSet.OK
    );
  }
}

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Mutual Fund Valuation & Return Matrix')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Calendar offsets back from Sale Date. annualized applies to trailing CAGR only.
// SIP XIRR is always an annualized yearly rate, so 1M–9M read higher than the period's price change.
// 5D and 15D are CAGR only. Do not fill tenureXirr for those keys.
// 7Y is intentional. Do not put 8Y back.
var TENURE_CONFIG = [
  { key: '5D', label: '5D', unit: 'day', n: 5, annualized: false, group: 'short' },
  { key: '15D', label: '15D', unit: 'day', n: 15, annualized: false, group: 'short' },
  { key: '1M', label: '1M', unit: 'month', n: 1, annualized: false, group: 'short' },
  { key: '3M', label: '3M', unit: 'month', n: 3, annualized: false, group: 'short' },
  { key: '6M', label: '6M', unit: 'month', n: 6, annualized: false, group: 'short' },
  { key: '9M', label: '9M', unit: 'month', n: 9, annualized: false, group: 'short' },
  { key: '1Y', label: '1Y', unit: 'year', n: 1, annualized: true, group: 'short' },
  { key: '2Y', label: '2Y', unit: 'year', n: 2, annualized: true, group: 'long' },
  { key: '3Y', label: '3Y', unit: 'year', n: 3, annualized: true, group: 'long' },
  { key: '5Y', label: '5Y', unit: 'year', n: 5, annualized: true, group: 'long' },
  { key: '7Y', label: '7Y', unit: 'year', n: 7, annualized: true, group: 'long' },
  { key: '10Y', label: '10Y', unit: 'year', n: 10, annualized: true, group: 'long' },
  { key: '12Y', label: '12Y', unit: 'year', n: 12, annualized: true, group: 'long' }
];

/**
 * Returns one row per scheme_codes entry.
 * Defaults: SIP 1000 on day 5, From = today minus 1 year, To = today, Sale = today.
 */
function getFundMatrixData(params) {
  params = params || {};
  var today = todayDay();
  var sipAmount = Number(params.sipAmount) || 1000;
  var sipDay = Number(params.sipDay) || 5;
  var fromDate = params.fromDate ? toDay(params.fromDate) : shiftBack(today, 'year', 1);
  var toDate = params.toDate ? toDay(params.toDate) : today;
  var saleDate = params.saleDate ? toDay(params.saleDate) : today;

  if (sipDay < 1) sipDay = 1;
  if (sipDay > 31) sipDay = 31;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var schemeSheet = ss.getSheetByName('scheme_codes');
  if (!schemeSheet) {
    throw new Error('Sheet "scheme_codes" not found in the spreadsheet.');
  }

  var rawSchemes = schemeSheet.getDataRange().getValues();
  if (rawSchemes.length <= 1) {
    return { funds: [], categories: [], amcs: [], tenures: TENURE_CONFIG };
  }

  var headers = rawSchemes[0].map(function(h) { return String(h).trim().toLowerCase(); });
  var codeIdx = headers.indexOf('scheme_code');
  if (codeIdx === -1) codeIdx = 0;
  var nameIdx = headers.indexOf('scheme_name');
  if (nameIdx === -1) nameIdx = 1;
  var amcIdx = headers.indexOf('amc_name');
  var catIdx = headers.indexOf('fund_category');
  var planIdx = headers.indexOf('plan');
  var optIdx = headers.indexOf('option');
  var minNavIdx = headers.indexOf('min_nav_date');

  var schemes = [];
  var wanted = {};
  for (var i = 1; i < rawSchemes.length; i++) {
    var row = rawSchemes[i];
    var code = String(row[codeIdx] || '').trim();
    if (!code) continue;
    wanted[code] = true;
    schemes.push({
      code: code,
      name: String(row[nameIdx] || 'Unnamed Fund').trim(),
      amc: amcIdx !== -1 ? String(row[amcIdx] || '').trim() : '',
      category: catIdx !== -1 ? String(row[catIdx] || 'Equity').trim() : 'Equity',
      plan: planIdx !== -1 ? String(row[planIdx] || 'Direct').trim() : 'Direct',
      option: optIdx !== -1 ? String(row[optIdx] || 'Growth').trim() : 'Growth',
      minNavDate: minNavIdx !== -1 && row[minNavIdx] ? formatDay(toDay(row[minNavIdx])) : ''
    });
  }

  // Index only the scheme codes listed on scheme_codes. Match by code, not by fund name.
  var navIndex = loadNavIndex(ss, wanted);
  var categoriesSet = {};
  var amcSet = {};
  var fundList = [];

  for (var s = 0; s < schemes.length; s++) {
    var scheme = schemes[s];
    var series = navIndex[scheme.code] || null;
    if (scheme.category) categoriesSet[scheme.category] = true;
    if (scheme.amc) amcSet[scheme.amc] = true;

    // Custom range is From Date through To Date. Sale Date is not an input to these figures.
    var custom = buildSip(series, sipAmount, sipDay, fromDate, toDate);
    var navStart = lookupNav(series, fromDate);
    var navEnd = lookupNav(series, toDate);
    var saleNav = lookupNav(series, saleDate);
    var tenureXirr = {};
    var tenureCagr = {};

    for (var t = 0; t < TENURE_CONFIG.length; t++) {
      var tenure = TENURE_CONFIG[t];
      // Full window required. A fund that starts inside the window does not get this tenure.
      var windowStart = shiftBack(saleDate, tenure.unit, tenure.n);
      var covered = lookupNav(series, windowStart) !== null;
      if (!covered) {
        tenureXirr[tenure.key] = null;
        tenureCagr[tenure.key] = null;
        continue;
      }
      var sip = buildSip(series, sipAmount, sipDay, windowStart, saleDate);
      tenureXirr[tenure.key] = (tenure.key === '5D' || tenure.key === '15D') ? null : sip.xirr;
      tenureCagr[tenure.key] = trailingCagr(series, windowStart, saleDate, tenure.annualized);
    }

    if (!scheme.minNavDate && series && series.dates.length) {
      scheme.minNavDate = formatDay(new Date(series.dates[0]));
    }

    fundList.push({
      code: scheme.code,
      name: scheme.name,
      amc: scheme.amc,
      category: scheme.category,
      type: scheme.plan + ' · ' + scheme.option,
      minNavDate: scheme.minNavDate,
      navStart: roundNav(navStart),
      navEnd: roundNav(navEnd),
      saleNav: roundNav(saleNav),
      units: custom.units !== null ? Number(custom.units.toFixed(4)) : null,
      investedAmount: custom.invested !== null ? roundRupee(custom.invested) : null,
      sipCorpus: custom.corpus !== null ? roundRupee(custom.corpus) : null,
      netProfit: custom.profit !== null ? roundRupee(custom.profit) : null,
      absoluteReturn: custom.absoluteReturn,
      customRangeXirr: custom.xirr,
      tenureXirr: tenureXirr,
      tenureCagr: tenureCagr,
      redeem: custom.redeem || null,
      switchPath: threeYearPath(series, saleDate),
      risk: buildRiskStats(series, saleDate)
    });
  }

  attachPeerRisk(fundList, navIndex, saleDate);

  var pathMonths = [];
  for (var back = 35; back >= 0; back--) {
    pathMonths.push(formatDay(shiftBack(saleDate, 'month', back)).slice(0, 7));
  }

  return {
    funds: fundList,
    categories: Object.keys(categoriesSet).sort(),
    amcs: Object.keys(amcSet).sort(),
    tenures: TENURE_CONFIG,
    pathMonths: pathMonths
  };
}

// Risk window and the cash rate used inside Sharpe and Sortino. Do not change these silently.
var RISK_YEARS = 3;
var RISK_FREE = 0.065;

function buildRiskStats(series, saleDate) {
  var empty = blankRisk(saleDate);
  var pts = windowPoints(series, shiftBack(saleDate, 'year', RISK_YEARS), saleDate);
  if (pts.length < 61) return empty;
  var rets = returnSeries(pts);
  if (rets.length < 60) return empty;
  var values = rets.map(function(item) { return item.r; });
  var vol = stdev(values);
  var downside = downsideDev(values, RISK_FREE / 252);
  var cagr = trailingCagr(series, new Date(pts[0].t), saleDate, true);
  empty.windowStart = formatDay(new Date(pts[0].t));
  empty.windowEnd = formatDay(saleDate);
  empty.observations = rets.length;
  empty.cagr = cagr;
  empty.volatility = vol === null ? null : vol * Math.sqrt(252);
  empty.downsideDeviation = downside === null ? null : downside * Math.sqrt(252);
  if (cagr !== null && empty.volatility) empty.sharpe = (cagr - RISK_FREE) / empty.volatility;
  if (cagr !== null && empty.downsideDeviation) empty.sortino = (cagr - RISK_FREE) / empty.downsideDeviation;
  var path = drawPath(pts);
  empty.maxDrawdown = path.drawdown;
  empty.maxDrawup = path.drawup;
  var months = monthReturns(pts);
  if (months.length) {
    var best = months[0];
    var worst = months[0];
    var wins = 0;
    months.forEach(function(item) {
      if (item.r > best.r) best = item;
      if (item.r < worst.r) worst = item;
      if (item.r > 0) wins += 1;
    });
    empty.bestMonth = best.r;
    empty.worstMonth = worst.r;
    empty.positiveMonthShare = wins / months.length;
  }
  var sorted = values.slice().sort(function(a, b) { return a - b; });
  empty.var95 = sorted[Math.floor(0.05 * (sorted.length - 1))];
  var rolling = rollingYearStats(series, saleDate);
  empty.rollingReturn = rolling.rollingReturn;
  empty.rollingHit = rolling.rollingHit;
  return empty;
}

function blankRisk(saleDate) {
  return {
    windowStart: null,
    windowEnd: saleDate ? formatDay(saleDate) : null,
    observations: 0,
    cagr: null,
    volatility: null,
    downsideDeviation: null,
    sharpe: null,
    sortino: null,
    maxDrawdown: null,
    maxDrawup: null,
    bestMonth: null,
    worstMonth: null,
    positiveMonthShare: null,
    upsideCapture: null,
    downsideCapture: null,
    beta: null,
    alpha: null,
    informationRatio: null,
    rSquared: null,
    correlation: null,
    trackingError: null,
    trackingDifference: null,
    treynor: null,
    var95: null,
    rollingReturn: null,
    rollingHit: null
  };
}

function rollingYearStats(series, saleDate) {
  var out = { rollingReturn: null, rollingHit: null };
  if (!series || !series.dates.length) return out;
  var yearMs = 365.25 * 86400000;
  var endMs = saleDate.getTime();
  var firstMs = endMs - 4 * yearMs;
  var windows = [];
  var backIndex = 0;
  var lastSample = 0;
  for (var i = 0; i < series.dates.length; i++) {
    var t = series.dates[i];
    if (t > endMs) break;
    if (t < firstMs) continue;
    if (lastSample && t - lastSample < 28 * 86400000) continue;
    var back = t - yearMs;
    while (backIndex + 1 < i && series.dates[backIndex + 1] <= back) backIndex += 1;
    if (series.dates[backIndex] > back || !(series.navs[backIndex] > 0) || !(series.navs[i] > 0)) continue;
    windows.push(series.navs[i] / series.navs[backIndex] - 1);
    lastSample = t;
  }
  if (windows.length < 6) return out;
  var hits = 0;
  windows.forEach(function(value) { if (value > 0) hits += 1; });
  out.rollingReturn = mean(windows);
  out.rollingHit = hits / windows.length;
  return out;
}

function attachPeerRisk(funds, navIndex, saleDate) {
  var byCat = {};
  funds.forEach(function(fund) {
    var category = fund.category || 'Other';
    if (!byCat[category]) byCat[category] = [];
    byCat[category].push(fund.code);
  });
  var start = shiftBack(saleDate, 'year', RISK_YEARS);
  funds.forEach(function(fund) {
    if (!fund.risk) fund.risk = blankRisk(saleDate);
    var peers = [];
    (byCat[fund.category || 'Other'] || []).forEach(function(code) {
      if (code !== fund.code && navIndex[code]) peers.push(navIndex[code]);
    });
    var rel = peerRelative(navIndex[fund.code], peers, start, saleDate);
    fund.risk.upsideCapture = rel.upsideCapture;
    fund.risk.downsideCapture = rel.downsideCapture;
    fund.risk.beta = rel.beta;
    fund.risk.alpha = rel.alpha;
    fund.risk.informationRatio = rel.informationRatio;
    fund.risk.rSquared = rel.rSquared;
    fund.risk.correlation = rel.correlation;
    fund.risk.trackingError = rel.trackingError;
    fund.risk.trackingDifference = rel.trackingDifference;
    if (rel.beta > 0 && fund.risk.cagr !== null && fund.risk.cagr !== undefined) {
      fund.risk.treynor = (fund.risk.cagr - RISK_FREE) / rel.beta;
    }
  });
}

function peerRelative(series, peers, start, end) {
  var out = {
    upsideCapture: null, downsideCapture: null, beta: null, alpha: null,
    informationRatio: null, rSquared: null, correlation: null,
    trackingError: null, trackingDifference: null
  };
  if (!series || !peers.length) return out;
  var mine = returnMap(returnSeries(windowPoints(series, start, end)));
  var peerMaps = peers.map(function(peer) { return returnMap(returnSeries(windowPoints(peer, start, end))); });
  var upF = [];
  var upP = [];
  var downF = [];
  var downP = [];
  var xs = [];
  var ys = [];
  Object.keys(mine).forEach(function(key) {
    var peerVals = [];
    peerMaps.forEach(function(map) {
      if (map[key] !== undefined) peerVals.push(map[key]);
    });
    if (!peerVals.length) return;
    var peer = 0;
    peerVals.forEach(function(value) { peer += value; });
    peer /= peerVals.length;
    var fund = mine[key];
    xs.push(peer);
    ys.push(fund);
    if (peer > 0) { upP.push(peer); upF.push(fund); }
    else if (peer < 0) { downP.push(peer); downF.push(fund); }
  });
  if (xs.length < 30) return out;
  var peerVar = variance(xs);
  if (peerVar) out.beta = covariance(xs, ys) / peerVar;
  var meanF = mean(ys) * 252;
  var meanP = mean(xs) * 252;
  if (out.beta !== null) out.alpha = (meanF - RISK_FREE) - out.beta * (meanP - RISK_FREE);
  if (upP.length >= 10 && mean(upP)) out.upsideCapture = mean(upF) / mean(upP);
  if (downP.length >= 10 && mean(downP)) out.downsideCapture = mean(downF) / mean(downP);
  var fundVar = variance(ys);
  var cov = covariance(xs, ys);
  if (peerVar && fundVar) {
    out.rSquared = Math.pow(cov, 2) / (peerVar * fundVar);
    out.correlation = cov / Math.sqrt(peerVar * fundVar);
  }
  var excess = [];
  for (var i = 0; i < ys.length; i++) excess.push(ys[i] - xs[i]);
  var tracking = stdev(excess);
  var excessMean = mean(excess);
  if (excessMean !== null) out.trackingDifference = excessMean * 252;
  if (tracking) {
    out.trackingError = tracking * Math.sqrt(252);
    out.informationRatio = (excessMean * Math.sqrt(252)) / tracking;
  }
  return out;
}

/**
 * Manager, tenure, and portfolio weights are not in nav_data.
 * This asks the public mfdata.in profile. Morningstar and Dhan have no keyless API;
 * the client opens those sites for the selected fund name when this call cannot fill a field.
 * codes: [{ code, name }], at most 8.
 */
function getFundPublishedProfile(codes) {
  codes = codes || [];
  if (codes.length > 8) codes = codes.slice(0, 8);
  var profiles = [];
  for (var i = 0; i < codes.length; i++) {
    profiles.push(fetchPublishedProfile(codes[i]));
  }
  return {
    profiles: profiles,
    source: 'Published fields come from the public scheme profile when it responds. Morningstar and Dhan open in a new tab for the selected fund.'
  };
}

function fetchPublishedProfile(item) {
  var code = String(item.code || item);
  var name = item.name || code;
  var profile = {
    code: code,
    name: name,
    links: {
      morningstar: 'https://www.google.com/search?q=' + encodeURIComponent(name + ' site:morningstar.in'),
      dhan: 'https://www.google.com/search?q=' + encodeURIComponent(name + ' site:dhan.co/mutual-funds')
    },
    manager: null,
    managerSince: null,
    expenseRatio: null,
    aumCr: null,
    morningstar: null,
    holdings: [],
    publishedRatios: null,
    error: null
  };
  try {
    var response = UrlFetchApp.fetch('https://mfdata.in/api/v1/schemes/' + encodeURIComponent(code), { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) {
      profile.error = 'Published profile is unavailable right now. Use Morningstar or Dhan for manager, tenure, and weights.';
      return profile;
    }
    var body = JSON.parse(response.getContentText());
    var data = body.data || body;
    profile.expenseRatio = numberOrNull(data.expense_ratio);
    profile.aumCr = numberOrNull(data.aum_cr);
    profile.morningstar = data.morningstar || data.rating || null;
    profile.publishedRatios = data.ratios || null;
    if (data.family_id) fillFamilyProfile(profile, data.family_id);
    if (!profile.manager && !profile.holdings.length && !profile.error) {
      profile.error = 'This scheme has no manager or portfolio on the public profile. Use Morningstar or Dhan.';
    }
  } catch (err) {
    profile.error = 'Published profile could not be loaded. Use Morningstar or Dhan.';
  }
  return profile;
}

function fillFamilyProfile(profile, familyId) {
  try {
    var people = UrlFetchApp.fetch('https://mfdata.in/api/v1/families/' + familyId + '/people', { muteHttpExceptions: true });
    if (people.getResponseCode() === 200) {
      var parsed = JSON.parse(people.getContentText());
      var list = parsed.data || parsed;
      if (Object.prototype.toString.call(list) === '[object Array]' && list.length) {
        var person = list[0];
        profile.manager = person.name || person.manager || null;
        profile.managerSince = person.start_date || person.since || person.tenure || null;
      }
    }
  } catch (err) {}
  try {
    var holdings = UrlFetchApp.fetch('https://mfdata.in/api/v1/families/' + familyId + '/holdings', { muteHttpExceptions: true });
    if (holdings.getResponseCode() === 200) {
      var parsedHold = JSON.parse(holdings.getContentText());
      var data = parsedHold.data || parsedHold;
      var equity = data.equity || [];
      equity.sort(function(a, b) { return (b.weight_pct || 0) - (a.weight_pct || 0); });
      profile.holdings = equity.slice(0, 8).map(function(row) {
        return { name: row.name, weight: row.weight_pct, sector: row.sector || '' };
      });
    }
  } catch (err) {}
}

/**
 * Pairwise portfolio overlap for at most 20 schemes.
 * codes: [{ code, name }]. Holdings are the latest published book, not the NAV sheet.
 * A scheme is used only when the published scheme code matches. Weights are not invented.
 */
function getFundOverlap(codes) {
  codes = codes || [];
  if (codes.length > 20) {
    return {
      funds: [],
      matrix: [],
      common: [],
      error: 'Select at most 20 funds.',
      source: overlapSourceNote()
    };
  }
  var funds = [];
  var books = [];
  var queries = [];
  var queryOwner = [];
  for (var i = 0; i < codes.length; i++) {
    var code = String(codes[i].code || codes[i]);
    var name = codes[i].name || code;
    funds.push({
      code: code,
      name: name,
      asOf: null,
      equityCount: 0,
      equityWeight: null,
      topHoldings: [],
      error: null
    });
    books.push(null);
    var variants = overlapQueries(name);
    for (var q = 0; q < variants.length; q++) {
      queries.push(growwSearchUrl(variants[q]));
      queryOwner.push(i);
    }
  }
  var searchHits = overlapFetchJson(queries);
  var grouped = funds.map(function() { return []; });
  for (var h = 0; h < searchHits.length; h++) grouped[queryOwner[h]].push(searchHits[h]);
  var detailUrls = [];
  var detailIndex = [];
  var pending = [];
  for (var s = 0; s < funds.length; s++) {
    var searchId = growwSearchId(grouped[s], funds[s].code);
    if (searchId) {
      detailIndex.push(s);
      detailUrls.push(growwPortfolioUrl(searchId));
    } else {
      pending.push(s);
    }
  }
  var fallbackIds = [];
  var fallbackOwner = [];
  for (var p = 0; p < pending.length; p++) {
    var candidates = growwSchemeCandidates(grouped[pending[p]]);
    for (var cnd = 0; cnd < candidates.length; cnd++) {
      fallbackIds.push(growwPortfolioUrl(candidates[cnd]));
      fallbackOwner.push(pending[p]);
    }
  }
  var details = overlapFetchJson(detailUrls.concat(fallbackIds));
  for (var d = 0; d < detailIndex.length; d++) {
    var idx = detailIndex[d];
    var book = equityBook(details[d], funds[idx].code);
    if (!book) {
      funds[idx].error = 'The published portfolio did not load.';
      continue;
    }
    if (!book.count) {
      funds[idx].error = 'This published portfolio has no equity holdings.';
      funds[idx].asOf = book.asOf;
      continue;
    }
    fillEquityFund(funds[idx], book);
    books[idx] = book.map;
  }
  for (var f = 0; f < fallbackOwner.length; f++) {
    var owner = fallbackOwner[f];
    if (books[owner] || funds[owner].error) continue;
    var fallbackBook = equityBook(details[detailIndex.length + f], funds[owner].code);
    if (!fallbackBook) continue;
    if (!fallbackBook.count) {
      funds[owner].error = 'This published portfolio has no equity holdings.';
      funds[owner].asOf = fallbackBook.asOf;
      continue;
    }
    fillEquityFund(funds[owner], fallbackBook);
    books[owner] = fallbackBook.map;
  }
  for (var missed = 0; missed < funds.length; missed++) {
    if (!books[missed] && !funds[missed].error) {
      funds[missed].error = 'No published portfolio matched this scheme code.';
    }
  }
  var n = funds.length;
  var matrix = [];
  var common = [];
  for (var r = 0; r < n; r++) {
    matrix[r] = [];
    common[r] = [];
    for (var c = 0; c < n; c++) {
      if (r === c || !books[r] || !books[c]) {
        matrix[r][c] = null;
        common[r][c] = 0;
      } else {
        var pair = pairOverlap(books[r], books[c]);
        matrix[r][c] = pair.overlap;
        common[r][c] = pair.common;
      }
    }
  }
  return { funds: funds, matrix: matrix, common: common, source: overlapSourceNote() };
}

function overlapSourceNote() {
  return 'Overlap uses the latest published equity portfolio. It is the sum of the smaller weight of each stock held by both funds.';
}

function overlapQueries(name) {
  var queries = [];
  function add(value) {
    var text = String(value || '').replace(/\s+/g, ' ').trim();
    if (text && queries.indexOf(text) === -1) queries.push(text);
  }
  add(name);
  var base = String(name || '').split(/\s+-\s+/)[0].replace(/\s*-\s*/g, ' ');
  add(base);
  add(base + ' Direct');
  return queries;
}

function growwSearchUrl(name) {
  return 'https://groww.in/v1/api/search/v3/query/global/st_query?page=0&size=12&web=true&query=' + encodeURIComponent(name);
}

function growwPortfolioUrl(searchId) {
  return 'https://groww.in/v1/api/data/mf/web/v4/scheme/search/' + encodeURIComponent(searchId);
}

function overlapFetchJson(urls) {
  if (!urls.length) return [];
  var requests = urls.map(function(url) {
    return {
      url: url,
      muteHttpExceptions: true,
      headers: { Accept: 'application/json', 'User-Agent': 'Mozilla/5.0' }
    };
  });
  var responses = UrlFetchApp.fetchAll(requests);
  return responses.map(function(response) {
    if (response.getResponseCode() !== 200) return null;
    try {
      return JSON.parse(response.getContentText());
    } catch (err) {
      return null;
    }
  });
}

function growwSearchId(bodies, code) {
  var want = String(code);
  for (var b = 0; b < bodies.length; b++) {
    var content = bodies[b] && bodies[b].data && bodies[b].data.content;
    if (!content) continue;
    for (var i = 0; i < content.length; i++) {
      var item = content[i];
      if (item && item.entity_type === 'Scheme' && String(item.scheme_code) === want && item.search_id) {
        return item.search_id;
      }
    }
  }
  return null;
}

function growwSchemeCandidates(bodies) {
  var ids = [];
  for (var b = 0; b < bodies.length; b++) {
    var content = bodies[b] && bodies[b].data && bodies[b].data.content;
    if (!content) continue;
    for (var i = 0; i < content.length; i++) {
      var item = content[i];
      if (!item || item.entity_type !== 'Scheme' || !item.search_id) continue;
      if (ids.indexOf(item.search_id) === -1) ids.push(item.search_id);
      if (ids.length >= 4) return ids;
    }
  }
  return ids;
}

function publishedCodeMatches(body, code) {
  if (!body) return false;
  var want = String(code);
  return String(body.scheme_code) === want ||
    String(body.direct_scheme_code || '') === want ||
    String(body.regular_scheme_code || '') === want;
}

function equityBook(body, code) {
  if (!publishedCodeMatches(body, code)) return null;
  var holdings = body.holdings || [];
  var map = {};
  var labels = {};
  var count = 0;
  var equityWeight = 0;
  var asOf = null;
  for (var i = 0; i < holdings.length; i++) {
    var row = holdings[i];
    if (String(row.nature_name || '').toUpperCase() !== 'EQUITY') continue;
    var key = String(row.stock_search_id || row.company_name || '').toLowerCase().replace(/[^a-z0-9]+/g, '');
    if (!key) continue;
    var weight = Number(row.corpus_per);
    if (isNaN(weight) || weight <= 0) continue;
    if (!map[key]) {
      map[key] = weight;
      labels[key] = row.company_name || key;
      count += 1;
    } else {
      map[key] += weight;
    }
    equityWeight += weight;
    if (!asOf && row.portfolio_date) asOf = portfolioDay(row.portfolio_date);
  }
  var top = Object.keys(map).map(function(key) {
    return { name: labels[key], weight: map[key] };
  });
  top.sort(function(a, b) { return b.weight - a.weight; });
  return { map: map, count: count, equityWeight: equityWeight / 100, asOf: asOf, top: top.slice(0, 3) };
}

function fillEquityFund(fund, book) {
  fund.asOf = book.asOf;
  fund.equityCount = book.count;
  fund.equityWeight = book.equityWeight;
  fund.topHoldings = book.top || [];
}

function portfolioDay(iso) {
  try {
    return Utilities.formatDate(new Date(iso), 'Asia/Kolkata', 'yyyy-MM-dd');
  } catch (err) {
    return String(iso).slice(0, 10);
  }
}

function pairOverlap(left, right) {
  var sum = 0;
  var shared = 0;
  var keys = Object.keys(left);
  for (var i = 0; i < keys.length; i++) {
    var other = right[keys[i]];
    if (other === undefined) continue;
    sum += Math.min(left[keys[i]], other);
    shared += 1;
  }
  return { overlap: sum / 100, common: shared };
}

function numberOrNull(value) {
  var n = Number(value);
  return isNaN(n) ? null : n;
}

function windowPoints(series, start, end) {
  if (!series || !start || !end) return [];
  var startMs = start.getTime();
  var endMs = end.getTime();
  var base = null;
  var pts = [];
  for (var i = 0; i < series.dates.length; i++) {
    var t = series.dates[i];
    if (t <= startMs) base = { t: t, nav: series.navs[i] };
    else if (t <= endMs) pts.push({ t: t, nav: series.navs[i] });
    else break;
  }
  if (base) pts.unshift(base);
  return pts;
}

function returnSeries(pts) {
  var out = [];
  for (var i = 1; i < pts.length; i++) {
    if (pts[i - 1].nav > 0) out.push({ t: pts[i].t, r: pts[i].nav / pts[i - 1].nav - 1 });
  }
  return out;
}

function returnMap(rets) {
  var map = {};
  rets.forEach(function(item) { map[String(item.t)] = item.r; });
  return map;
}

function drawPath(pts) {
  var peak = pts[0].nav;
  var trough = pts[0].nav;
  var drawdown = 0;
  var drawup = 0;
  for (var i = 1; i < pts.length; i++) {
    var nav = pts[i].nav;
    if (nav > peak) peak = nav;
    if (nav < trough) trough = nav;
    if (peak > 0) {
      var dd = nav / peak - 1;
      if (dd < drawdown) drawdown = dd;
    }
    if (trough > 0) {
      var du = nav / trough - 1;
      if (du > drawup) drawup = du;
    }
  }
  return { drawdown: drawdown, drawup: drawup };
}

function monthReturns(pts) {
  var out = [];
  var monthKey = null;
  var monthStart = null;
  var lastNav = null;
  for (var i = 0; i < pts.length; i++) {
    var date = new Date(pts[i].t);
    var key = date.getFullYear() + '-' + date.getMonth();
    if (monthKey === null) {
      monthKey = key;
      monthStart = pts[i].nav;
      lastNav = pts[i].nav;
      continue;
    }
    if (key !== monthKey) {
      if (monthStart > 0) out.push({ r: lastNav / monthStart - 1 });
      monthKey = key;
      monthStart = pts[i].nav;
    }
    lastNav = pts[i].nav;
  }
  if (monthStart > 0 && lastNav !== null) out.push({ r: lastNav / monthStart - 1 });
  return out;
}

function mean(values) {
  if (!values.length) return null;
  var sum = 0;
  values.forEach(function(value) { sum += value; });
  return sum / values.length;
}

function variance(values) {
  if (values.length < 2) return null;
  var avg = mean(values);
  var sum = 0;
  values.forEach(function(value) { sum += (value - avg) * (value - avg); });
  return sum / (values.length - 1);
}

function covariance(xs, ys) {
  var avgX = mean(xs);
  var avgY = mean(ys);
  var sum = 0;
  for (var i = 0; i < xs.length; i++) sum += (xs[i] - avgX) * (ys[i] - avgY);
  return sum / (xs.length - 1);
}

function stdev(values) {
  var v = variance(values);
  return v === null ? null : Math.sqrt(v);
}

function downsideDev(values, mar) {
  var sum = 0;
  var n = 0;
  values.forEach(function(value) {
    var gap = value - mar;
    if (gap < 0) sum += gap * gap;
    n += 1;
  });
  if (n < 2) return null;
  return Math.sqrt(sum / (n - 1));
}

/** Groups nav_data into { dates, navs } per scheme_code. Dates are local midnights, sorted ascending. */
function loadNavIndex(ss, wanted) {
  var sheet = ss.getSheetByName('nav_data');
  if (!sheet) {
    throw new Error('Sheet "nav_data" not found in the spreadsheet.');
  }

  var values = sheet.getDataRange().getValues();
  var index = {};
  if (values.length <= 1) return index;

  var headers = values[0].map(function(h) { return String(h).trim().toLowerCase(); });
  var codeIdx = headers.indexOf('scheme_code');
  var dateIdx = headers.indexOf('date');
  var navIdx = headers.indexOf('nav');
  if (codeIdx === -1) codeIdx = 0;
  if (dateIdx === -1) dateIdx = 2;
  if (navIdx === -1) navIdx = 3;

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var code = String(row[codeIdx] || '').trim();
    if (!code || !wanted[code]) continue;
    var nav = Number(row[navIdx]);
    if (!isFinite(nav) || nav <= 0) continue;
    var day = toDay(row[dateIdx]);
    if (!day || isNaN(day.getTime())) continue;
    if (!index[code]) index[code] = { dates: [], navs: [] };
    index[code].dates.push(day.getTime());
    index[code].navs.push(nav);
  }

  var codes = Object.keys(index);
  for (var c = 0; c < codes.length; c++) {
    sortSeries(index[codes[c]]);
  }
  return index;
}

function sortSeries(series) {
  var order = [];
  for (var i = 0; i < series.dates.length; i++) order.push(i);
  order.sort(function(a, b) { return series.dates[a] - series.dates[b]; });

  var dates = [];
  var navs = [];
  for (var j = 0; j < order.length; j++) {
    var stamp = series.dates[order[j]];
    var nav = series.navs[order[j]];
    if (dates.length && dates[dates.length - 1] === stamp) {
      navs[navs.length - 1] = nav;
    } else {
      dates.push(stamp);
      navs.push(nav);
    }
  }
  series.dates = dates;
  series.navs = navs;
}

/** Last published NAV on or before dateObj. Null when the scheme did not exist yet. */
function lookupNav(series, dateObj) {
  if (!series || !series.dates.length || !dateObj) return null;
  var target = dateObj.getTime();
  var lo = 0;
  var hi = series.dates.length - 1;
  var found = -1;
  while (lo <= hi) {
    var mid = (lo + hi) >> 1;
    if (series.dates[mid] <= target) {
      found = mid;
      lo = mid + 1;
    } else {
      hi = mid - 1;
    }
  }
  if (found < 0) return null;
  return series.navs[found];
}

/**
 * Monthly SIP from startDate through endDate, redeemed at the endDate NAV.
 * Cash flows are negative installments plus one positive corpus.
 * Returns xirr null when no installment falls in the window.
 */
function buildSip(series, sipAmount, sipDay, startDate, endDate) {
  var empty = { units: 0, invested: 0, corpus: null, profit: null, absoluteReturn: null, xirr: null, redeem: null };
  if (!series || !startDate || !endDate || endDate < startDate) return empty;

  var endNav = lookupNav(series, endDate);
  if (endNav === null) return empty;

  var units = 0;
  var invested = 0;
  var flows = [];
  var lots = [];
  var cur = firstSipOnOrAfter(startDate, sipDay);

  while (cur && cur.getTime() <= endDate.getTime()) {
    var nav = lookupNav(series, cur);
    if (nav !== null && nav > 0) {
      var bought = sipAmount / nav;
      units += bought;
      invested += sipAmount;
      lots.push({ t: cur.getTime(), cost: sipAmount, units: bought });
      flows.push({ date: new Date(cur.getTime()), amount: -sipAmount });
    }
    cur = nextSipDate(cur, sipDay);
  }

  if (invested <= 0 || units <= 0) return empty;

  var corpus = units * endNav;
  flows.push({ date: new Date(endDate.getTime()), amount: corpus });
  var stcgGain = 0;
  var ltcgGain = 0;
  var exitLoad = 0;
  var yearMs = 365.25 * 86400000;
  for (var lotIndex = 0; lotIndex < lots.length; lotIndex++) {
    var lot = lots[lotIndex];
    var value = lot.units * endNav;
    var gain = value - lot.cost;
    if (endDate.getTime() - lot.t >= yearMs) ltcgGain += gain;
    else {
      stcgGain += gain;
      exitLoad += Math.max(0, value) * 0.01;
    }
  }
  return {
    units: units,
    invested: invested,
    corpus: corpus,
    profit: corpus - invested,
    absoluteReturn: (corpus - invested) / invested,
    xirr: calculateXIRR(flows),
    redeem: {
      stcgGain: roundRupee(stcgGain),
      ltcgGain: roundRupee(ltcgGain),
      exitLoad: roundRupee(exitLoad)
    }
  };
}

function threeYearPath(series, saleDate) {
  var path = [];
  for (var back = 35; back >= 0; back--) {
    var end = shiftBack(saleDate, 'month', back);
    var start = shiftBack(end, 'year', 3);
    var cagr = trailingCagr(series, start, end, true);
    path.push(cagr === null ? null : Math.round(cagr * 10000) / 10000);
  }
  return path;
}

/**
 * Point-to-point NAV change from startDate to endDate.
 * Under 1 year (annualized false) this is the absolute change. From 1Y it is the annualized CAGR.
 * This can be present when buildSip returns a null XIRR.
 */
function trailingCagr(series, startDate, endDate, annualized) {
  var startNav = lookupNav(series, startDate);
  var endNav = lookupNav(series, endDate);
  if (startNav === null || endNav === null || startNav <= 0) return null;
  var pointToPoint = (endNav - startNav) / startNav;
  if (!annualized) return pointToPoint;
  var days = (endDate.getTime() - startDate.getTime()) / 86400000;
  if (days <= 0) return null;
  var base = 1 + pointToPoint;
  if (base <= 0) return null;
  return Math.pow(base, 365.25 / days) - 1;
}

function firstSipOnOrAfter(startDate, sipDay) {
  var candidate = sipDateInMonth(startDate.getFullYear(), startDate.getMonth(), sipDay);
  if (candidate.getTime() < startDate.getTime()) {
    candidate = nextSipDate(candidate, sipDay);
  }
  return candidate;
}

function nextSipDate(current, sipDay) {
  return sipDateInMonth(current.getFullYear(), current.getMonth() + 1, sipDay);
}

function sipDateInMonth(year, month, sipDay) {
  var first = new Date(year, month, 1);
  var last = new Date(first.getFullYear(), first.getMonth() + 1, 0).getDate();
  return new Date(first.getFullYear(), first.getMonth(), Math.min(sipDay, last));
}

function shiftBack(dateObj, unit, amount) {
  if (unit === 'day') {
    return new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate() - amount);
  }
  if (unit === 'month') {
    return clampDay(dateObj.getFullYear(), dateObj.getMonth() - amount, dateObj.getDate());
  }
  return clampDay(dateObj.getFullYear() - amount, dateObj.getMonth(), dateObj.getDate());
}

function clampDay(year, month, day) {
  var first = new Date(year, month, 1);
  var last = new Date(first.getFullYear(), first.getMonth() + 1, 0).getDate();
  return new Date(first.getFullYear(), first.getMonth(), Math.min(day, last));
}

function calculateXIRR(cashFlows) {
  if (!cashFlows || cashFlows.length < 2) return null;
  var d0 = cashFlows[0].date.getTime();
  var rate = 0.1;

  for (var iter = 0; iter < 50; iter++) {
    var npv = 0;
    var dnpv = 0;
    var failed = false;
    for (var i = 0; i < cashFlows.length; i++) {
      var cf = cashFlows[i];
      var dt = (cf.date.getTime() - d0) / (365.25 * 86400000);
      var factor = Math.pow(1 + rate, dt);
      if (!isFinite(factor) || factor <= 0) {
        failed = true;
        break;
      }
      npv += cf.amount / factor;
      dnpv -= (dt * cf.amount) / (factor * (1 + rate));
    }
    if (failed || !isFinite(npv) || !isFinite(dnpv)) break;
    if (Math.abs(npv) < 1e-4) return rate;
    if (Math.abs(dnpv) < 1e-8) break;
    var newRate = rate - npv / dnpv;
    if (!isFinite(newRate) || newRate <= -0.99 || newRate > 10) break;
    if (Math.abs(newRate - rate) < 1e-7) return newRate;
    rate = newRate;
  }
  return null;
}

function todayDay() {
  var now = new Date();
  return new Date(now.getFullYear(), now.getMonth(), now.getDate());
}

function toDay(value) {
  if (value === null || value === undefined || value === '') return null;
  if (Object.prototype.toString.call(value) === '[object Date]') {
    if (isNaN(value.getTime())) return null;
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }
  return parseDate(value);
}

function parseDate(str) {
  if (!str) return todayDay();
  var parts = String(str).split(/[-/]/);
  if (parts.length >= 3) {
    var year = parseInt(parts[0], 10);
    var month = parseInt(parts[1], 10);
    var day = parseInt(parts[2], 10);
    if (year < 100) {
      day = year;
      year = parseInt(parts[2], 10);
    }
    if (parts[0].length === 4) {
      return new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
    }
    return new Date(year, month - 1, day);
  }
  var parsed = new Date(str);
  if (!isNaN(parsed.getTime())) {
    return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
  }
  return todayDay();
}

function formatDay(dateObj) {
  var month = dateObj.getMonth() + 1;
  var day = dateObj.getDate();
  return dateObj.getFullYear() + '-' + (month < 10 ? '0' : '') + month + '-' + (day < 10 ? '0' : '') + day;
}

function roundNav(value) {
  if (value === null || value === undefined || !isFinite(value)) return null;
  return Number(value.toFixed(4));
}

function roundRupee(value) {
  if (value === null || value === undefined || !isFinite(value)) return null;
  return Math.round(value * 100) / 100;
}
