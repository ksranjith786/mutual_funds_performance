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
      tenureCagr: tenureCagr
    });
  }

  return {
    funds: fundList,
    categories: Object.keys(categoriesSet).sort(),
    amcs: Object.keys(amcSet).sort(),
    tenures: TENURE_CONFIG
  };
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
  var empty = { units: 0, invested: 0, corpus: null, profit: null, absoluteReturn: null, xirr: null };
  if (!series || !startDate || !endDate || endDate < startDate) return empty;

  var endNav = lookupNav(series, endDate);
  if (endNav === null) return empty;

  var units = 0;
  var invested = 0;
  var flows = [];
  var cur = firstSipOnOrAfter(startDate, sipDay);

  while (cur && cur.getTime() <= endDate.getTime()) {
    var nav = lookupNav(series, cur);
    if (nav !== null && nav > 0) {
      units += sipAmount / nav;
      invested += sipAmount;
      flows.push({ date: new Date(cur.getTime()), amount: -sipAmount });
    }
    cur = nextSipDate(cur, sipDay);
  }

  if (invested <= 0 || units <= 0) return empty;

  var corpus = units * endNav;
  flows.push({ date: new Date(endDate.getTime()), amount: corpus });
  return {
    units: units,
    invested: invested,
    corpus: corpus,
    profit: corpus - invested,
    absoluteReturn: (corpus - invested) / invested,
    xirr: calculateXIRR(flows)
  };
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
