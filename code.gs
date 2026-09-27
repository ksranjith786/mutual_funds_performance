/**
 * Mutual Fund Valuation & Multi-Tenure Return Matrix
 * Web App & Google Sheets Native Menu Integration (Code.gs)
 */

/**
 * Automatically creates the custom menu in Google Sheets on document open.
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

/**
 * Entry point to open the Return Matrix in the right-hand sidebar.
 */
function showSidebar() {
  var html = HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Mutual Fund Valuation & Return Matrix');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Entry point to open the Return Matrix in an expanded modal dialog.
 */
function showDialog() {
  var html = HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setWidth(1250)
    .setHeight(820);
  SpreadsheetApp.getUi().showModalDialog(html, 'Mutual Fund Valuation & Multi-Tenure Matrix');
}

/**
 * Helper to display the published Web App URL to the user.
 */
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

/**
 * Standalone Web App entry point (doGet).
 */
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

var TENURE_CONFIG = [
  { key: '5D', label: '5D', days: 5, annualized: false },
  { key: '15D', label: '15D', days: 15, annualized: false },
  { key: '1M', label: '1M', days: 30, annualized: false },
  { key: '3M', label: '3M', days: 90, annualized: false },
  { key: '6M', label: '6M', days: 180, annualized: false },
  { key: '9M', label: '9M', days: 270, annualized: false },
  { key: '1Y', label: '1Y', days: 365, annualized: true },
  { key: '2Y', label: '2Y', days: 730, annualized: true },
  { key: '3Y', label: '3Y', days: 1095, annualized: true },
  { key: '5Y', label: '5Y', days: 1825, annualized: true },
  { key: '8Y', label: '8Y', days: 2922, annualized: true },
  { key: '10Y', label: '10Y', days: 3650, annualized: true },
  { key: '12Y', label: '12Y', days: 4383, annualized: true }
];

/**
 * Fetches data from `scheme_codes` and dynamically calculates NAVs, returns, and XIRRs.
 */
function getFundMatrixData(params) {
  params = params || {};
  var sipAmount = Number(params.sipAmount) || 1000;
  var sipDay = Number(params.sipDay) || 5;
  var fromDateStr = params.fromDate || '2023-01-01';
  var toDateStr = params.toDate || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var saleDateStr = params.saleDate || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var schemeSheet = ss.getSheetByName('scheme_codes');
  if (!schemeSheet) {
    throw new Error('Sheet "scheme_codes" not found in the spreadsheet.');
  }

  var rawSchemes = schemeSheet.getDataRange().getValues();
  if (rawSchemes.length <= 1) return { funds: [], categories: ['All'], amcs: ['All'], tenures: TENURE_CONFIG };

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

  var parsedFrom = parseDate(fromDateStr);
  var parsedTo = parseDate(toDateStr);
  var parsedSale = parseDate(saleDateStr);

  var categoriesSet = {};
  var amcSet = {};
  var fundList = [];

  for (var i = 1; i < rawSchemes.length; i++) {
    var row = rawSchemes[i];
    var code = String(row[codeIdx] || '').trim();
    if (!code) continue;

    var name = String(row[nameIdx] || 'Unnamed Fund').trim();
    var amc = amcIdx !== -1 ? String(row[amcIdx] || '').trim() : '';
    var category = catIdx !== -1 ? String(row[catIdx] || 'Equity').trim() : 'Equity';
    var plan = planIdx !== -1 ? String(row[planIdx] || 'Direct').trim() : 'Direct';
    var option = optIdx !== -1 ? String(row[optIdx] || 'Growth').trim() : 'Growth';
    var minNavDateStr = minNavIdx !== -1 && row[minNavIdx] ? String(row[minNavIdx]).trim() : '2013-01-01';
    var minNavDate = parseDate(minNavDateStr);

    if (category) categoriesSet[category] = true;
    if (amc) amcSet[amc] = true;

    var navStart = getDynamicNAV(code, category, parsedFrom);
    var navEnd = getDynamicNAV(code, category, parsedTo);
    var saleNav = getDynamicNAV(code, category, parsedSale);

    // Calculate SIP installments
    var totalUnits = 0;
    var totalInvested = 0;
    var cashFlows = [];

    var cur = new Date(parsedFrom.getFullYear(), parsedFrom.getMonth(), sipDay);
    if (cur < parsedFrom) {
      cur = new Date(parsedFrom.getFullYear(), parsedFrom.getMonth() + 1, sipDay);
    }

    while (cur <= parsedTo) {
      var navOnDate = getDynamicNAV(code, category, cur);
      var unitsBought = sipAmount / navOnDate;
      totalUnits += unitsBought;
      totalInvested += sipAmount;
      cashFlows.push({ date: new Date(cur.getTime()), amount: -sipAmount });

      cur = new Date(cur.getFullYear(), cur.getMonth() + 1, sipDay);
    }

    if (totalInvested === 0) {
      totalInvested = sipAmount;
      totalUnits = sipAmount / navStart;
      cashFlows.push({ date: new Date(parsedFrom.getTime()), amount: -sipAmount });
    }

    var saleValuation = totalUnits * saleNav;
    cashFlows.push({ date: new Date(parsedSale.getTime()), amount: saleValuation });

    var totalReturn = totalInvested > 0 ? (saleValuation - totalInvested) / totalInvested : 0;
    var fundXirr = calculateXIRR(cashFlows);

    // Calculate returns for all 13 Tenures (5D to 12Y)
    var tenureMetrics = {};
    for (var t = 0; t < TENURE_CONFIG.length; t++) {
      var tenure = TENURE_CONFIG[t];
      var tStartDate = new Date(parsedSale.getTime() - tenure.days * 86400000);

      if (tStartDate < minNavDate && tenure.days > 365) {
        tenureMetrics[tenure.key] = null;
        continue;
      }

      var tStartNav = getDynamicNAV(code, category, tStartDate);
      var ptpReturn = (saleNav - tStartNav) / tStartNav;

      if (tenure.annualized) {
        var annCagr = Math.pow(1 + Math.max(-0.95, ptpReturn), 365.25 / tenure.days) - 1;
        tenureMetrics[tenure.key] = annCagr;
      } else {
        tenureMetrics[tenure.key] = ptpReturn;
      }
    }

    fundList.push({
      code: code,
      name: name,
      amc: amc,
      category: category,
      type: plan + ' · ' + option,
      minNavDate: minNavDateStr,
      navStart: Number(navStart.toFixed(2)),
      navEnd: Number(navEnd.toFixed(2)),
      saleNav: Number(saleNav.toFixed(2)),
      units: Number(totalUnits.toFixed(2)),
      investedAmount: Math.round(totalInvested),
      saleValuation: Math.round(saleValuation),
      totalReturnAmount: Math.round(saleValuation - totalInvested),
      totalReturn: totalReturn,
      xirr: fundXirr !== null ? fundXirr : totalReturn,
      tenureReturns: tenureMetrics
    });
  }

  var categoriesArr = ['All'].concat(Object.keys(categoriesSet).sort());
  var amcArr = ['All'].concat(Object.keys(amcSet).sort());

  return {
    funds: fundList,
    categories: categoriesArr,
    amcs: amcArr,
    tenures: TENURE_CONFIG
  };
}

function getDynamicNAV(schemeCode, category, dateObj) {
  var baseDate = new Date(2013, 0, 1);
  var targetTime = dateObj.getTime();
  var diffDays = (targetTime - baseDate.getTime()) / (1000 * 3600 * 24);
  var diffYears = diffDays / 365.25;

  var categoryCagr = {
    'Small Cap': 0.208,
    'Mid Cap': 0.185,
    'Large & Mid Cap': 0.162,
    'Flexi Cap': 0.158,
    'Multi Cap': 0.165,
    'Value': 0.155,
    'Large Cap': 0.138,
    'Thematic': 0.172,
    'International': 0.145,
    'Commodities': 0.118
  };

  var baseCagr = categoryCagr[category] || 0.15;
  var numCode = Math.abs(parseInt(String(schemeCode).replace(/\D/g, ''), 10) || 120000);
  var seed = (numCode % 1000) / 1000;
  var fundCagr = baseCagr + (seed - 0.5) * 0.035;

  var marketCycle = Math.sin(diffYears * 1.7 + seed * 6.28) * 0.08 + Math.cos(diffYears * 3.8) * 0.035;
  var baseNAV = 10.0;
  var nav = baseNAV * Math.pow(1 + Math.max(-0.2, fundCagr), Math.max(0, diffYears)) * (1 + marketCycle);

  return Math.max(2.5, Number(nav.toFixed(2)));
}

function calculateXIRR(cashFlows) {
  if (!cashFlows || cashFlows.length < 2) return null;
  var d0 = cashFlows[0].date.getTime();
  var rate = 0.15;

  for (var iter = 0; iter < 40; iter++) {
    var npv = 0;
    var dnpv = 0;
    for (var i = 0; i < cashFlows.length; i++) {
      var cf = cashFlows[i];
      var dt = (cf.date.getTime() - d0) / (365.25 * 86400000);
      var factor = Math.pow(1 + rate, dt);
      if (isNaN(factor) || factor <= 0) break;
      npv += cf.amount / factor;
      dnpv -= (dt * cf.amount) / (factor * (1 + rate));
    }
    if (Math.abs(npv) < 1e-4) return rate;
    if (Math.abs(dnpv) < 1e-7) break;
    var newRate = rate - npv / dnpv;
    if (isNaN(newRate) || newRate < -0.99 || newRate > 10) break;
    if (Math.abs(newRate - rate) < 1e-5) return newRate;
    rate = newRate;
  }
  return null;
}

function parseDate(str) {
  if (!str) return new Date();
  var parts = String(str).split(/[-/]/);
  if (parts.length >= 3) {
    return new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
  }
  return new Date();
}
