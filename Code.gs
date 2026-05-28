/**
 * Family Treasury Dashboard — Google Apps Script backend
 *
 * Setup: open your Google Sheet → Extensions → Apps Script.
 * Paste this into Code.gs, paste Index.html into a new HTML file named "Index".
 * Then Deploy → New deployment → Web app.
 *
 * The script reads three sheets from the bound workbook:
 *   - "Account"                  (current balances per account)
 *   - "Account Monthly History"  (monthly snapshots for the trend chart)
 *   - "Transactions"             (for cash flow and recent activity)
 *
 * Column F "Include in Net Worth" in the Account sheet controls which rows
 * count toward the headline net worth figure. Set TRUE or FALSE per row.
 * The toggle in the UI can additionally exclude/include Real Estate on the fly.
 */

const SHEETS = {
  ACCOUNTS:  'Account',
  HISTORY:   'Account Monthly History',
  TXNS:      'Transactions',
  STOCKS_MF: 'Stocks + MF Summary'
};

const EUR_CCYS = ['EUR', 'Eur', 'Euro', 'EURO'];

const PALETTE = {
  green:     '#22c55e',
  greenDeep: '#15803d',
  greenSoft: '#84cc16',
  orange:    '#f97316',
  orangeDeep:'#ea580c',
  red:       '#ef4444',
  gold:      '#eab308',
  purple:    '#8b5cf6',
  cyan:      '#06b6d4',
  slate:     '#64748b',
  slateLight:'#94a3b8',
  dark:      '#1a1d1f'
};

// ============================================================
// ENTRY POINT
// ============================================================
function doGet() {
  const t = HtmlService.createTemplateFromFile('Index');
  t.data = JSON.stringify(buildDashboardData());
  return t.evaluate()
    .setTitle('Family Treasury')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================
// MAIN DATA BUILDER
// ============================================================
function buildDashboardData() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const accounts = readSheet(ss, SHEETS.ACCOUNTS);
  const history  = readSheet(ss, SHEETS.HISTORY);
  const txns     = readSheet(ss, SHEETS.TXNS);

  fillCategory(accounts);
  fillCategory(history);

  return {
    asOf:          new Date().toISOString(),
    eurRate:       latestEurRate(accounts),
    summary:       summary(accounts),
    trend:         trend(history),
    allocation:    allocation(accounts),
    currencySplit: currencySplit(accounts),
    perHolder:     perHolder(accounts),
    cashflow:      cashflow(txns),
    income:        incomeBreakdown(txns),
    expense:       expenseBreakdown(txns),
    accounts:      accountGroups(accounts),
    recentTxns:    recentTxns(txns, 20),
    stocksDetail:  readStocksDetail(ss)
  };
}

// ============================================================
// SHEET READER
// ============================================================
function readSheet(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) { Logger.log('Sheet missing: ' + name); return []; }
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(h => String(h).trim());
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (r.every(v => v === '' || v === null)) continue;
    const obj = {};
    headers.forEach((h, j) => obj[h] = r[j]);
    out.push(obj);
  }
  return out;
}

// ============================================================
// HELPERS
// ============================================================
function isEur(c) { return EUR_CCYS.indexOf(String(c || '').trim()) >= 0; }
function num(v)   { const n = Number(v); return isNaN(n) ? 0 : n; }

function includeInNetWorth(row) {
  // Reads column F "Include in Net Worth" — accepts TRUE (boolean) or 'TRUE' (string)
  const v = row['Include in Net Worth'];
  if (typeof v === 'boolean') return v;
  return String(v).trim().toUpperCase() === 'TRUE';
}

function isRealEstate(row) {
  return String(row.Label || '').trim() === 'Real Estate';
}

// Forward-fill Label and Category columns
function fillCategory(rows) {
  let lastLabel = '', lastCat = '';
  rows.forEach(r => {
    if (r.Label    && String(r.Label).trim())    lastLabel = String(r.Label).trim();
    else r.Label = lastLabel;
    if (r.Category && String(r.Category).trim()) lastCat   = String(r.Category).trim();
    else r.Category = lastCat;
  });
}

function latestEurRate(accounts) {
  const rates = accounts.map(a => num(a['Euro rate'])).filter(r => r > 0);
  return rates.length ? rates[rates.length - 1] : 112;
}

// ============================================================
// HERO SUMMARY
// Now driven by "Include in Net Worth" column (col F).
// Real estate rows are passed through separately so the UI
// toggle can add/remove them client-side without a reload.
// ============================================================
function summary(accounts) {
  let base = 0;
  let realEstate = 0;
  let liquid = 0, stocksMF = 0, gold = 0, disputed = 0, fixedDeposit = 0;

  accounts.forEach(a => {
    const v    = num(a['INR Gross Balance']);
    const lbl  = String(a.Label || '').trim();
    const incl = includeInNetWorth(a);  // col G must be TRUE

    // Skip rows where Label (col B) is blank — not counted anywhere
    if (!lbl) return;

    // Real Estate: accumulate separately for the UI toggle, only if col G = TRUE
    if (lbl === 'Real Estate') {
      if (incl) realEstate += v;
      return;
    }

    // Skip rows excluded from net worth (col G = FALSE)
    if (!incl) return;

    // Net worth base
    base += v;

    // Hero KPI buckets — driven directly by Label
    switch (lbl) {
      case 'Liquid':         liquid       += v; break;
      case 'Stocks & MF':   stocksMF     += v; break;
      case 'Fixed Deposit':  fixedDeposit += v; break;
      case 'Gold':           gold         += v; break;
      case 'Falcon':         disputed     += v; break;
    }
  });

  return {
    total:        Math.round(base + realEstate),
    base:         Math.round(base),
    realEstate:   Math.round(realEstate),
    liquid:       Math.round(liquid),
    stocksMF:     Math.round(stocksMF),
    gold:         Math.round(gold),
    fixedDeposit: Math.round(fixedDeposit),
    disputed:     Math.round(disputed)
  };
}

// ============================================================
// TREND
// ============================================================
function trend(history) {
  const buckets = {};
  history.forEach(r => {
    const d = r.Date;
    if (!(d instanceof Date)) return;
    const key = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    buckets[key] = (buckets[key] || 0) + num(r['INR Gross Balance']);
  });
  return Object.keys(buckets).sort().map(k => ({
    date:  k,
    total: Math.round(buckets[k])
  }));
}

// ============================================================
// ALLOCATION
// ============================================================
function allocation(accounts) {
  const colorMap = {
    'Liquid':         '#e8714a',
    'Stocks & MF':    PALETTE.greenDeep,
    'Fixed Deposit':  '#3b82f6',
    'Gold':           PALETTE.gold,
    'Real Estate':    PALETTE.slate,
    'Falcon':         PALETTE.slateLight,
    'Debt':           PALETTE.red
  };
  const sums = {};
  accounts.forEach(a => {
    const lbl = String(a.Label || '').trim();
    if (!lbl) return;
    const incl = includeInNetWorth(a);
    if (!incl && lbl !== 'Real Estate') return;
    sums[lbl] = (sums[lbl] || 0) + num(a['INR Gross Balance']);
  });
  return Object.keys(sums).map(label => ({
    label,
    value: Math.round(sums[label]),
    color: colorMap[label] || PALETTE.slate
  }));
}

// ============================================================
// CURRENCY SPLIT
// ============================================================
function currencySplit(accounts) {
  let inr = 0, eur = 0;
  accounts.forEach(a => {
    const v = num(a['INR Gross Balance']);
    if (isEur(a.Currency)) eur += v;
    else inr += v;
  });
  return { inr: Math.round(inr), eur: Math.round(eur) };
}

// ============================================================
// PER HOLDER
// ============================================================
function perHolder(accounts) {
  let hiral = 0, divyang = 0, family = 0, debt = 0;
  accounts.forEach(a => {
    const v   = num(a['INR Gross Balance']);
    const lbl = String(a.Label || '').trim();
    const name = String(a.Account || '').toLowerCase();
    if (lbl === 'Debt') debt += v;
    else if (name.indexOf('hiral') >= 0) hiral += v;
    else if (name.indexOf('divyang') >= 0) divyang += v;
    else family += v;
  });
  return {
    hiral:   Math.round(hiral),
    divyang: Math.round(divyang),
    family:  Math.round(family),
    debt:    Math.round(debt)
  };
}

// ============================================================
// CASH FLOW
// ============================================================
function cashflow(txns) {
  const cutoff = new Date();
  cutoff.setMonth(cutoff.getMonth() - 12);
  const buckets = {};
  txns.forEach(t => {
    const d = t.Date;
    if (!(d instanceof Date) || d < cutoff) return;
    if (!isEur(t.Currency)) return;
    const key = d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2);
    if (!buckets[key]) buckets[key] = { inc: 0, exp: 0 };
    buckets[key].inc += num(t.Income);
    buckets[key].exp += num(t.Expense);
  });
  return Object.keys(buckets).sort().map(k => {
    const [y, m] = k.split('-');
    const dt = new Date(Number(y), Number(m) - 1, 1);
    return {
      m:   Utilities.formatDate(dt, Session.getScriptTimeZone(), 'MMM yy'),
      inc: Math.round(buckets[k].inc),
      exp: Math.round(buckets[k].exp)
    };
  });
}

// ============================================================
// INCOME / EXPENSE BREAKDOWN
// ============================================================
function incomeBreakdown(txns) {
  return breakdown_(txns, 'income', 'Income',
    [PALETTE.green, '#16a34a', PALETTE.greenDeep, PALETTE.greenSoft, PALETTE.gold, PALETTE.cyan, PALETTE.slate]);
}

function expenseBreakdown(txns) {
  return breakdown_(txns, 'expense', 'Expense',
    ['#f97316', '#fb923c', '#ea580c', '#c2410c', '#9a3412', '#7c2d12', '#ef4444']);
}

function breakdown_(txns, typeWanted, amountCol, palette) {
  const cutoff = new Date();
  cutoff.setMonth(cutoff.getMonth() - 12);
  const sums = {};
  txns.forEach(t => {
    const d = t.Date;
    if (!(d instanceof Date) || d < cutoff) return;
    if (String(t.Type || '').toLowerCase() !== typeWanted) return;
    if (!isEur(t.Currency)) return;
    const cat = String(t['Txn_Category'] || 'Other').trim() || 'Other';
    sums[cat] = (sums[cat] || 0) + num(t[amountCol]);
  });
  return Object.keys(sums)
    .map(cat => ({ cat, amt: Math.round(sums[cat]) }))
    .filter(x => x.amt > 0)
    .sort((a, b) => b.amt - a.amt)
    .map((x, i) => Object.assign(x, { color: palette[i % palette.length] }));
}

// ============================================================
// ACCOUNT GROUPS
// ============================================================
function accountGroups(accounts) {
  const colorMap = {
    'Liquid':         '#e8714a',
    'Stocks & MF':    '#15803d',
    'Fixed Deposit':  '#3b82f6',
    'Gold':           '#eab308',
    'Real Estate':    '#c4a882',
    'Falcon':         '#94a3b8',
    'Debt':           '#ef4444'
  };
  const groups = {};
  accounts.forEach(a => {
    const lbl = String(a.Label || '').trim();
    const acct = String(a.Account || a.Category || '').trim();
    if (!lbl || !acct) return;
    const v = num(a['INR Gross Balance']);
    if (v === 0) return;
    if (!groups[lbl]) groups[lbl] = [];
    groups[lbl].push({
      name: acct,
      cat:  String(a.Category || '').trim(),
      bal:  num(a.Balance),
      ccy:  isEur(a.Currency) ? 'EUR' : 'INR',
      inr:  Math.round(v),
      incl: includeInNetWorth(a)
    });
  });
  Object.keys(groups).forEach(g =>
    groups[g].sort((a, b) => Math.abs(b.inr) - Math.abs(a.inr))
  );
  const order = ['Liquid', 'Fixed Deposit', 'Stocks & MF', 'Gold', 'Real Estate', 'Falcon', 'Debt', 'Other'];
  const out = {};
  order.forEach(k => { if (groups[k]) out[k] = groups[k]; });
  return out;
}

// ============================================================
// RECENT TRANSACTIONS
// ============================================================
function recentTxns(txns, limit) {
  return txns
    .filter(t => t.Date instanceof Date)
    .sort((a, b) => b.Date - a.Date)
    .slice(0, limit)
    .map(t => ({
      date: Utilities.formatDate(t.Date, Session.getScriptTimeZone(), 'MMM dd'),
      type: String(t.Type || ''),
      acct: String(t.Account || ''),
      amt:  num(t.Amount),
      ccy:  isEur(t.Currency) ? 'EUR' : 'INR',
      note: String(t.Note || t['Txn_Category'] || '')
    }));
}

// ============================================================
// AI ADVISOR — builds prompt from live sheet data + calls Gemini
// Called from Index.html via google.script.run
// ============================================================
function runGeminiAdvisor() {
  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) return JSON.stringify({ error: 'GEMINI_API_KEY not set in Script Properties' });

    const ss       = SpreadsheetApp.getActiveSpreadsheet();
    const accounts = readSheet(ss, SHEETS.ACCOUNTS);
    const txns     = readSheet(ss, SHEETS.TXNS);
    fillCategory(accounts);

    const s   = summary(accounts);
    const nw  = s.total || 1;
    const cf  = cashflow(txns);
    const ccy = currencySplit(accounts);
    const sd  = readStocksDetail(ss);
    const ft  = sd ? (sd.familyTotal || {}) : {};

    var avgInc = cf.length ? cf.reduce(function(a,d){return a+d.inc;},0)/cf.length : 0;
    var avgExp = cf.length ? cf.reduce(function(a,d){return a+d.exp;},0)/cf.length : 0;
    var savingsRate = avgInc ? ((avgInc-avgExp)/avgInc*100).toFixed(1) : 0;

    function pct(v){ return (v/nw*100).toFixed(1); }
    function inr(v){
      var abs = Math.abs(v), sign = v < 0 ? '-' : '';
      if(abs>=10000000) return sign+'Rs '+(abs/10000000).toFixed(2)+' Cr';
      if(abs>=100000)   return sign+'Rs '+(abs/100000).toFixed(2)+' L';
      return sign+'Rs '+Math.round(abs).toLocaleString();
    }

    var topSectors = sd ? (sd.sectors||[]).slice(0,3).map(function(x){return x.label+' ('+x.pct.toFixed(1)+'%)';}).join(', ') : 'N/A';
    var top2Pct    = sd ? (sd.sectors||[]).slice(0,2).reduce(function(a,x){return a+x.pct;},0).toFixed(1) : 0;
    var smallCap   = sd ? (sd.capSplit||[]).filter(function(c){return c.label==='Small Cap';})[0] : null;
    var largeCap   = sd ? (sd.capSplit||[]).filter(function(c){return c.label==='Large Cap';})[0] : null;
    var accts      = sd ? (sd.accounts||[]) : [];
    var worstAcct  = accts.length ? accts.slice().sort(function(a,b){return a.plPct-b.plPct;})[0] : null;
    var bestAcct   = accts.length ? accts.slice().sort(function(a,b){return b.plPct-a.plPct;})[0] : null;

    var prompt = 'You are a sharp, experienced CA (Chartered Accountant) advising an Indian family. Be direct and specific with numbers — like a trusted advisor who respects the client. Identify problems clearly but frame actions constructively. No flattery, no harsh language, no drama. Just clear diagnosis and actionable steps.\n\n';
    prompt += 'WEALTH SNAPSHOT:\n';
    prompt += 'Net Worth: ' + inr(nw) + '\n';
    prompt += 'Liquid (savings+cash+IPO fund): ' + inr(s.liquid) + ' = ' + pct(s.liquid) + '% of NW\n';
    prompt += 'Stocks & MF: ' + inr(s.stocksMF) + ' = ' + pct(s.stocksMF) + '% of NW\n';
    prompt += 'Real Estate: ' + inr(s.realEstate) + ' = ' + pct(s.realEstate) + '% of NW\n';
    prompt += 'Fixed Deposit: ' + inr(s.fixedDeposit||0) + ' = ' + pct(s.fixedDeposit||0) + '% of NW\n';
    prompt += 'Disputed (excluded from NW): ' + inr(s.disputed||0) + '\n\n';
    prompt += 'ZERODHA PORTFOLIO:\n';
    prompt += 'Invested: ' + inr(ft.invested||0) + ' | Value: ' + inr(ft.value||0) + ' | P&L: ' + inr(ft.pl||0) + ' (' + ((ft.plPct||0)).toFixed(2) + '%)\n';
    prompt += 'Best account: ' + (bestAcct ? bestAcct.name + ' +' + bestAcct.plPct.toFixed(2) + '%' : 'N/A') + '\n';
    prompt += 'Worst account: ' + (worstAcct ? worstAcct.name + ' ' + worstAcct.plPct.toFixed(2) + '%' : 'N/A') + '\n';
    prompt += 'Cap split: Large ' + (largeCap?largeCap.pct.toFixed(1):0) + '% | Small ' + (smallCap?smallCap.pct.toFixed(1):0) + '%\n';
    prompt += 'Top sectors: ' + topSectors + ' | Top 2 concentration: ' + top2Pct + '%\n\n';
    prompt += 'CASH FLOW (EUR avg/month, last 12 months):\n';
    prompt += 'Income: EUR ' + Math.round(avgInc) + ' | Expense: EUR ' + Math.round(avgExp) + ' | Savings rate: ' + savingsRate + '%\n\n';
    prompt += 'CURRENCY: ' + pct(ccy.inr) + '% INR | ' + pct(ccy.eur) + '% EUR\n\n';
    prompt += 'Generate exactly 6 insight cards. Keep each field under 20 words. Return ONLY valid JSON with no markdown, no backticks, no extra text:\n';
    prompt += 'Return ONLY valid JSON with no markdown, no backticks, no extra text:\n';
    prompt += '{"score":65,"verdict":"one blunt line","cards":[{"category":"LIQUIDITY","title":"short title","severity":"critical","observation":"blunt 1-2 sentences with numbers","impact":"what this costs them","action":"specific action with amount and timeline"},{"category":"ALLOCATION","title":"...","severity":"warning","observation":"...","impact":"...","action":"..."},{"category":"PORTFOLIO","title":"...","severity":"warning","observation":"...","impact":"...","action":"..."},{"category":"SECTORS","title":"...","severity":"advisory","observation":"...","impact":"...","action":"..."},{"category":"CASH FLOW","title":"...","severity":"healthy","observation":"...","impact":"...","action":"..."},{"category":"CURRENCY","title":"...","severity":"advisory","observation":"...","impact":"...","action":"..."}]}';

    Logger.log('Prompt length: ' + prompt.length);

    var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey;
    var payload = { contents: [{ parts: [{ text: prompt }] }], generationConfig: { temperature: 0.3, maxOutputTokens: 4000 } };
    var options = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true };

    var response = UrlFetchApp.fetch(url, options);
    Logger.log('HTTP status: ' + response.getResponseCode());
    var raw  = response.getContentText();
    Logger.log('Raw (first 600): ' + raw.substring(0, 600));
    var data = JSON.parse(raw);
    if (data.error) return JSON.stringify({ error: data.error.message });

    var text = '';
    if (data.candidates && data.candidates[0] && data.candidates[0].content &&
        data.candidates[0].content.parts && data.candidates[0].content.parts[0]) {
      text = data.candidates[0].content.parts[0].text || '';
    }
    Logger.log('Extracted text (first 400): ' + text.substring(0, 400));
    text = text.replace(/^```json\s*/,'').replace(/^```\s*/,'').replace(/\s*```$/,'').trim();
    return text;

  } catch(e) {
    Logger.log('runGeminiAdvisor error: ' + e.toString());
    return JSON.stringify({ error: e.toString() });
  }
}

function testGeminiAdvisor() {
  var result = runGeminiAdvisor();
  Logger.log('Result (first 500): ' + result.substring(0, 500));
  try {
    var parsed = JSON.parse(result);
    Logger.log('Parse SUCCESS — score: ' + parsed.score + ', cards: ' + parsed.cards.length);
  } catch(e) {
    Logger.log('Parse FAILED: ' + e.toString());
  }
}
// Fixed row map (1-indexed, as in the sheet):
//   Section A  : D7  = Zerodha total, D8 = Others total
//   Accounts   : header row 19, data rows 20-24, total row 25  cols B-F
//   Asset Class: header row 28, data rows 29-32               cols B-D
//   Cap Split  : header row 28, data rows 29-31               cols E-G
//   Sectors    : header row 35, data rows 36-44 (9 rows)      cols B-D + E-G (merged)
//   IPO list   : header row 64, data rows 65-104              cols B-G
// ============================================================
function readStocksDetail(ss) {
  try {
    const sh = ss.getSheetByName(SHEETS.STOCKS_MF);
    if (!sh) { Logger.log('Sheet missing: ' + SHEETS.STOCKS_MF); return null; }

    // Helper: get single cell value, 1-indexed row/col (A=1)
    const cellVal = function(r, c) { return sh.getRange(r, c).getValue(); };
    // Helper: get a single row as array, 1-indexed
    const rowVals = function(r, startC, cnt) { return sh.getRange(r, startC, 1, cnt).getValues()[0]; };
    // Helper: get a 2-D block, 1-indexed
    const blk = function(r, c, nr, nc) { return sh.getRange(r, c, nr, nc).getValues(); };

    // ── L2 totals from Section A ─────────────────────────────
    const zerodhaTotalVal = num(cellVal(7, 4));  // D7
    const othersTotalVal  = num(cellVal(8, 4));  // D8

    // ── Account breakdown rows 20-24, total row 25, cols B-F ─
    const acctData   = blk(20, 2, 5, 5);
    const totalData  = rowVals(25, 2, 5);
    const accounts_  = acctData
      .filter(function(r) { return r[0] !== '' && r[0] !== null; })
      .map(function(r) {
        return { name: String(r[0]), invested: num(r[1]), value: num(r[2]), pl: num(r[3]), plPct: num(r[4]) };
      });
    const familyTotal = {
      invested: num(totalData[1]),
      value:    num(totalData[2]),
      pl:       num(totalData[3]),
      plPct:    num(totalData[4])
    };

    // ── Asset Class split rows 29-32, cols B-D ───────────────
    const assetClass = blk(29, 2, 4, 3)
      .filter(function(r) { return r[0] !== '' && r[0] !== null; })
      .map(function(r) { return { label: String(r[0]), value: num(r[1]), pct: num(r[2]) }; });

    // ── Cap split rows 29-31, cols E-G ───────────────────────
    const capSplit = blk(29, 5, 3, 3)
      .filter(function(r) { return r[0] !== '' && r[0] !== null; })
      .map(function(r) { return { label: String(r[0]), value: num(r[1]), pct: num(r[2]) }; });

    // ── Sectors rows 36-44, left cols B-D + right cols E-G ───
    const secLeft  = blk(36, 2, 9, 3);
    const secRight = blk(36, 5, 9, 3);
    const sectors  = [];
    for (var i = 0; i < 9; i++) {
      if (secLeft[i][0]  !== '' && secLeft[i][0]  !== null)
        sectors.push({ label: String(secLeft[i][0]),  value: num(secLeft[i][1]),  pct: num(secLeft[i][2])  });
      if (secRight[i][0] !== '' && secRight[i][0] !== null)
        sectors.push({ label: String(secRight[i][0]), value: num(secRight[i][1]), pct: num(secRight[i][2]) });
    }
    sectors.sort(function(a, b) { return b.value - a.value; });

    // ── IPO list rows 65-104, cols B-G, filter Zerodha=FALSE ─
    const ipoOthers = blk(65, 2, 40, 6)
      .filter(function(r) { return r[0] !== '' && r[0] !== null; })
      .filter(function(r) {
        var z = r[5]; // col G = Zerodha?
        if (typeof z === 'boolean') return !z;
        return String(z).trim().toUpperCase() !== 'TRUE';
      })
      .map(function(r) {
        return { account: String(r[0]), stock: String(r[1]), invested: num(r[2]), value: num(r[3]), pl: num(r[4]) };
      });

    return { zerodhaTotalVal, othersTotalVal, accounts: accounts_, familyTotal, assetClass, capSplit, sectors, ipoOthers };

  } catch(e) {
    Logger.log('readStocksDetail error: ' + e.toString());
    return null;  // returns null safely — UI will skip the section
  }
}
