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
  ACCOUNTS: 'Account',
  HISTORY:  'Account Monthly History',
  TXNS:     'Transactions'
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
    recentTxns:    recentTxns(txns, 20)
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
  return String(row.Category || '').trim() === 'RealEstate';
}

// Forward-fill Category column
function fillCategory(rows) {
  let last = '';
  rows.forEach(r => {
    if (r.Category && String(r.Category).trim()) last = String(r.Category).trim();
    else r.Category = last;
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
  let base = 0;        // net worth excluding real estate
  let realEstate = 0;  // real estate total (sent separately for the toggle)
  let liquid = 0, invested = 0, gold = 0, debt = 0, disputed = 0;

  accounts.forEach(a => {
    const v = num(a['INR Gross Balance']);
    const include = includeInNetWorth(a);

    // Accumulate real estate separately regardless of the Include flag
    // (the flag is the sheet's "default"; the UI toggle overrides it live)
    if (isRealEstate(a)) {
      realEstate += v;
    } else if (include) {
      base += v;
    }

    // KPI breakdown tiles — still category-driven for display
    switch (String(a.Category || '').trim()) {
      case 'Euro':
      case 'Savings A/C':
        liquid += v; break;
      case 'Stocks':
      case 'IPO Fund':
      case 'Investment':
        invested += v; break;
      case 'Stocks covered in 15':
        break; // skip — already in ID 15 total
      case 'Falcon':
        disputed += v; break;
      case 'Gold':
        gold += v; break;
      case 'Debt':
        debt += v; break;
    }
  });

  // total = base + realEstate (both included by default since sheet has TRUE)
  // The UI toggle will subtract realEstate client-side when switched off
  return {
    total:      Math.round(base + realEstate),
    base:       Math.round(base),       // net worth without real estate
    realEstate: Math.round(realEstate), // real estate portion, for the toggle
    liquid:     Math.round(liquid),
    invested:   Math.round(invested),
    gold:       Math.round(gold),
    debt:       Math.round(debt),
    disputed:   Math.round(disputed)
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
  const labelMap = {
    'Stocks':       'Stocks & IPO',
    'IPO Fund':     'Stocks & IPO',
    'Investment':   'Mutual funds',
    'Euro':         'Euro bank',
    'Savings A/C':  'INR savings',
    'Gold':         'Gold',
    'RealEstate':   'Real estate',
    'Falcon':       'Disputed',
    'Debt':         'Debt (cards)'
  };
  const colorMap = {
    'Stocks & IPO':  PALETTE.green,
    'Euro bank':     PALETTE.orange,
    'INR savings':   PALETTE.greenDeep,
    'Gold':          PALETTE.gold,
    'Mutual funds':  PALETTE.purple,
    'Real estate':   PALETTE.slate,
    'Disputed':      PALETTE.slateLight,
    'Debt (cards)':  PALETTE.red
  };
  const sums = {};
  accounts.forEach(a => {
    if (a.Category === 'Stocks covered in 15') return;
    const label = labelMap[a.Category] || a.Category || 'Other';
    sums[label] = (sums[label] || 0) + num(a['INR Gross Balance']);
  });
  return Object.keys(sums).map(label => ({
    label,
    value: Math.round(sums[label]),
    color: colorMap[label] || PALETTE.taupe
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
    const v = num(a['INR Gross Balance']);
    const name = String(a.Account || '').toLowerCase();
    if (a.Category === 'Debt') debt += v;
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
  const groupMap = {
    'Euro':         'Euro bank accounts',
    'Savings A/C':  'INR savings accounts',
    'Stocks':       'Stocks & IPO',
    'IPO Fund':     'Stocks & IPO',
    'Investment':   'Stocks & IPO',
    'Falcon':       'Disputed (under recovery)',
    'Gold':         'Gold',
    'RealEstate':   'Real estate',
    'Debt':         'Debt (credit cards)'
  };
  const groups = {};
  accounts.forEach(a => {
    if (!a.Account) return;
    if (a.Category === 'Stocks covered in 15') return;
    const v = num(a['INR Gross Balance']);
    if (v === 0) return;
    const grp = groupMap[a.Category] || 'Other';
    if (!groups[grp]) groups[grp] = [];
    groups[grp].push({
      name: String(a.Account),
      bal:  num(a.Balance),
      ccy:  isEur(a.Currency) ? 'EUR' : 'INR',
      inr:  Math.round(v)
    });
  });
  Object.keys(groups).forEach(g =>
    groups[g].sort((a, b) => Math.abs(b.inr) - Math.abs(a.inr))
  );
  const order = ['Euro bank accounts', 'INR savings accounts', 'Stocks & IPO', 'Gold', 'Real estate', 'Disputed (under recovery)', 'Debt (credit cards)', 'Other'];
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
