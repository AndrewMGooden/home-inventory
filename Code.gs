/**
 * Gooden's Home Inventory — the server that talks to your Google Sheet.
 *
 * The app page (index.html) is hosted on GitHub Pages and calls this script's
 * web app URL. Every request must carry a Google sign-in token for the Sheet's
 * owner or someone listed under "household" in the Settings tab.
 *
 * You can edit the tabs by hand, reorder columns, or add your own columns —
 * the app finds columns by their header name.
 */

// Where the app page lives, and the Google sign-in client it uses (see README.md).
const APP_URL = 'https://andrewmgooden.github.io/home-inventory/';
const GOOGLE_CLIENT_ID = '820879358248-q71o9tp8cgov7cqr9jm7r0qbmg7sa13c.apps.googleusercontent.com';

const SCHEMA = {
  items: {
    sheet: 'Inventory',
    cols: [
      ['photo', 'Photo'], ['name', 'Name'], ['category', 'Category'], ['location', 'Location'],
      ['qty', 'Quantity'], ['min', 'Restock At'], ['unit', 'Unit'], ['price', 'Price'],
      ['store', 'Store'], ['notes', 'Notes'], ['updated', 'Updated'], ['photoId', 'Photo ID'],
      ['id', 'ID'],
    ],
  },
  shopping: {
    sheet: 'Shopping List',
    cols: [
      ['name', 'Item'], ['qty', 'Qty'], ['category', 'Category'], ['added', 'Added'],
      ['itemId', 'Item ID'], ['id', 'ID'],
    ],
  },
  bills: {
    sheet: 'Bills',
    cols: [
      ['name', 'Bill'], ['category', 'Category'], ['amount', 'Usual Amount'], ['repeat', 'Repeats'],
      ['due', 'Next Due'], ['notes', 'Notes'], ['id', 'ID'],
    ],
  },
  purchases: {
    sheet: 'Purchases',
    cols: [
      ['date', 'Date'], ['name', 'Item'], ['category', 'Category'], ['qty', 'Qty'],
      ['price', 'Price Each'], ['total', 'Total'], ['store', 'Store'], ['itemId', 'Item ID'],
      ['id', 'ID'], ['billId', 'Bill ID'],
    ],
  },
  settings: {
    sheet: 'Settings',
    cols: [['key', 'Setting'], ['value', 'Value']],
  },
};

const DEFAULT_SETTINGS = { budget: 0, currency: 'JMD', household: '' };
const SETUP_VERSION = '4';
const BILL_REPEATS = ['Monthly', 'Every 2 months', 'Every 3 months', 'Every 4 months', 'Every 6 months', 'Yearly', 'As needed'];
const PHOTO_FOLDER = 'Home Inventory Photos';

// Per-request caches so we only look up each sheet and its headers once.
const cache_ = { sheets: {}, cols: {}, width: {}, tz: null };


/* ───────────────────────── Web API + menu ───────────────────────── */

// Everything the app page may call. Each call needs a valid sign-in token (see checkUser_).
const API = {
  getData, saveItem, setQty, deleteItem, addToShopping, removeFromShopping,
  buy, deletePurchase, saveSettings, saveBill, deleteBill,
};

/**
 * The app sends POST {token, fn, args} as plain text (so browsers skip the CORS
 * pre-check) and gets back {ok, data} or {ok: false, error, auth}.
 */
function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    checkUser_(req.token);
    if (!Object.prototype.hasOwnProperty.call(API, req.fn)) throw new Error('Unknown action: ' + req.fn);
    out = { ok: true, data: API[req.fn].apply(null, req.args || []) };
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err), auth: !!(err && err.auth) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/** Opening the script's own link just points people to the real app. */
function doGet() {
  const html = '<p style="font-family:Arial,sans-serif">This is the Gooden\'s Home Inventory server. ' +
    'The app is at <a href="' + APP_URL + '" target="_top">' + APP_URL + '</a></p>';
  return HtmlService.createHtmlOutput(html).setTitle("Gooden's Home Inventory server");
}

/** Checks a Google sign-in (ID) token and that its email is allowed. Returns the email. */
function checkUser_(token) {
  if (!GOOGLE_CLIENT_ID) throw authError_('Google sign-in is not set up yet (GOOGLE_CLIENT_ID is empty).');
  if (!token) throw authError_('Please sign in.');
  const cache = CacheService.getScriptCache();
  const key = 'tok:' + Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(token)));
  let email = cache.get(key);
  if (!email) {
    const res = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(token),
      { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) throw authError_('Your sign-in has expired. Please sign in again.');
    const info = JSON.parse(res.getContentText());
    const issuerOk = info.iss === 'accounts.google.com' || info.iss === 'https://accounts.google.com';
    if (info.aud !== GOOGLE_CLIENT_ID || !issuerOk || String(info.email_verified) !== 'true') {
      throw authError_('That sign-in was not accepted.');
    }
    email = String(info.email).toLowerCase();
    // Remember a checked token until it expires (at most 6 hours) to keep the app fast.
    const ttl = Math.min(21600, Number(info.exp) - Math.floor(Date.now() / 1000));
    if (ttl > 60) cache.put(key, email, ttl);
  }
  if (allowedEmails_().indexOf(email) === -1) {
    throw authError_(email + ' is not allowed to use this app. The owner can add it under "household" in the Settings tab.');
  }
  return email;
}

/** The Sheet's owner plus any emails listed (comma-separated) under "household" in Settings. */
function allowedEmails_() {
  const row = readAll_('settings').find(r => r.key === 'household');
  const extra = String(row ? row.value : '').split(/[\s,;]+/).filter(Boolean);
  return [Session.getEffectiveUser().getEmail()].concat(extra).map(e => e.toLowerCase());
}

function authError_(message) {
  const err = new Error(message);
  err.auth = true;
  return err;
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏠 Inventory')
    .addItem('Open the app', 'showAppLink')
    .addItem('Rebuild dashboard & colors', 'setup')
    .addToUi();
}

function showAppLink() {
  const style = 'font-family:Arial,sans-serif;font-size:14px';
  const html = '<p style="' + style + '">Your app link:</p>' +
    '<p style="' + style + ';word-break:break-all"><a href="' + APP_URL + '" target="_blank">' + APP_URL + '</a></p>' +
    '<p style="' + style + ';color:#666">On Android: open it in Chrome, tap ⋮ → Install app (or Add to Home screen).</p>';
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(440).setHeight(200), "Gooden's Home Inventory");
}


/* ───────────────────────── Setup and colors ───────────────────────── */

// [header color, stripe color] for each tab.
const THEME = {
  items: ['#2e6a4c', '#eef6f1'],
  shopping: ['#d9822b', '#fdf1e5'],
  bills: ['#7b4fa0', '#f4eef9'],
  purchases: ['#2f6fb0', '#eaf2fb'],
  settings: ['#6b7280', '#f2f3f5'],
};
const FONT = 'Nunito';
const GREY = '#9aa0a6';

/** Creates the tabs, colors everything and rebuilds the Dashboard. Safe to run again anytime. */
function setup() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(SCHEMA).forEach(kind => sheet_(kind));
  const blank = ss.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0 && blank.getLastColumn() === 0) ss.deleteSheet(blank);

  const have = readAll_('settings').map(r => r.key);
  Object.keys(DEFAULT_SETTINGS).forEach(key => {
    if (have.indexOf(key) === -1) write_('settings', { key: key, value: DEFAULT_SETTINGS[key] });
  });

  const money = moneyFormat_();
  Object.keys(SCHEMA).forEach(styleSheet_);

  const inv = sheet_('items'), c = cache_.cols.items, invRows = bodyRows_(inv);
  if (Math.max(c.photo, c.name) <= 3) inv.setFrozenColumns(Math.max(c.photo, c.name));
  inv.setColumnWidth(c.photo, 64);
  inv.setColumnWidth(c.name, 170);
  inv.setColumnWidth(c.notes, 220);
  inv.getRange(2, c.qty, invRows, 1).setHorizontalAlignment('center').setFontWeight('bold');
  inv.getRange(2, c.min, invRows, 1).setHorizontalAlignment('center');
  inv.getRange(2, c.price, invRows, 1).setNumberFormat(money);
  inv.getRange(2, c.updated, invRows, 1).setNumberFormat('mmm d, h:mm am/pm').setFontColor(GREY);
  [c.photoId, c.id].forEach(col => inv.getRange(2, col, invRows, 1).setFontColor(GREY));
  inv.getRange(1, c.min).setNote('When Quantity drops to this number or lower, the row turns red and the item goes on the shopping list.');
  inv.getRange(1, c.photo).setNote('Photos you add in the app show up here.');
  const qty = colA1_(c.qty), min = colA1_(c.min);
  inv.setConditionalFormatRules([
    // Completely out: bold red number (listed first so it wins on the Quantity cell).
    formulaRule_('=AND($' + qty + '2<>"", $' + qty + '2=0)', inv.getRange(2, c.qty, invRows, 1), '#fde2dc', '#b42318'),
    formulaRule_('=AND($' + min + '2<>"", $' + qty + '2<=$' + min + '2)', inv.getRange(2, 1, invRows, cache_.width.items), '#fde2dc'),
  ]);

  const pur = sheet_('purchases'), p = cache_.cols.purchases, purRows = bodyRows_(pur);
  pur.getRange(2, p.date, purRows, 1).setNumberFormat('ddd, mmm d, yyyy');
  pur.getRange(2, p.price, purRows, 1).setNumberFormat(money);
  pur.getRange(2, p.total, purRows, 1).setNumberFormat(money).setFontWeight('bold');
  pur.setColumnWidth(p.date, 130);
  pur.setColumnWidth(p.name, 180);
  [p.itemId, p.id, p.billId].forEach(col => pur.getRange(2, col, purRows, 1).setFontColor(GREY));
  pur.getRange(1, p.billId).setNote('Rows with a Bill ID are bill payments. They are shaded purple.');
  pur.setConditionalFormatRules([
    formulaRule_('=$' + colA1_(p.billId) + '2<>""', pur.getRange(2, 1, purRows, cache_.width.purchases), THEME.bills[1]),
  ]);

  const bills = sheet_('bills'), b = cache_.cols.bills, billRows = bodyRows_(bills);
  bills.getRange(2, b.amount, billRows, 1).setNumberFormat(money);
  bills.getRange(2, b.due, billRows, 1).setNumberFormat('mmm d, yyyy');
  bills.getRange(2, b.repeat, billRows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(BILL_REPEATS, true).build());
  bills.getRange(2, b.id, billRows, 1).setFontColor(GREY);
  bills.setColumnWidth(b.name, 190);
  bills.setColumnWidth(b.notes, 240);
  bills.getRange(1, b.due).setNote('The next date this bill is due. The app works out the following months from "Repeats".');

  const shop = sheet_('shopping'), s = cache_.cols.shopping, shopRows = bodyRows_(shop);
  shop.getRange(2, s.added, shopRows, 1).setNumberFormat('mmm d, h:mm am/pm').setFontColor(GREY);
  [s.itemId, s.id].forEach(col => shop.getRange(2, col, shopRows, 1).setFontColor(GREY));
  shop.setColumnWidth(s.name, 190);

  const set = sheet_('settings');
  set.setColumnWidth(cache_.cols.settings.key, 140);
  set.getRange(1, cache_.cols.settings.key)
    .setNote('budget = your monthly budget (0 means none). currency = a 3-letter code such as JMD. ' +
      'household = Google emails (comma-separated) that may also use the app, besides the owner.');

  buildDashboard_(money);
  PropertiesService.getDocumentProperties().setProperty('setupVersion', SETUP_VERSION);
}

/** Header color, white bold headers, striped rows and the tab color. */
function styleSheet_(kind) {
  const sh = sheet_(kind), width = cache_.width[kind], [color, stripe] = THEME[kind];
  const all = sh.getRange(1, 1, sh.getMaxRows(), width);
  sh.setFrozenRows(1);
  sh.setTabColor(color);
  sh.getBandings().forEach(band => band.remove());
  all.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false)
    .setHeaderRowColor(color).setFirstRowColor('#ffffff').setSecondRowColor(stripe);
  all.setFontFamily(FONT).setVerticalAlignment('middle');
  sh.getRange(1, 1, 1, width).setBackground(null).setFontColor('#ffffff').setFontWeight('bold').setFontSize(11);
  sh.setRowHeight(1, 36);
  if (kind !== 'items') sh.setRowHeights(2, bodyRows_(sh), 26);
}

function formulaRule_(formula, range, background, fontColor) {
  const rule = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(formula).setBackground(background);
  if (fontColor) rule.setFontColor(fontColor).setBold(true);
  return rule.setRanges([range]).build();
}

function moneyFormat_() {
  const row = readAll_('settings').find(r => r.key === 'currency');
  const currency = String(row ? row.value : DEFAULT_SETTINGS.currency).trim().toUpperCase();
  return currency === 'JMD' ? '"J$"#,##0' : '#,##0.00';
}

function bodyRows_(sh) {
  return Math.max(sh.getMaxRows() - 1, 1);
}


/* ───────────────────────── Dashboard ───────────────────────── */

const DASHBOARD = 'Dashboard';
const PALETTE = ['#2e6a4c', '#d9822b', '#7b4fa0', '#2f6fb0', '#e2b93b', '#c2410c', '#0f9d8a', '#b23a7a', '#6b7280', '#8fbf6a'];
// Row layout of the Dashboard tab.
const DASH = {
  cards1: 4, cards2: 8, charts: [12, 29, 46],
  bills: 64, billRows: 25, recent: 93, recentRows: 10,
  data1: 107, data2: 123, footer: 136,
};

/**
 * Deletes and redraws the Dashboard tab. Everything on it is a formula that reads the
 * other tabs, so it stays up to date by itself.
 */
function buildDashboard_(money) {
  const ss = SpreadsheetApp.getActive();
  const old = ss.getSheetByName(DASHBOARD);
  if (old) ss.deleteSheet(old);
  const sh = ss.insertSheet(DASHBOARD, 0);
  const box = (row, col, rows, cols) => sh.getRange(row, col, rows || 1, cols || 1);
  const signed = '+' + money + ';-' + money + ';"same"';

  sh.setTabColor('#e2b93b');
  sh.setHiddenGridlines(true);
  sh.setColumnWidths(1, 12, 96);
  box(1, 1, sh.getMaxRows(), 12).setFontFamily(FONT).setVerticalAlignment('middle').setFontColor('#1c221e');

  // Column references into the other tabs, found by header name.
  const P = ref_('purchases'), I = ref_('items'), B = ref_('bills'), S = ref_('settings');
  const MS = '(EOMONTH(TODAY(),-1)+1)', ME = 'EOMONTH(TODAY(),0)', YS = 'DATE(YEAR(TODAY()),1,1)';
  const inMonth = P.date + ',">="&' + MS + ',' + P.date + ',"<="&' + ME;
  const budget = 'N(IFERROR(VLOOKUP("budget",' + S._ + '$' + S.keyL + ':$' + S.valueL + ',' +
    (cache_.cols.settings.value - cache_.cols.settings.key + 1) + ',FALSE),0))';

  // ── Banner
  box(1, 1, 1, 12).merge().setValue('   🏠  Home Dashboard').setBackground('#2e6a4c')
    .setFontColor('#ffffff').setFontSize(24).setFontWeight('bold');
  box(2, 1, 1, 12).merge()
    .setFormula('="     "&TEXT(TODAY(),"dddd, mmmm d, yyyy")&"   ·   everything in your house and every dollar you spend, in one place"')
    .setBackground('#dff0e6').setFontColor('#2e6a4c').setFontSize(11).setFontStyle('italic');
  sh.setRowHeight(1, 62);
  sh.setRowHeight(2, 30);
  sh.setRowHeight(3, 14);
  sh.setFrozenRows(2);

  // ── Summary cards: [label, value, hint, background, text color, number format]
  const bl = DASH.bills + 2, bh = bl + DASH.billRows - 1; // first/last bill row
  const card = (row, i, c) => {
    const col = 1 + i * 2;
    box(row, col, 1, 2).merge().setValue(c[0]).setFontWeight('bold').setFontSize(10);
    box(row + 1, col, 1, 2).merge().setValue(c[1]).setNumberFormat(c[5]).setFontSize(22).setFontWeight('bold');
    box(row + 2, col, 1, 2).merge().setValue(c[2]).setFontSize(9).setFontStyle('italic');
    box(row, col, 3, 2).setBackground(c[3]).setFontColor(c[4]).setHorizontalAlignment('center')
      .setBorder(true, true, true, true, null, null, '#ffffff', SpreadsheetApp.BorderStyle.SOLID_THICK);
  };
  [
    ['💸 Spent', '=SUMIFS(' + P.total + ',' + inMonth + ')', '="in "&TEXT(TODAY(),"mmmm")', '#fde7d3', '#b4501a', money],
    ['✅ Bills paid', '=SUMIFS(' + P.total + ',' + inMonth + ',' + P.billId + ',"<>")', 'this month', '#dff3e6', '#2e6a4c', money],
    ['⏳ Bills left to pay', '=SUM(F' + bl + ':F' + bh + ')',
      '=COUNTIF(G' + bl + ':G' + bh + ',"*Overdue*")&" overdue · "&COUNTIF(G' + bl + ':G' + bh + ',"*⏳*")&" coming up"', '#fde2dc', '#b42318', money],
    ['🎯 Budget left', '=IF(' + budget + '>0,' + budget + '-A5,"Not set")',
      '=IF(' + budget + '>0,"of "&TEXT(' + budget + ',"#,##0")&" this month","set it in the app ⚙️")', '#e3edfb', '#2f6fb0', money],
    ['🛒 Running low', '=SUMPRODUCT((' + I.name2 + '<>"")*(' + I.min2 + '<>"")*(' + I.qty2 + '<=' + I.min2 + '))',
      'items to restock', '#fff4d6', '#9a6b00', '0'],
    ['🏠 At home', '=COUNTA(' + I.name2 + ')', 'things you keep track of', '#efe6f7', '#7b4fa0', '0'],
  ].forEach((c, i) => card(DASH.cards1, i, c));
  const m0 = DASH.data1 + 2; // first month row in the data table
  [
    ['📈 vs last month', '=A5-SUMIFS(' + P.total + ',' + P.date + ',">="&EDATE(' + MS + ',-1),' + P.date + ',"<"&' + MS + ')',
      'more (+) or less (−) than last month', '#f1f3f0', '#1c221e', signed],
    ['📊 Monthly average', '=IFERROR(AVERAGEIF(D' + m0 + ':D' + (m0 + 10) + ',">0"),0)', 'over the last 11 months', '#eaf2fb', '#2f6fb0', money],
    ['🗓️ Spent this year', '=SUMIFS(' + P.total + ',' + P.date + ',">="&' + YS + ',' + P.date + ',"<="&' + ME + ')',
      '="since January 1, "&YEAR(TODAY())', '#f4eef9', '#7b4fa0', money],
    ['🏷️ Biggest buy', '=IFERROR(MAXIFS(' + P.total + ',' + inMonth + '),0)',
      '=IFERROR(INDEX(FILTER(' + P.name + ',' + P.total + '=G9,' + P.date + '>=' + MS + ',' + P.date + '<=' + ME + '),1),"this month")',
      '#fdf1e5', '#b4501a', money],
    ['🧾 Purchases', '=COUNTIFS(' + inMonth + ')', 'buys and payments this month', '#fff4d6', '#9a6b00', '0'],
    ['📦 Stock value', '=ARRAYFORMULA(SUM(IFERROR(' + I.qty2 + '*' + I.price2 + ',0)))', 'what is in the house is worth', '#eef6f1', '#2e6a4c', money],
  ].forEach((c, i) => card(DASH.cards2, i, c));
  [4, 8].forEach(r => { sh.setRowHeight(r, 26); sh.setRowHeight(r + 1, 46); sh.setRowHeight(r + 2, 22); });
  sh.setRowHeight(7, 10);
  sh.setRowHeight(11, 16);

  // ── Bills this month (one formula per cell, so the table never breaks when bills are deleted)
  const section = (row, col, cols, text, color) => box(row, col, 1, cols).merge().setValue(text)
    .setBackground(color).setFontColor('#ffffff').setFontWeight('bold').setFontSize(13);
  const heads = (row, col, values, color) => box(row, col, 1, values.length).setValues([values])
    .setBackground(color).setFontWeight('bold');
  section(DASH.bills, 1, 7, '  📅  Bills this month', THEME.bills[0]);
  heads(DASH.bills + 1, 1, ['Bill', '', 'Due', 'Usual', 'Paid', 'Left to pay', 'Status'], THEME.bills[1]);
  const names = [], rest = [];
  for (let k = 0; k < DASH.billRows; k++) {
    const row = billRowFormulas_(B, P, inMonth, k + 2);
    names.push([row[0]]);
    rest.push(row.slice(1));
  }
  box(bl, 1, DASH.billRows, 1).setFormulas(names);
  box(bl, 3, DASH.billRows, 5).setFormulas(rest);
  box(DASH.bills + 1, 1, DASH.billRows + 1, 2).mergeAcross();
  box(bl, 3, DASH.billRows, 1).setNumberFormat('mmm d');
  box(bl, 4, DASH.billRows, 3).setNumberFormat(money);
  box(bl, 1, DASH.billRows, 7).setBorder(null, null, true, null, null, true, '#ececec', SpreadsheetApp.BorderStyle.SOLID);

  // ── Running low
  section(DASH.bills, 9, 4, '  🛒  Running low', THEME.shopping[0]);
  heads(DASH.bills + 1, 9, ['Item', 'Have', 'Restock at', 'Price'], THEME.shopping[1]);
  box(bl, 9).setFormula('=IFERROR(SORT(FILTER({' + [I.name2, I.qty2, I.min2, I.price2].join(',') + '},' +
    I.name2 + '<>"",' + I.min2 + '<>"",' + I.qty2 + '<=' + I.min2 + '),2,TRUE),"Nothing running low 🎉")');
  box(bl, 9, DASH.billRows, 1).setWrap(true);
  box(bl, 10, DASH.billRows, 2).setHorizontalAlignment('center');
  box(bl, 12, DASH.billRows, 1).setNumberFormat(money);
  box(bl, 9, DASH.billRows, 4).setBorder(null, null, true, null, null, true, '#ececec', SpreadsheetApp.BorderStyle.SOLID);

  // ── Recent purchases and what you buy most
  const rl = DASH.recent + 2;
  section(DASH.recent, 1, 5, '  🧾  Latest purchases', THEME.purchases[0]);
  heads(DASH.recent + 1, 1, ['Date', 'Item', 'Category', 'Where', 'Total'], THEME.purchases[1]);
  box(rl, 1).setFormula('=IFERROR(ARRAY_CONSTRAIN(SORT(FILTER({' + [P.date2, P.name2, P.category2, P.store2, P.total2].join(',') + '},' +
    P.date2 + '<>""),1,FALSE),' + DASH.recentRows + ',5),"No purchases yet")');
  box(rl, 1, DASH.recentRows, 1).setNumberFormat('ddd mmm d');
  box(rl, 2, DASH.recentRows, 1).setWrap(true);
  box(rl, 5, DASH.recentRows, 1).setNumberFormat(money);
  box(rl, 1, DASH.recentRows, 5).setBorder(null, null, true, null, null, true, '#ececec', SpreadsheetApp.BorderStyle.SOLID);

  const PDATA = '{' + [P.date2, P.name2 + '&""', 'IF(' + P.category2 + '="","Other",' + P.category2 + ')',
    P.total2, P.store2 + '&""', P.billId2 + '&""'].join(',') + '}'; // Col1 date … Col6 bill id
  const IDATA = '{' + ['IF(' + I.category2 + '="","Other",' + I.category2 + ')', I.name2 + '&""',
    'IFERROR(' + I.qty2 + '*' + I.price2 + ',0)'].join(',') + '}'; // Col1 category, Col2 name, Col3 value
  const day = d => 'date \'"&TEXT(' + d + ',"yyyy-mm-dd")&"\'';
  const query = (data, select, empty) => '=IFERROR(ARRAYFORMULA(QUERY(' + data + ',"' + select + '",0)),{' + empty + '})';

  section(DASH.recent, 9, 4, '  ⭐  What you buy most this year', THEME.items[0]);
  box(DASH.recent + 1, 9, 1, 4).setBackground(THEME.items[1]).setFontWeight('bold');
  box(DASH.recent + 1, 9).setFormula(query(PDATA,
    'select Col2, count(Col2), sum(Col4) where Col6 = \'\' and Col1 >= ' + day(YS) + ' and Col4 > 0 ' +
    'group by Col2 order by sum(Col4) desc limit ' + DASH.recentRows + " label Col2 'Item', count(Col2) 'Times', sum(Col4) 'Spent'",
    '"Item","Times","Spent";"Nothing yet",0,0'));
  box(rl, 9, DASH.recentRows, 1).setWrap(true);
  box(rl, 10, DASH.recentRows, 1).setHorizontalAlignment('center');
  box(rl, 11, DASH.recentRows, 1).setNumberFormat(money);
  box(rl, 9, DASH.recentRows, 4).setBorder(null, null, true, null, null, true, '#ececec', SpreadsheetApp.BorderStyle.SOLID);

  // ── The numbers behind the charts
  section(DASH.data1, 1, 12, '  📊  The numbers behind the charts', THEME.settings[0]);
  const h1 = DASH.data1 + 1, h2 = DASH.data2;
  [h1, h2].forEach(r => box(r, 1, 1, 12).setBackground('#f2f3f5').setFontWeight('bold'));
  box(h1, 1, 1, 5).setValues([['Month', 'Bills', 'Shopping & other', 'Total', 'Month starts']]);
  const months = [];
  for (let k = 0; k < 12; k++) {
    const r = m0 + k, start = 'E' + r;
    months.push([
      '=TEXT(' + start + ',"mmm yy")',
      '=SUMIFS(' + P.total + ',' + P.date + ',">="&' + start + ',' + P.date + ',"<"&EDATE(' + start + ',1),' + P.billId + ',"<>")',
      '=D' + r + '-B' + r,
      '=SUMIFS(' + P.total + ',' + P.date + ',">="&' + start + ',' + P.date + ',"<"&EDATE(' + start + ',1))',
      '=EDATE(' + MS + ',' + (k - 11) + ')',
    ]);
  }
  box(m0, 1, 12, 5).setFormulas(months);
  box(m0, 2, 12, 3).setNumberFormat(money);
  box(m0, 5, 12, 1).setNumberFormat('mmm d, yyyy').setFontColor(GREY);

  box(h1, 7).setFormula(query(PDATA,
    'select Col3, sum(Col4) where Col1 >= ' + day(MS) + ' and Col1 <= ' + day(ME) + ' and Col4 > 0 ' +
    "group by Col3 order by sum(Col4) desc limit 10 label Col3 'Category', sum(Col4) 'This month'",
    '"Category","This month";"Nothing yet",0'));
  box(h1, 10).setFormula(query(PDATA,
    "select Col2, sum(Col4) where Col6 <> '' and Col1 >= " + day(YS) + ' and Col4 > 0 ' +
    "group by Col2 order by sum(Col4) desc limit 10 label Col2 'Bill', sum(Col4) 'Paid this year'",
    '"Bill","Paid this year";"No bills paid yet",0'));
  box(h2, 1).setFormula(query(PDATA,
    "select Col5, sum(Col4) where Col5 <> '' and Col1 >= " + day(YS) + ' and Col4 > 0 ' +
    "group by Col5 order by sum(Col4) desc limit 10 label Col5 'Store', sum(Col4) 'Spent this year'",
    '"Store","Spent this year";"No stores yet",0'));
  box(h2, 4).setFormula(query(PDATA,
    'select Col3, sum(Col4) where Col1 >= ' + day(YS) + ' and Col4 > 0 ' +
    "group by Col3 order by sum(Col4) desc limit 10 label Col3 'Category', sum(Col4) 'This year'",
    '"Category","This year";"Nothing yet",0'));
  box(h2, 7).setFormula(query(IDATA,
    "select Col1, count(Col2) where Col2 <> '' group by Col1 order by count(Col2) desc limit 10 " +
    "label Col1 'Category', count(Col2) 'Items'", '"Category","Items";"No items yet",0'));
  box(h2, 10).setFormula(query(IDATA,
    "select Col1, sum(Col3) where Col2 <> '' group by Col1 order by sum(Col3) desc limit 10 " +
    "label Col1 'Category', sum(Col3) 'Stock value'", '"Category","Stock value";"No items yet",0'));
  [[h1 + 1, 8], [h1 + 1, 11], [h2 + 1, 2], [h2 + 1, 5], [h2 + 1, 11]].forEach(([r, c]) => box(r, c, 10, 1).setNumberFormat(money));

  box(DASH.footer, 1, 1, 12).merge()
    .setValue('This page builds itself from your other tabs, so there is nothing to type here. ' +
      'To redraw the colors and charts: menu 🏠 Inventory → Rebuild dashboard & colors.')
    .setFontColor(GREY).setFontStyle('italic').setFontSize(9);

  // ── Status colors and +/- colors
  const status = box(bl, 7, DASH.billRows, 1);
  const tag = (text, bg, fg) => SpreadsheetApp.newConditionalFormatRule().whenTextContains(text)
    .setBackground(bg).setFontColor(fg).setBold(true).setRanges([status]).build();
  sh.setConditionalFormatRules([
    tag('⚠', '#fde2dc', '#b42318'), tag('✅', '#dff3e6', '#2e6a4c'), tag('⏳', '#fff4d6', '#9a6b00'),
    tag('💤', '#f2f3f5', '#6b7280'), tag('🔄', '#eaf2fb', '#2f6fb0'),
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setFontColor('#b42318')
      .setRanges([box(DASH.cards1 + 1, 7, 1, 2)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setFontColor('#b42318')
      .setRanges([box(DASH.cards2 + 1, 1, 1, 2)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setFontColor('#2e6a4c')
      .setRanges([box(DASH.cards2 + 1, 1, 1, 2)]).build(),
  ]);

  // ── Charts
  const chart = (type, range, row, col, title, options) => {
    let builder = sh.newChart().setChartType(type).addRange(range).setNumHeaders(1)
      .setPosition(row, col, 6, 0)
      .setOption('title', title)
      .setOption('width', 560).setOption('height', 330)
      .setOption('fontName', FONT)
      .setOption('titleTextStyle', { fontSize: 15, bold: true, color: '#1c221e' })
      .setOption('legend', { position: 'bottom' });
    Object.keys(options).forEach(k => { builder = builder.setOption(k, options[k]); });
    sh.insertChart(builder.build());
  };
  const [c1, c2, c3] = DASH.charts;
  chart(Charts.ChartType.COLUMN, box(h1, 1, 13, 3), c1, 1, 'Spending by month: bills vs shopping',
    { isStacked: true, colors: [THEME.bills[0], THEME.shopping[0]] });
  chart(Charts.ChartType.PIE, box(h1, 7, 11, 2), c1, 7, 'This month by category',
    { pieHole: 0.45, colors: PALETTE, legend: { position: 'right' } });
  chart(Charts.ChartType.COLUMN, box(h2, 4, 11, 2), c2, 1, 'This year by category',
    { colors: [THEME.purchases[0]], legend: { position: 'none' } });
  chart(Charts.ChartType.BAR, box(h1, 10, 11, 2), c2, 7, 'Bills paid this year',
    { colors: [THEME.bills[0]], legend: { position: 'none' } });
  chart(Charts.ChartType.BAR, box(h2, 1, 11, 2), c3, 1, 'Where your money goes (stores)',
    { colors: [THEME.shopping[0]], legend: { position: 'none' } });
  chart(Charts.ChartType.PIE, box(h2, 7, 11, 2), c3, 7, 'What you keep at home',
    { pieHole: 0.45, colors: PALETTE, legend: { position: 'right' } });

  // Typing here would break formulas, so warn (but don't block) anyone who tries.
  sh.protect().setDescription('Dashboard builds itself').setWarningOnly(true);
  ss.setActiveSheet(sh);
}

/**
 * Formulas for one row of "Bills this month", reading bill number n (its row in the Bills tab):
 * [name, due date, usual amount, paid, left to pay, status].
 */
function billRowFormulas_(B, P, inMonth, n) {
  const at = key => 'INDEX(' + B[key] + ',' + n + ')';
  const repeats = '{' + BILL_REPEATS.map(r => '"' + r + '"').join(',') + '}';
  const vars = [
    'b_name', at('name'),
    'b_amt', 'N(' + at('amount') + ')',
    'b_due', at('due'),
    'b_id', at('id'),
    'b_every', 'IFERROR(CHOOSE(MATCH(' + at('repeat') + ',' + repeats + ',0),1,2,3,4,6,12,0),1)',
    'b_diff', 'IF(ISNUMBER(b_due),(YEAR(TODAY())-YEAR(b_due))*12+MONTH(TODAY())-MONTH(b_due),0)',
    'b_now', 'IF(b_every=0,FALSE,IF(ISNUMBER(b_due),AND(b_diff>=0,MOD(b_diff,b_every)=0),b_every=1))',
    'b_day', 'IF(AND(b_now,ISNUMBER(b_due)),DATE(YEAR(TODAY()),MONTH(TODAY()),MIN(DAY(b_due),DAY(EOMONTH(TODAY(),0)))),"")',
    'b_paid', 'IF(b_id="",0,SUMIFS(' + P.total + ',' + P.billId + ',b_id,' + inMonth + '))',
    'b_left', 'IF(b_every=0,MAX(0,b_amt-b_paid),IF(AND(b_now,b_paid=0),b_amt,0))',
    'b_status', 'IF(b_every=0,"🔄 As you go",IF(b_paid>0,"✅ Paid",IF(NOT(b_now),"💤 Not due",' +
      'IF(AND(ISNUMBER(b_day),b_day<TODAY()),"⚠️ Overdue","⏳ Due"))))',
  ];
  const cell = expr => '=LET(' + vars.join(',') + ',IF(b_name="","",' + expr + '))';
  return [cell('b_name'), cell('b_day'), cell('IF(b_amt=0,"",b_amt)'), cell('IF(b_paid=0,"",b_paid)'),
    cell('IF(b_left=0,"",b_left)'), cell('b_status')];
}

/** For each column of a tab: whole column (key), body without header (key2) and letter (keyL). */
function ref_(kind) {
  sheet_(kind);
  const out = { _: "'" + SCHEMA[kind].sheet + "'!" };
  Object.keys(cache_.cols[kind]).forEach(key => {
    const l = colA1_(cache_.cols[kind][key]);
    out[key] = out._ + '$' + l + ':$' + l;
    out[key + '2'] = out._ + '$' + l + '$2:$' + l;
    out[key + 'L'] = l;
  });
  return out;
}

function ensureSetup_() {
  if (PropertiesService.getDocumentProperties().getProperty('setupVersion') !== SETUP_VERSION) setup();
}


/* ───────────────────────── API used by the app ───────────────────────── */

function getData() {
  ensureSetup_();
  const settings = Object.assign({}, DEFAULT_SETTINGS);
  readAll_('settings').forEach(r => { if (r.key) settings[r.key] = r.value; });
  return {
    items: readAll_('items').map(r => ({
      id: r.id, name: String(r.name), category: String(r.category), location: String(r.location),
      qty: num_(r.qty), min: num_(r.min, ''), unit: String(r.unit), price: num_(r.price, ''),
      store: String(r.store), notes: String(r.notes), updated: r.updated,
      photoId: String(r.photoId), photoUrl: r.photoId ? photoUrl_(r.photoId) : '',
    })),
    shopping: readAll_('shopping').map(strip_),
    bills: readAll_('bills').map(r => ({
      id: r.id, name: String(r.name), category: String(r.category), amount: num_(r.amount, ''),
      repeat: String(r.repeat || 'Monthly'), due: String(r.due).slice(0, 10), notes: String(r.notes),
    })),
    purchases: readAll_('purchases').map(strip_),
    settings: { budget: num_(settings.budget), currency: String(settings.currency || 'JMD') },
    sheetUrl: SpreadsheetApp.getActive().getUrl(),
  };
}

function saveItem(item, photoDataUrl) {
  return locked_(() => {
    const name = String(item.name || '').trim();
    if (!name) throw new Error('Please give the item a name.');
    const old = item.id ? readAll_('items').find(r => r.id === item.id) : null;
    const rec = {
      id: old ? old.id : '',
      name: text_(name), category: text_(item.category), location: text_(item.location),
      qty: num_(item.qty), min: num_(item.min, ''), unit: text_(item.unit),
      price: num_(item.price, ''), store: text_(item.store), notes: text_(item.notes),
      updated: now_(),
    };
    if (photoDataUrl) {
      rec.photoId = savePhoto_(photoDataUrl, name);
      if (old && old.photoId) trash_(old.photoId);
    } else if (item.removePhoto && old && old.photoId) {
      trash_(old.photoId);
      rec.photoId = '';
    }
    if ('photoId' in rec) rec.photo = rec.photoId ? '=IMAGE("' + photoUrl_(rec.photoId) + '")' : '';
    const row = write_('items', rec);
    if (rec.photoId) sheet_('items').setRowHeight(row, 60);
    return getData();
  });
}

function setQty(id, qty) {
  return locked_(() => {
    const row = findRow_('items', id);
    if (!row) throw new Error('That item no longer exists — pull to refresh.');
    const sh = sheet_('items'), c = cache_.cols.items;
    sh.getRange(row, c.qty).setValue(num_(qty));
    sh.getRange(row, c.updated).setValue(now_());
    return { id: id, qty: num_(qty) };
  });
}

function deleteItem(id) {
  return locked_(() => {
    const item = readAll_('items').find(r => r.id === id);
    if (item) {
      if (item.photoId) trash_(item.photoId);
      deleteRow_('items', item._row);
    }
    readAll_('shopping').filter(s => s.itemId === id).reverse()
      .forEach(s => deleteRow_('shopping', s._row));
    return getData();
  });
}

function addToShopping(entry) {
  return locked_(() => {
    const name = String(entry.name || '').trim();
    if (!name) throw new Error('What do you need to buy?');
    const existing = readAll_('shopping').find(s =>
      (entry.itemId && s.itemId === entry.itemId) || String(s.name).toLowerCase() === name.toLowerCase());
    write_('shopping', {
      id: existing ? existing.id : '', name: text_(name), qty: num_(entry.qty) || 1,
      category: text_(entry.category || (existing && existing.category)),
      itemId: entry.itemId || (existing && existing.itemId) || '', added: now_(),
    });
    return getData();
  });
}

function removeFromShopping(id) {
  return locked_(() => {
    const row = findRow_('shopping', id);
    if (row) deleteRow_('shopping', row);
    return getData();
  });
}

/**
 * Logs a purchase, adds it to stock and takes it off the shopping list.
 * Bill payments come through here too (with a billId) and never touch stock.
 */
function buy(p) {
  return locked_(() => {
    const name = String(p.name || '').trim();
    if (!name) throw new Error('What did you buy?');
    const qty = num_(p.qty) || 1;
    const price = num_(p.price, '');
    const items = p.billId ? [] : readAll_('items');
    const item = items.find(i => p.itemId && i.id === p.itemId) ||
      items.find(i => String(i.name).toLowerCase() === name.toLowerCase());

    write_('purchases', {
      date: /^\d{4}-\d{2}-\d{2}$/.test(p.date) ? p.date : today_(),
      name: text_(name),
      category: text_(p.category || (item && item.category)),
      qty: qty,
      price: price,
      total: (row, c) => '=IF(' + colA1_(c.price) + row + '="","",' +
        colA1_(c.qty) + row + '*' + colA1_(c.price) + row + ')',
      store: text_(p.store),
      itemId: item ? item.id : '',
      billId: p.billId || '',
    });

    if (item) {
      const sh = sheet_('items'), c = cache_.cols.items;
      if (p.addToStock !== false) sh.getRange(item._row, c.qty).setValue(num_(item.qty) + qty);
      if (price !== '') sh.getRange(item._row, c.price).setValue(price);
      if (String(p.store || '').trim()) sh.getRange(item._row, c.store).setValue(text_(p.store));
      sh.getRange(item._row, c.updated).setValue(now_());
    }

    readAll_('shopping')
      .filter(s => (p.shoppingId && s.id === p.shoppingId) || (item && s.itemId === item.id))
      .reverse()
      .forEach(s => deleteRow_('shopping', s._row));
    return getData();
  });
}

function saveBill(bill) {
  return locked_(() => {
    const name = String(bill.name || '').trim();
    if (!name) throw new Error('Please give the bill a name.');
    const repeat = BILL_REPEATS.indexOf(bill.repeat) === -1 ? 'Monthly' : bill.repeat;
    write_('bills', {
      id: bill.id || '', name: text_(name), category: text_(bill.category),
      amount: num_(bill.amount, ''), repeat: repeat,
      due: repeat !== 'As needed' && /^\d{4}-\d{2}-\d{2}$/.test(bill.due) ? bill.due : '',
      notes: text_(bill.notes),
    });
    return getData();
  });
}

/** Removes the bill. Its past payments stay in Purchases as spending history. */
function deleteBill(id) {
  return locked_(() => {
    const row = findRow_('bills', id);
    if (row) deleteRow_('bills', row);
    return getData();
  });
}

function deletePurchase(id) {
  return locked_(() => {
    const row = findRow_('purchases', id);
    if (row) deleteRow_('purchases', row);
    return getData();
  });
}

function saveSettings(s) {
  return locked_(() => {
    const rows = readAll_('settings');
    const moneyBefore = moneyFormat_();
    Object.keys(DEFAULT_SETTINGS).forEach(key => {
      if (!(key in s)) return;
      const value = key === 'budget' ? num_(s.budget)
        : key === 'currency' ? String(s[key]).trim().toUpperCase() : String(s[key]).trim();
      const r = rows.find(x => x.key === key);
      if (r) sheet_('settings').getRange(r._row, cache_.cols.settings.value).setValue(value);
      else write_('settings', { key: key, value: value });
    });
    // A new currency changes how money looks everywhere in the sheet.
    if (moneyFormat_() !== moneyBefore) setup();
    return getData();
  });
}


/* ───────────────────────── Photos (stored in Google Drive) ───────────────────────── */

function savePhoto_(dataUrl, name) {
  const m = /^data:(image\/[\w+.-]+);base64,(.+)$/.exec(String(dataUrl));
  if (!m) throw new Error('That photo could not be read.');
  const fileName = (name.replace(/[^\w -]/g, '').trim() || 'item') + '.jpg';
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], fileName);
  const file = photoFolder_().createFile(blob);
  // Link-only sharing lets the app and the sheet's IMAGE() show the photo.
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  return file.getId();
}

function photoFolder_() {
  const props = PropertiesService.getDocumentProperties();
  const id = props.getProperty('photoFolderId');
  if (id) {
    try {
      const folder = DriveApp.getFolderById(id);
      if (!folder.isTrashed()) return folder;
    } catch (e) {}
  }
  const folder = DriveApp.createFolder(PHOTO_FOLDER);
  props.setProperty('photoFolderId', folder.getId());
  return folder;
}

function photoUrl_(id) {
  return 'https://drive.google.com/thumbnail?id=' + id + '&sz=w600';
}

function trash_(id) {
  try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {}
}


/* ───────────────────────── Sheet helpers ───────────────────────── */

/** Returns the sheet for a kind, creating it and any missing header columns. */
function sheet_(kind) {
  if (cache_.sheets[kind]) return cache_.sheets[kind];
  const s = SCHEMA[kind], ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(s.sheet) || ss.insertSheet(s.sheet);
  const lastCol = sh.getLastColumn();
  const have = lastCol ? sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim()) : [];
  const missing = s.cols.map(c => c[1]).filter(h => have.indexOf(h) === -1);
  if (missing.length) {
    const need = have.length + missing.length;
    if (sh.getMaxColumns() < need) sh.insertColumnsAfter(sh.getMaxColumns(), need - sh.getMaxColumns());
    sh.getRange(1, have.length + 1, 1, missing.length).setValues([missing]);
    have.push.apply(have, missing);
  }
  const cols = {};
  s.cols.forEach(([key, header]) => { cols[key] = have.indexOf(header) + 1; });
  cache_.sheets[kind] = sh;
  cache_.cols[kind] = cols;
  cache_.width[kind] = have.length;
  return sh;
}

/** Reads every non-empty row as an object. Rows typed in by hand get an ID. */
function readAll_(kind) {
  const sh = sheet_(kind), cols = cache_.cols[kind], width = cache_.width[kind];
  const last = sh.getLastRow();
  if (last < 2) return [];
  const values = sh.getRange(2, 1, last - 1, width).getValues();
  const rows = [];
  values.forEach((r, i) => {
    const o = { _row: i + 2 };
    let any = false;
    SCHEMA[kind].cols.forEach(([key]) => {
      if (key === 'photo') return;
      o[key] = cols[key] ? plain_(r[cols[key] - 1]) : '';
      if (o[key] !== '' && key !== 'id') any = true;
    });
    if (!any) return;
    if (cols.id && !o.id) {
      o.id = newId_();
      sh.getRange(o._row, cols.id).setValue(o.id);
    }
    rows.push(o);
  });
  return rows;
}

/**
 * Creates or updates a row. Values may be functions of (row, cols) for
 * formulas that refer to their own row. Returns the row number.
 */
function write_(kind, obj) {
  const sh = sheet_(kind), cols = cache_.cols[kind], width = cache_.width[kind];
  const value = (key, row) => typeof obj[key] === 'function' ? obj[key](row, cols) : obj[key];
  let row = findRow_(kind, obj.id);
  if (row) {
    // Only touch our own cells so hand-made formulas and extra columns survive.
    Object.keys(obj).forEach(key => {
      if (cols[key] && key !== 'id') sh.getRange(row, cols[key]).setValue(value(key, row));
    });
    return row;
  }
  if (cols.id && !obj.id) obj.id = newId_();
  row = sh.getLastRow() + 1;
  if (row > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 1);
  const vals = new Array(width).fill('');
  Object.keys(obj).forEach(key => { if (cols[key]) vals[cols[key] - 1] = value(key, row); });
  sh.getRange(row, 1, 1, width).setValues([vals]);
  return row;
}

function findRow_(kind, id) {
  if (!id) return 0;
  const sh = sheet_(kind), c = cache_.cols[kind].id, last = sh.getLastRow();
  if (!c || last < 2) return 0;
  const ids = sh.getRange(2, c, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return 0;
}

function deleteRow_(kind, row) {
  const sh = sheet_(kind);
  // Sheets refuses to delete the last remaining body row, so clear it instead.
  if (sh.getMaxRows() <= 2) sh.getRange(row, 1, 1, sh.getMaxColumns()).clearContent();
  else sh.deleteRow(row);
}

function locked_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureSetup_();
    return fn();
  } finally {
    SpreadsheetApp.flush();
    lock.releaseLock();
  }
}

function strip_(r) {
  const o = Object.assign({}, r);
  delete o._row;
  return o;
}

/** Sheet values → things google.script.run can send (it can't send Dates). */
function plain_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, tz_(), "yyyy-MM-dd'T'HH:mm:ss");
  if (v === null || v === undefined || typeof v === 'object') return '';
  return v;
}

function num_(v, ifBlank) {
  if (v === '' || v === null || v === undefined) return ifBlank === undefined ? 0 : ifBlank;
  const n = typeof v === 'number' ? v : Number(String(v).replace(/[^0-9.\-]/g, ''));
  return isFinite(n) ? n : (ifBlank === undefined ? 0 : ifBlank);
}

/** A leading apostrophe makes Sheets keep the text exactly as typed. */
function text_(v) {
  const s = String(v === null || v === undefined ? '' : v).trim();
  return s ? "'" + s : '';
}

function newId_() {
  return 'x' + Utilities.getUuid().replace(/-/g, '').slice(0, 11);
}

function tz_() {
  return cache_.tz || (cache_.tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone());
}

function now_() {
  return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm');
}

function today_() {
  return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd');
}

function colA1_(n) {
  let s = '';
  for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s;
  return s;
}
