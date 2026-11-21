// AmtBuddy survey → Google Sheet
// 1. Put your Sheet ID below (the long part of the Sheet URL between /d/ and /edit).
// 2. Run setup() once (function menu → setup → Run). It is SAFE to run again: it never deletes responses.
// 3. Deploy → New deployment → Web app. Execute as: Me. Who has access: Anyone. Copy the /exec URL.
//    Later changes: Deploy → Manage deployments → pencil → Version: New version → Deploy (URL stays the same).

const SHEET_ID = '1zR5wPbpi3hhtlgX-tF-QLDWYGALMjyBrDiEzwHvKLbU';
const TZ = 'Europe/Berlin';

const COLUMNS = [
  ['Submitted (Berlin time)', 150],
  ['Email',                   220],
  ['City',                    130],
  ['Situation',               170],
  ['Biggest stress',          210],
  ['Letters (last 3 months)', 150],
  ['Worth for 3 months',      150],
  ['Language',                110],
  ['Source',                  200],
];

const LABELS = {
  pain: {
    'appointment': 'Getting an appointment',
    'documents':   'Which documents I need',
    'letters':     'Understanding letters',
    'no-answer':   'Office not answering',
  },
  letters: { '0': 'None', '1-2': '1 to 2', '3-5': '3 to 5', '6+': '6 or more' },
  price:   { '0': 'Free only', '5': 'About €5', '9': 'About €9', '15': 'About €15', '20+': '€20 or more' },
};
const label_ = (field, raw) => (raw === undefined || raw === '') ? '' : (LABELS[field][raw] || String(raw));

function book_() { return SpreadsheetApp.openById(SHEET_ID); }

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const sh = book_().getSheetByName('Responses') || setup().getSheetByName('Responses');
    sh.appendRow([
      Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'),
      String(d.email || '').trim(),
      d.city || '',
      d.situation || '',
      label_('pain', d.pain),
      label_('letters', d.letters),
      label_('price', d.price),
      d.lang || '',
      d.source || '',
    ]);
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput('ok');
}

// Creates the Responses and Summary tabs. Never deletes existing responses.
function setup() {
  const ss = book_();
  let sh = ss.getSheetByName('Responses');
  if (!sh) {
    const first = ss.getSheets()[0];
    sh = (first.getLastRow() === 0 && ss.getSheets().length === 1) ? first.setName('Responses') : ss.insertSheet('Responses', 0);
  }
  sh.getRange(1, 1, 1, COLUMNS.length).setValues([COLUMNS.map(c => c[0])])
    .setFontWeight('bold').setBackground('#16181d').setFontColor('#ffdb3d');
  COLUMNS.forEach((c, i) => sh.setColumnWidth(i + 1, c[1]));
  sh.setFrozenRows(1);
  if (sh.getFilter() === null) sh.getRange(1, 1, Math.max(sh.getMaxRows(), 2), COLUMNS.length).createFilter();

  const sum = ss.getSheetByName('Summary') || ss.insertSheet('Summary');
  sum.clear();
  const rows = [['AmtBuddy survey summary', ''], ['Total responses', '=COUNTA(Responses!B2:B)'],
                ['Would pay something (€5+)', '=IFERROR((COUNTA(Responses!G2:G)-COUNTIF(Responses!G2:G,"Free only"))/COUNTA(Responses!G2:G),0)'], ['', '']];
  const block = (title, col, options) => {
    rows.push([title, 'Count']);
    options.forEach(o => rows.push([o, `=COUNTIF(Responses!${col}2:${col},"${o}")`]));
    rows.push(['', '']);
  };
  block('Situation', 'D', ['EU Blue Card', 'Chancenkarte', 'Work permit (other)', 'Student visa', 'Family reunion', 'Permanent residence', 'Citizenship', 'Other']);
  block('City', 'C', ['Berlin', 'Munich', 'Hamburg', 'Frankfurt', 'Cologne', 'Stuttgart', 'Düsseldorf', 'Other city', 'Not in Germany yet']);
  block('Biggest stress', 'E', Object.values(LABELS.pain));
  block('Letters (last 3 months)', 'F', Object.values(LABELS.letters));
  block('Worth for 3 months', 'G', Object.values(LABELS.price));
  block('Language', 'H', ['English', 'Bangla', 'Hindi', 'Urdu', 'Arabic', 'Turkish', 'Other']);
  sum.getRange(1, 1, rows.length, 2).setValues(rows);
  sum.getRange('A1').setFontSize(14).setFontWeight('bold');
  sum.getRange('B3').setNumberFormat('0%');
  rows.forEach((r, i) => { if (r[1] === 'Count') sum.getRange(i + 1, 1, 1, 2).setFontWeight('bold').setBackground('#ffdb3d'); });
  sum.setColumnWidth(1, 240); sum.setColumnWidth(2, 90);
  return ss;
}
