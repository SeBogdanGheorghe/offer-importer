/* Self-check for the pure server logic in Code.gs. Run: node selfcheck.js */
const fs = require('fs');
eval(fs.readFileSync(__dirname + '/Code.gs', 'utf8'));

let failures = 0;
function check(label, actual, expected) {
  const ok = JSON.stringify(actual) === JSON.stringify(expected);
  if (!ok) { failures++; console.log(`FAIL ${label}: expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`); }
  else console.log(`ok   ${label}`);
}

/* --- 3PP form vs update form -------------------------------------------- */
check('3PP form detected from Offer Analysis + Offer Details',
  detectFileType_(['Offer Analysis', 'Offer Details', 'Rates (1) INAO', 'Allocation (1)']), '3PP form');
check('update form when only rates/allocation tabs',
  detectFileType_(['Rates (1)', 'Allocation (1)']), 'update form');
check('update form when only Offer Details present',
  detectFileType_(['Offer Details', 'Rates (1)']), 'update form');
check('3PP detection is case/space tolerant',
  detectFileType_(['  offer   analysis ', 'OFFER DETAILS']), '3PP form');

/* rows logged before the rename must not disappear from the filter */
check('legacy "EF" rows read back as 3PP form', normalizeFileType_('EF'), '3PP form');
check('legacy value is matched case-insensitively', normalizeFileType_('ef'), '3PP form');
check('update form rows are left alone', normalizeFileType_('update form'), 'update form');
check('already-renamed rows are left alone', normalizeFileType_('3PP form'), '3PP form');
check('a blank type stays blank', normalizeFileType_(''), '');

/* --- search ------------------------------------------------------------ */
const target = { fileName: 'RATE UPDATE_Amsterdam & Berlin with European Sleeper 13.10.2026',
                 email: 'b.gheorghe@secretescapes.com', fileType: 'update form' };

const matches = (q) => matchesTokens_(searchTokens_(q), target);

check('the required case: "Amsterdam & Berlin" finds the file', matches('Amsterdam & Berlin'), true);
check('partial word run across the underscore', matches('RATE UPDATE Amsterdam'), true);
check('out-of-order tokens still match', matches('Berlin Amsterdam'), true);
check('case insensitive', matches('amsterdam berlin'), true);
check('date fragment matches', matches('13.10.2026'), true);
check('searching by person works', matches('gheorghe'), true);
check('unrelated term does not match', matches('Paris'), false);
check('one good + one bad token does not match', matches('Amsterdam Paris'), false);
check('empty search matches everything', matches(''), true);

/* --- date ranges -------------------------------------------------------- */
function expectedStart(days) {
  const date = new Date();
  date.setDate(date.getDate() - (days - 1));
  return new Date(date.getFullYear(), date.getMonth(), date.getDate()).getTime();
}

const sevenDays = resolveRange_({});
check('default period starts 7 calendar days back (today + 6)', sevenDays.from.getTime(), expectedStart(7));
check('default period has no upper bound', sevenDays.to, null);
check('30-day period', resolveRange_({ days: 30 }).from.getTime(), expectedStart(30));
check('period always starts at midnight', [sevenDays.from.getHours(), sevenDays.from.getMinutes()], [0, 0]);

const custom = resolveRange_({ mode: 'custom', from: '2026-10-01', to: '2026-10-07' });
check('custom from = start of day', [custom.from.getFullYear(), custom.from.getMonth() + 1, custom.from.getDate(), custom.from.getHours()], [2026, 10, 1, 0]);
check('custom to = end of day (inclusive)', [custom.to.getDate(), custom.to.getHours(), custom.to.getMinutes()], [7, 23, 59]);

/* --- misc --------------------------------------------------------------- */
check('extension stripped for the Drive file name',
  stripExtension_('OFFER FORM - Mein Schiff 7 LPA (3) (1).xlsx'), 'OFFER FORM - Mein Schiff 7 LPA (3) (1)');

/* --- unique download names (pulled out of App.html) --------------------- */
const appSource = fs.readFileSync(__dirname + '/App.html', 'utf8');
const fnStart = appSource.indexOf('function withUniqueFilenames');
const fnEnd = appSource.indexOf('\n}', fnStart) + 2;
eval(appSource.slice(fnStart, fnEnd));

const names = (files) => withUniqueFilenames(files).map((file) => file.filename);

check('the reported case: 5 Nts / 5 Nts / 7 Nts / 7 Nts all stay distinct',
  names([
    { filename: '5 Nts upload template - RATE UPDATE.xlsx' },
    { filename: '5 Nts upload template - RATE UPDATE.xlsx' },
    { filename: '7 Nts upload template - RATE UPDATE.xlsx' },
    { filename: '7 Nts upload template - RATE UPDATE.xlsx' },
  ]),
  [
    '5 Nts upload template - RATE UPDATE.xlsx',
    '5 Nts upload template - RATE UPDATE (2).xlsx',
    '7 Nts upload template - RATE UPDATE.xlsx',
    '7 Nts upload template - RATE UPDATE (2).xlsx',
  ]);

check('already-distinct names are left alone',
  names([{ filename: 'Inao.xlsx' }, { filename: 'Auao.xlsx' }]),
  ['Inao.xlsx', 'Auao.xlsx']);

check('three of a kind counts up',
  names([{ filename: 'a.xlsx' }, { filename: 'a.xlsx' }, { filename: 'a.xlsx' }]),
  ['a.xlsx', 'a (2).xlsx', 'a (3).xlsx']);

check('the counter goes before the extension, not after',
  names([{ filename: 'x.y.xlsx' }, { filename: 'x.y.xlsx' }])[1], 'x.y (2).xlsx');

/* --- who may see the log spreadsheet link ------------------------------- */
const realGetUserEmail = getUserEmail_;

getUserEmail_ = () => 'b.gheorghe@secretescapes.com';
check('owner gets the log link', (getLogLink() || {}).url,
  'https://docs.google.com/spreadsheets/d/1Iw4_A5SnVXxgxkTVrVkkcUWGOMyqMo0cSflhEvnumkw/edit');

getUserEmail_ = () => 'B.Gheorghe@SecretEscapes.com';
check('owner match ignores capitalisation', !!getLogLink(), true);

getUserEmail_ = () => '  b.gheorghe@secretescapes.com  ';
check('owner match ignores stray spaces', !!getLogLink(), true);

getUserEmail_ = () => 'someone.else@secretescapes.com';
check('colleague gets nothing', getLogLink(), null);

getUserEmail_ = () => 'unknown';
check('unknown user gets nothing', getLogLink(), null);

check('the log link is not handed to every visitor',
  Object.keys(getBootstrapShape()).indexOf('logUrl'), -1);

getUserEmail_ = realGetUserEmail;

function getBootstrapShape() {
  // getBootstrap() itself needs Apps Script globals; read its literal instead.
  const source = fs.readFileSync(__dirname + '/Code.gs', 'utf8');
  const body = source.slice(source.indexOf('function getBootstrap()'));
  const keys = {};
  body.slice(0, body.indexOf('}')).replace(/(\w+):/g, (_, key) => { keys[key] = true; });
  return keys;
}

console.log(failures ? `\n${failures} FAILED` : '\nAll checks passed.');
process.exit(failures ? 1 : 0);
