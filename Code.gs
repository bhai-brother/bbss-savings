/**
 * BBSS Google Sheets backend
 * Deploy this bound Apps Script as a Web App.
 * The website stays on GitHub Pages; this script stores the shared state in Google Sheets.
 */

const BBSS = {
  STATE_SHEET: 'State',
  SESSION_SHEET: 'Sessions',
  STATE_KEY: 'app_state',
  SESSION_HOURS_REMEMBER: 24 * 30,
  SESSION_HOURS_SHORT: 12,
  SHEETS: {
    README: ['Sheet', 'Purpose'],
    Meta: ['Key', 'Value'],
    Members: ['MemberID', 'Name', 'AliasesJSON', 'Active', 'MustChangePassword', 'Mobile', 'Blood', 'Email', 'Address', 'CreatedAt'],
    Deposits: ['Year', 'LedgerID', 'SL', 'MemberID', 'Name', 'January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December', 'Total'],
    Funds: ['Year', 'FundID', 'Label', 'Amount'],
    AnnualReview: ['Year', 'Meeting', 'Previous', 'Deposits', 'Profit', 'Loss', 'GrandTotal'],
    Investments: ['ID', 'Group', 'SL', 'Name', 'Date', 'Description', 'Principal', 'Profit', 'ProfitDate', 'Paid'],
    PersonalInfo: ['ID', 'Name', 'Mobile', 'Blood'],
    Transactions: ['ID', 'Year', 'MemberAccountID', 'MemberID', 'MemberName', 'Month', 'Amount', 'FundID', 'FundLabel', 'At', 'CreatedBy'],
    ProfileRequests: ['ID', 'MemberID', 'SubmittedAt', 'Status', 'ChangesJSON', 'PhotoIncluded', 'ReviewedAt'],
    PasswordResetRequests: ['ID', 'MemberID', 'Name', 'RequestedAt', 'Status', 'ReviewedAt'],
    Notes: ['ID', 'Text'],
    Audit: ['At', 'Action'],
    State: ['Key', 'ChunkIndex', 'JSONChunk', 'UpdatedAt'],
    Sessions: ['TokenHash', 'Role', 'MemberID', 'ExpiresAt', 'CreatedAt', 'LastSeen']
  }
};

function setupBBSS() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open the Google Sheet first, then run setupBBSS().');
  const props = PropertiesService.getScriptProperties();
  props.setProperty('BBSS_SHEET_ID', ss.getId());

  Object.keys(BBSS.SHEETS).forEach(name => ensureSheet_(ss, name, BBSS.SHEETS[name]));
  writeReadme_(ss);

  let key = props.getProperty('BBSS_WRITE_KEY');
  if (!key) {
    key = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 16);
    props.setProperty('BBSS_WRITE_KEY', key);
  }

  Logger.log('BBSS WRITE KEY: ' + key);
  Logger.log('Copy this key and keep it private. Enter it only in the Admin browser.');
  return 'Setup complete. WRITE KEY: ' + key;
}

function resetWriteKey() {
  const key = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 16);
  PropertiesService.getScriptProperties().setProperty('BBSS_WRITE_KEY', key);
  Logger.log('NEW BBSS WRITE KEY: ' + key);
  return key;
}

function doGet(e) {
  try {
    const action = String((e && e.parameter && e.parameter.action) || 'health');
    if (action !== 'health') return json_({ok:false, error:'Use POST for this action.'});
    return json_({
      ok: true,
      service: 'BBSS Google Sheets API',
      initialized: !!loadState_(),
      time: new Date().toISOString()
    });
  } catch (err) {
    return json_({ok:false, error:String(err && err.message || err)});
  }
}

function doPost(e) {
  try {
    const body = parseBody_(e);
    const action = String(body.action || '');

    if (action === 'health') return json_({ok:true, initialized:!!loadState_()});
    if (action === 'bootstrapSave') return json_(bootstrapSave_(body));
    if (action === 'login') return json_(login_(body));
    if (action === 'session') return json_(sessionLoad_(body));
    if (action === 'logout') return json_(logout_(body));
    if (action === 'save') return json_(adminSave_(body));
    if (action === 'requestReset') return json_(requestReset_(body));
    if (action === 'profileRequest') return json_(profileRequest_(body));
    if (action === 'changePassword') return json_(changePassword_(body));

    return json_({ok:false, error:'Unknown action'});
  } catch (err) {
    return json_({ok:false, error:String(err && err.message || err)});
  }
}

function bootstrapSave_(body) {
  requireWriteKey_(body.writeKey);
  validateState_(body.state);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    saveState_(body.state);
  } finally {
    lock.releaseLock();
  }
  return {ok:true, initialized:true, message:'Initial/shared state saved'};
}

function login_(body) {
  const state = loadState_();
  if (!state) return {ok:false, code:'not_initialized', error:'Database is not initialized yet. Admin must bootstrap-save the current BBSS state first.'};
  const role = String(body.role || 'member');
  const password = String(body.password || '');

  if (role === 'admin') {
    if (password !== String((state.meta && state.meta.adminPin) || '1234')) return {ok:false, error:'Admin password/PIN সঠিক নয়।'};
    const token = createSession_('admin', '', !!body.remember);
    return {ok:true, role:'admin', token:token, state:sanitizeState_(state, {role:'admin', memberId:''})};
  }

  const account = resolveMember_(state, body.login);
  if (account.error) return {ok:false, error:account.error};
  const a = account.member;
  if (!a || a.active === false) return {ok:false, error:'এই Member ID বর্তমানে নিষ্ক্রিয়।'};
  if (sha256_(password) !== String(a.passwordHash || '')) return {ok:false, error:'পাসওয়ার্ড সঠিক নয়।'};

  const token = createSession_('member', a.memberId, !!body.remember);
  return {ok:true, role:'member', memberId:a.memberId, token:token, state:sanitizeState_(state, {role:'member', memberId:a.memberId})};
}

function sessionLoad_(body) {
  const sess = validateSession_(body.token);
  if (!sess) return {ok:false, code:'session_expired', error:'Session expired'};
  let state = loadState_();
  if (!state) return {ok:false, code:'not_initialized', error:'Database not initialized'};
  if (recalculateAnnualFromDeposits_(state)) {
    // Re-load under the script lock before self-healing, so a concurrent Admin save is never overwritten.
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const latest = loadState_();
      if (latest) {
        if (recalculateAnnualFromDeposits_(latest)) {
          latest.meta = latest.meta || {};
          latest.meta.lastUpdated = new Date().toISOString();
          latest.audit = Array.isArray(latest.audit) ? latest.audit : [];
          latest.audit.unshift({at:latest.meta.lastUpdated, action:'Annual totals auto-recalculated from live deposits'});
          latest.audit = latest.audit.slice(0, 150);
          saveState_(latest);
        }
        state = latest;
      }
    } finally {
      lock.releaseLock();
    }
  }
  touchSession_(sess.row);
  return {ok:true, role:sess.role, memberId:sess.memberId, state:sanitizeState_(state, sess)};
}

function logout_(body) {
  deleteSession_(body.token);
  return {ok:true};
}

function adminSave_(body) {
  const sess = validateSession_(body.token);
  if (!sess || sess.role !== 'admin') return {ok:false, error:'Admin session required'};
  validateState_(body.state);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const current = loadState_();
    const incoming = body.state;
    // Preserve member password hashes if a future client intentionally omits them.
    if (current && current.auth && Array.isArray(current.auth.members) && incoming.auth && Array.isArray(incoming.auth.members)) {
      const oldMap = {};
      current.auth.members.forEach(m => oldMap[String(m.memberId || '')] = m);
      incoming.auth.members.forEach(m => {
        const old = oldMap[String(m.memberId || '')];
        if (old && !m.passwordHash) m.passwordHash = old.passwordHash;
      });
    }
    saveState_(incoming);
  } finally {
    lock.releaseLock();
  }
  return {ok:true, state:sanitizeState_(loadState_(), sess)};
}

function requestReset_(body) {
  const state = loadState_();
  if (!state) return {ok:false, error:'Database not initialized'};
  const account = resolveMember_(state, body.login);
  if (account.error) return {ok:false, error:account.error};
  const a = account.member;
  state.auth = state.auth || {};
  state.auth.passwordResetRequests = Array.isArray(state.auth.passwordResetRequests) ? state.auth.passwordResetRequests : [];
  const now = new Date().toISOString();
  const old = state.auth.passwordResetRequests.find(r => r.memberId === a.memberId && r.status === 'pending');
  if (old) old.requestedAt = now;
  else state.auth.passwordResetRequests.unshift({id:'pr'+Date.now(), memberId:a.memberId, name:a.name, requestedAt:now, status:'pending'});
  saveState_(state);
  return {ok:true, message:'Reset request sent'};
}

function profileRequest_(body) {
  const sess = validateSession_(body.token);
  if (!sess || sess.role !== 'member') return {ok:false, error:'Member session required'};
  const state = loadState_();
  const a = findMemberById_(state, sess.memberId);
  if (!a) return {ok:false, error:'Member not found'};
  state.auth = state.auth || {};
  state.auth.profileRequests = Array.isArray(state.auth.profileRequests) ? state.auth.profileRequests : [];
  const now = new Date().toISOString();
  let req = state.auth.profileRequests.find(r => r.memberId === a.memberId && r.status === 'pending');
  const changes = body.changes || {};
  const photoData = String(body.photoData || '');
  if (photoData.length > 2200000) return {ok:false, error:'ছবির size বেশি। 1.5MB-এর মধ্যে রাখুন।'};
  if (req) {
    req.changes = changes;
    if (photoData) req.photoData = photoData;
    req.submittedAt = now;
  } else {
    state.auth.profileRequests.unshift({id:'pf'+Date.now(), memberId:a.memberId, submittedAt:now, status:'pending', changes:changes, photoData:photoData});
  }
  saveState_(state);
  return {ok:true, state:sanitizeState_(state, sess)};
}

function changePassword_(body) {
  const sess = validateSession_(body.token);
  if (!sess || sess.role !== 'member') return {ok:false, error:'Member session required'};
  const state = loadState_();
  const a = findMemberById_(state, sess.memberId);
  if (!a) return {ok:false, error:'Member not found'};
  if (sha256_(String(body.oldPassword || '')) !== String(a.passwordHash || '')) return {ok:false, error:'বর্তমান password সঠিক নয়'};
  const next = String(body.newPassword || '');
  if (next.length < 4) return {ok:false, error:'নতুন password কমপক্ষে 4 অক্ষর দিন'};
  a.passwordHash = sha256_(next);
  a.mustChangePassword = false;
  saveState_(state);
  return {ok:true, state:sanitizeState_(state, sess)};
}

function sanitizeState_(state, sess) {
  const s = JSON.parse(JSON.stringify(state));
  if (sess.role === 'admin') return s;

  const memberId = String(sess.memberId || '');
  const current = findMemberById_(s, memberId);
  if (s.meta) {
    delete s.meta.adminPin;
    s.meta.autoSync = false;
  }
  s.investments = [];
  s.notes = [];
  s.audit = [];
  s.personalInfo = (s.personalInfo || []).filter(p => current && baseName_(p.name) === baseName_(current.name));
  s.depositTransactions = (s.depositTransactions || []).filter(t => String(t.memberId || '') === memberId);

  s.auth = s.auth || {};
  s.auth.members = (s.auth.members || []).map(m => {
    const own = String(m.memberId || '') === memberId;
    return {
      memberId:m.memberId,
      name:m.name,
      aliases:Array.isArray(m.aliases) ? m.aliases : [m.name],
      active:m.active !== false,
      mustChangePassword:own ? !!m.mustChangePassword : false,
      photoData:own ? String(m.photoData || '') : '',
      profile:own ? (m.profile || {}) : {},
      createdAt:m.createdAt || ''
    };
  });
  s.auth.profileRequests = (s.auth.profileRequests || []).filter(r => String(r.memberId || '') === memberId);
  s.auth.passwordResetRequests = [];
  return s;
}

function resolveMember_(state, value) {
  const raw = String(value || '').trim();
  const id = raw.toLowerCase();
  const members = (state.auth && Array.isArray(state.auth.members)) ? state.auth.members : [];
  const byId = members.find(a => String(a.memberId || '').toLowerCase() === id);
  if (byId) return {member:byId};

  const mobile = normalizeMobile_(raw);
  if (!mobile) return {error:'Member ID বা মোবাইল নম্বর পাওয়া যায়নি।'};
  const list = members.filter(a => normalizeMobile_(memberMobile_(state, a)) === mobile);
  if (list.length === 1) return {member:list[0]};
  if (list.length > 1) return {error:'এই মোবাইল নম্বরে একাধিক Member ID যুক্ত আছে। নিরাপত্তার জন্য Member ID দিয়ে লগইন করুন।'};
  return {error:'Member ID বা মোবাইল নম্বর পাওয়া যায়নি।'};
}

function memberMobile_(state, a) {
  if (a && a.profile && a.profile.mobile) return a.profile.mobile;
  const name = baseName_(a && a.name);
  const p = (state.personalInfo || []).find(x => baseName_(x.name) === name);
  return p ? p.mobile : '';
}

function findMemberById_(state, memberId) {
  return ((state.auth && state.auth.members) || []).find(a => String(a.memberId || '') === String(memberId || '')) || null;
}

function normalizeMobile_(v) {
  let x = String(v || '').replace(/\D/g, '');
  if (x.indexOf('880') === 0 && x.length === 13) x = '0' + x.slice(3);
  return x;
}

function baseName_(v) {
  return String(v || '').replace(/-(01|02)$/i, '').trim().toLowerCase();
}

function createSession_(role, memberId, remember) {
  const ss = getSS_();
  const sh = ensureSheet_(ss, BBSS.SESSION_SHEET, BBSS.SHEETS.Sessions);
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  const now = new Date();
  const hours = remember ? BBSS.SESSION_HOURS_REMEMBER : BBSS.SESSION_HOURS_SHORT;
  const exp = new Date(now.getTime() + hours * 3600000);
  sh.appendRow([sha256_(token), role, memberId || '', exp.toISOString(), now.toISOString(), now.toISOString()]);
  pruneSessions_();
  return token;
}

function validateSession_(token) {
  const t = String(token || '');
  if (!t) return null;
  const sh = getSS_().getSheetByName(BBSS.SESSION_SHEET);
  if (!sh || sh.getLastRow() < 2) return null;
  const vals = sh.getRange(2, 1, sh.getLastRow()-1, 6).getValues();
  const h = sha256_(t);
  const now = Date.now();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]) !== h) continue;
    const exp = new Date(vals[i][3]).getTime();
    if (!exp || exp < now) {
      sh.deleteRow(i + 2);
      return null;
    }
    return {row:i+2, role:String(vals[i][1]), memberId:String(vals[i][2] || '')};
  }
  return null;
}

function touchSession_(row) {
  try { getSS_().getSheetByName(BBSS.SESSION_SHEET).getRange(row, 6).setValue(new Date().toISOString()); } catch (e) {}
}

function deleteSession_(token) {
  const sh = getSS_().getSheetByName(BBSS.SESSION_SHEET);
  if (!sh || sh.getLastRow() < 2) return;
  const h = sha256_(String(token || ''));
  const vals = sh.getRange(2, 1, sh.getLastRow()-1, 1).getValues();
  for (let i = vals.length - 1; i >= 0; i--) if (String(vals[i][0]) === h) sh.deleteRow(i + 2);
}

function pruneSessions_() {
  const sh = getSS_().getSheetByName(BBSS.SESSION_SHEET);
  if (!sh || sh.getLastRow() < 2) return;
  const vals = sh.getRange(2, 1, sh.getLastRow()-1, 6).getValues();
  const now = Date.now();
  for (let i = vals.length - 1; i >= 0; i--) {
    const exp = new Date(vals[i][3]).getTime();
    if (!exp || exp < now) sh.deleteRow(i + 2);
  }
}

function requireWriteKey_(key) {
  const expected = PropertiesService.getScriptProperties().getProperty('BBSS_WRITE_KEY');
  if (!expected) throw new Error('Run setupBBSS() first.');
  if (String(key || '') !== String(expected)) throw new Error('Write key is incorrect');
}

function recalculateAnnualFromDeposits_(state) {
  if (!state || !Array.isArray(state.annualReviews)) return false;
  const before = JSON.stringify(state.annualReviews);
  const totals = {};
  Object.keys(state.deposits || {}).forEach(year => {
    totals[String(year)] = (state.deposits[year] || []).reduce((sum, row) => {
      return sum + (Array.isArray(row.months) ? row.months : []).reduce((s, v) => s + (Number(v) || 0), 0);
    }, 0);
  });
  const rows = state.annualReviews.slice().sort((a,b) => Number(a.year) - Number(b.year));
  let carry = 0;
  rows.forEach((x, i) => {
    if (i > 0) x.previous = carry;
    else x.previous = Number(x.previous) || 0;
    x.deposits = Number(totals[String(x.year)] || 0);
    x.grandTotal = (Number(x.previous)||0) + x.deposits + (Number(x.profit)||0) - (Number(x.loss)||0);
    carry = x.grandTotal;
  });
  return before !== JSON.stringify(state.annualReviews);
}

function saveState_(state) {
  validateState_(state);
  recalculateAnnualFromDeposits_(state);
  const ss = getSS_();
  const sh = ensureSheet_(ss, BBSS.STATE_SHEET, BBSS.SHEETS.State);
  const json = JSON.stringify(state);
  const chunkSize = 35000; // safely below Google Sheets' per-cell text limit
  const now = new Date().toISOString();
  const rows = [];
  for (let i = 0, idx = 1; i < json.length; i += chunkSize, idx++) rows.push([BBSS.STATE_KEY, idx, json.slice(i, i + chunkSize), now]);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow()-1, 4).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, 4).setValues(rows);
  syncMirrorSheets_(ss, state);
}

function loadState_() {
  const ss = getSS_(true);
  if (!ss) return null;
  const sh = ss.getSheetByName(BBSS.STATE_SHEET);
  if (!sh || sh.getLastRow() < 2) return null;
  const vals = sh.getRange(2, 1, sh.getLastRow()-1, 4).getValues()
    .filter(r => String(r[0]) === BBSS.STATE_KEY && r[2] !== '')
    .sort((a,b) => Number(a[1]) - Number(b[1]));
  if (!vals.length) return null;
  const json = vals.map(r => String(r[2] || '')).join('');
  try { return JSON.parse(json); } catch (e) { throw new Error('State JSON is corrupted'); }
}

function syncMirrorSheets_(ss, state) {
  writeTable_(ss, 'Meta', BBSS.SHEETS.Meta, Object.keys(state.meta || {}).map(k => [k, serializeCell_(state.meta[k])]));

  const authMembers = ((state.auth && state.auth.members) || []);
  const members = authMembers.map(a => [
    a.memberId || '', a.name || '', JSON.stringify(a.aliases || []), a.active !== false, !!a.mustChangePassword,
    (a.profile && a.profile.mobile) || '', (a.profile && a.profile.blood) || '', (a.profile && a.profile.email) || '',
    (a.profile && a.profile.address) || '', a.createdAt || ''
  ]);
  writeTable_(ss, 'Members', BBSS.SHEETS.Members, members);

  const memberIdForName = name => {
    const n = String(name || '').trim().toLowerCase();
    const a = authMembers.find(m => String(m.name || '').trim().toLowerCase() === n || (m.aliases || []).some(x => String(x || '').trim().toLowerCase() === n));
    return a ? a.memberId : '';
  };

  const depRows = [];
  Object.keys(state.deposits || {}).sort().forEach(year => {
    (state.deposits[year] || []).forEach(r => {
      const months = Array.isArray(r.months) ? r.months.slice(0,12) : [];
      while (months.length < 12) months.push(0);
      depRows.push([year, r.id || '', r.sl || '', memberIdForName(r.name), r.name || ''].concat(months).concat([months.reduce((a,b)=>a+(Number(b)||0),0)]));
    });
  });
  writeTable_(ss, 'Deposits', BBSS.SHEETS.Deposits, depRows);

  const fundRows = [];
  Object.keys(state.funds || {}).sort().forEach(year => (state.funds[year] || []).forEach(f => fundRows.push([year, f.id || '', f.label || '', Number(f.amount)||0])));
  writeTable_(ss, 'Funds', BBSS.SHEETS.Funds, fundRows);

  writeTable_(ss, 'AnnualReview', BBSS.SHEETS.AnnualReview, (state.annualReviews || []).map(x => [x.year||'', x.meeting||'', Number(x.previous)||0, Number(x.deposits)||0, Number(x.profit)||0, Number(x.loss)||0, Number(x.grandTotal)||0]));
  writeTable_(ss, 'Investments', BBSS.SHEETS.Investments, (state.investments || []).map(x => [x.id||'', x.group||'', x.sl||'', x.name||'', x.date||'', x.description||'', Number(x.amount)||0, Number(x.profit)||0, x.profitDate||'', x.paid||'']));
  writeTable_(ss, 'PersonalInfo', BBSS.SHEETS.PersonalInfo, (state.personalInfo || []).map(x => [x.id||'', x.name||'', x.mobile||'', x.blood||'']));
  writeTable_(ss, 'Transactions', BBSS.SHEETS.Transactions, (state.depositTransactions || []).map(t => [t.id||'', t.year||'', t.memberAccountId||'', t.memberId||'', t.memberName||'', t.month??'', Number(t.amount)||0, t.fundId||'', t.fundLabel||'', t.at||'', t.createdBy||'']));
  writeTable_(ss, 'ProfileRequests', BBSS.SHEETS.ProfileRequests, (((state.auth||{}).profileRequests)||[]).map(r => [r.id||'', r.memberId||'', r.submittedAt||'', r.status||'', JSON.stringify(r.changes||{}), r.photoData?'YES':'', r.reviewedAt||'']));
  writeTable_(ss, 'PasswordResetRequests', BBSS.SHEETS.PasswordResetRequests, (((state.auth||{}).passwordResetRequests)||[]).map(r => [r.id||'', r.memberId||'', r.name||'', r.requestedAt||'', r.status||'', r.reviewedAt||'']));
  writeTable_(ss, 'Notes', BBSS.SHEETS.Notes, (state.notes || []).map(n => [n.id||'', n.text||'']));
  writeTable_(ss, 'Audit', BBSS.SHEETS.Audit, (state.audit || []).map(a => [a.at||'', a.action||'']));
}

function writeReadme_(ss) {
  const rows = [
    ['State', 'Canonical JSON used by the BBSS website, split into safe-size chunks. Do not manually edit these rows.'],
    ['Meta', 'Organization/app settings mirror.'],
    ['Members', 'Member directory mirror. Password hashes are intentionally not mirrored here.'],
    ['Deposits', 'Year-wise monthly deposit mirror.'],
    ['Funds', 'Fund/cash ledger mirror.'],
    ['AnnualReview', 'Annual review mirror.'],
    ['Investments', 'Investment records mirror.'],
    ['PersonalInfo', 'Member personal info mirror.'],
    ['Transactions', 'Deposit transaction log mirror.'],
    ['ProfileRequests', 'Member profile change requests mirror.'],
    ['PasswordResetRequests', 'Password reset requests mirror.'],
    ['Notes', 'Notes mirror.'],
    ['Audit', 'Audit log mirror.'],
    ['Sessions', 'Hashed login session tokens. Do not edit manually.']
  ];
  writeTable_(ss, 'README', BBSS.SHEETS.README, rows);
}

function writeTable_(ss, name, headers, rows) {
  const sh = ensureSheet_(ss, name, headers);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow()-1, Math.max(sh.getLastColumn(), headers.length)).clearContent();
  if (rows && rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows.map(r => {
    const a = r.slice(0, headers.length);
    while (a.length < headers.length) a.push('');
    return a;
  }));
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, headers.length);
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getMaxColumns() < headers.length) sh.insertColumnsAfter(sh.getMaxColumns(), headers.length - sh.getMaxColumns());
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sh.setFrozenRows(1);
  return sh;
}

function getSS_(allowMissing) {
  const id = PropertiesService.getScriptProperties().getProperty('BBSS_SHEET_ID');
  if (!id) {
    if (allowMissing) return null;
    throw new Error('Run setupBBSS() from the bound Google Sheet first.');
  }
  return SpreadsheetApp.openById(id);
}

function parseBody_(e) {
  const text = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
  try { return JSON.parse(text); } catch (err) { throw new Error('Invalid JSON body'); }
}

function validateState_(state) {
  if (!state || typeof state !== 'object' || !state.meta || !state.deposits) throw new Error('Invalid BBSS state');
}

function serializeCell_(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'object') return JSON.stringify(v);
  return v;
}

function sha256_(value) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(value || ''), Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + ((b < 0 ? b + 256 : b).toString(16))).slice(-2)).join('');
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
