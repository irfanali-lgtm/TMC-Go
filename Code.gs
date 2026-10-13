/**
 * TMC Visit Intelligence — Google Apps Script backend
 * ---------------------------------------------------
 * Sheet tabs (created by setup() if missing):
 *   AEs            name | city | active | coverageMeasured | notes | title | phone | email | photo   (profile fields are set by the AE in the app; photo = small data URL)
 *   Accounts       id | ae | account | tier | city | netNew | addedOn | notes | address | lat | lng | located
 *   Visits         id | date | ae | city | accounts | count | km | purpose | loggedBy | ts | route | kmMode | details | note
 *                    (note = manager's note on that day, set from the dashboard with the admin PIN)
 *                    (details = JSON list of {client, purpose} per visit, written by the app)
 *   Opportunities  id | ae | client | opportunity | stage | updated
 *   Aliases        alias | account | ae        (spellings AEs used → the account they meant; learned by the app)
 *   Categories     AE | Account | Category | City   (MANAGER-MAINTAINED: who owns which account and its A/B/C;
 *                  headers may also be AM / Location / Tier; "-" or blank = no category; rows whose Account starts
 *                  with "Total" are ignored; AE names are matched to the AEs tab even with spelling differences
 *                  ("Abdullah" + Karachi → Muhammad Abdullah, "Mudassir S" → Mudassar Shahid); "TBH"/blank = unassigned.
 *                  Synced into Accounts on every load — it is the source of truth for tiers.)
 *   Settings       key | value
 *                    adminPin       second PIN for managers: remove/restore AEs, office addresses, purpose list
 *                    mapsKey        optional Google Maps Platform key (Geocoding + Routes API). With it, routes use
 *                                   live traffic and the Platform quota; without it the built-in Apps Script Maps
 *                                   service is used (no key, Workspace daily quota, no traffic).
 *                    autoKm         Y = compute km from Google Maps when a visit is saved
 *                    purposes       extra purpose dropdown entries, separated by ;
 *                    aiProvider     gemini (the only provider supported)
 *                    aiKey          API key for that provider — stays in the sheet, never sent to the browser
 *                    aiModel        optional model override (default: gemini-3.5-flash)
 *                    office.Lahore  street address of the Lahore office (also office.Karachi, office.Islamabad …);
 *                                   setup() fills in TMC's three offices, change them in ⚙ in the app
 *
 * Kilometres: when a visit is saved the script geocodes each account (cached in the Accounts tab),
 * orders them nearest-first from the AE's city office, asks Google Maps for the driving route
 * office → … → office, and stores the total km plus the leg-by-leg route on the visit row.
 * This uses Apps Script's built-in Maps service — no API key needed (daily quota applies).
 *
 * SIGN-IN: every person signs in as themselves (email code, a manager's one-time code, or manager + admin PIN).
 *   Sessions live in the Sessions tab; the role column on the AEs tab marks managers. Set a strong adminPin first.
 * SETUP:
 *   1. Google Sheets → File → Import → upload the seed .xlsx → "Replace spreadsheet".
 *   2. Extensions → Apps Script → paste this file over Code.gs. Save. (No HTML files needed.)
 *   3. Deploy → New deployment → Web app · Execute as: Me · Who has access: Anyone → copy the /exec link.
 *   4. Put index.html online (Netlify) and paste the /exec link into the app when it asks.
 *   Tabs, columns and default settings are created automatically on first use.
 */

var TZ = 'Asia/Karachi'; // "today" for the team, whatever the script's own time zone is
var SETUP_VERSION = '8';  // bump when tabs/columns/settings change; setup runs itself on the next request

var TABS = {
  AEs:           ['name','city','active','coverageMeasured','notes','title','phone','email','photo','role'],
  Sessions:      ['token','ae','role','created','lastSeen','device','method'],
  Accounts:      ['id','ae','account','tier','city','netNew','addedOn','notes','address','lat','lng','located'],
  Visits:        ['id','date','ae','city','accounts','count','km','purpose','loggedBy','ts','route','kmMode','details','note'],
  Opportunities: ['id','ae','client','opportunity','stage','updated'],
  Aliases:       ['alias','account','ae'],
  Settings:      ['key','value']
};
// Manager-maintained category list. Optional, simple columns; synced into Accounts on every load.
var CATEGORIES_TAB = 'Categories', CATEGORIES_COLS = ['AE', 'Account', 'Category', 'City'];
var HEADER_ALIASES = { ae: ['ae', 'am', 'account manager', 'owner', 'name'], account: ['account', 'client', 'company'], category: ['category', 'tier', 'cat'], city: ['city', 'location', 'base'] };
function findCol_(hdr, key) { var opts = HEADER_ALIASES[key]; for (var i = 0; i < hdr.length; i++) if (opts.indexOf(String(hdr[i]).trim().toLowerCase()) !== -1) return i; return -1; }

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(TABS).forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(TABS[name]);
      sh.getRange(1, 1, 1, TABS[name].length).setFontWeight('bold');
      sh.setFrozenRows(1);
      ['date', 'addedOn', 'updated'].forEach(function (c) { var i = TABS[name].indexOf(c); if (i !== -1) sh.getRange(2, i + 1, Math.max(sh.getMaxRows() - 1, 1), 1).setNumberFormat('@'); });
    } else {
      // add any columns introduced after the sheet was created
      var have = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(String);
      TABS[name].forEach(function (c) {
        if (have.indexOf(c) === -1) { sh.getRange(1, have.length + 1).setValue(c).setFontWeight('bold'); have.push(c); }
      });
      // dates are stored as plain text (yyyy-mm-dd) so a time-zone setting can never move a visit to another day
      ['date', 'addedOn', 'updated'].forEach(function (c) { var i = have.indexOf(c); if (i !== -1) sh.getRange(2, i + 1, Math.max(Math.min(sh.getMaxRows() - 1, 2000), 1), 1).setNumberFormat('@'); });
    }
  });
  var cat = ss.getSheetByName(CATEGORIES_TAB);
  if (!cat) { cat = ss.insertSheet(CATEGORIES_TAB); cat.appendRow(CATEGORIES_COLS); cat.getRange(1, 1, 1, 4).setFontWeight('bold'); cat.setFrozenRows(1); }
  var st = ss.getSheetByName('Settings');
  var keys = {};
  readTab_('Settings').forEach(function (r) { keys[String(r.key)] = true; });
  [['adminPin', '0000'], ['autoKm', 'Y'], ['mapsKey', ''], ['aiProvider', 'gemini'], ['aiKey', ''], ['aiModel', ''],
   ['office.Lahore', "Floor No. 10, Zone B, DHA Business Hub, D.H.A. Commercial Broadway, DHA Phase 8, Lahore, Pakistan"],
   ['office.Karachi', "1st Floor, Nasreen Building, Main Shaheed-e-Millat Road, Block 7/8 Bangalore Town, Karachi, Pakistan"],
   ['office.Islamabad', "4th Floor, SNC Center, AKM Fazl-ul-Haq Road, Block D G-6/2, Blue Area, Islamabad, Pakistan"]].forEach(function (kv) {
    if (!keys[kv[0]]) st.appendRow(kv);
  });
}

/* ---------- helpers ---------- */
// Runs setup() once per SETUP_VERSION, automatically, so nobody has to click "Run" in the editor.
function ensureSetup_() {
  try {
    var props = PropertiesService.getScriptProperties();
    if (props.getProperty('setupVersion') === SETUP_VERSION) return;
    setup();
    props.setProperty('setupVersion', SETUP_VERSION);
  } catch (e) { /* a sheet still opening; the next request retries */ }
}
function sheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) { setup(); sh = ss.getSheetByName(name); }
  return sh;
}
function headers_(sh) {
  var n = Math.max(sh.getLastColumn(), 1);
  return sh.getRange(1, 1, 1, n).getValues()[0].map(String);
}
function readTab_(name) {
  var sh = sheet_(name), cols = TABS[name], hdr = headers_(sh);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, hdr.length).getValues();
  var out = [];
  vals.forEach(function (r) {
    if (r.every(function (c) { return c === '' || c === null; })) return;
    var o = {};
    cols.forEach(function (c) {
      var i = hdr.indexOf(c), v = i === -1 ? '' : r[i];
      if (v instanceof Date) v = Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
      o[c] = v === null ? '' : v;
    });
    out.push(o);
  });
  return out;
}
function settings_() {
  var o = {};
  readTab_('Settings').forEach(function (r) { o[String(r.key)] = String(r.value); });
  return o;
}
function checkPin_(pin) { /* normal use needs no PIN; manager actions check adminPin below */ }
function checkAdmin_(adminPin, sess) {
  if (sess && sess.role === 'manager') return;                     // signed-in managers need no PIN
  if (weakAdminPin_()) throw new Error('WEAKPIN: the admin PIN is still the default or too simple');   // 0000 unlocks nothing
  if (sess && !adminPin) throw new Error('MANAGER_ONLY: only managers can do that');
  if (sess && tooMany_('pin:' + sess.ae)) throw new Error('Too many tries — wait 15 minutes.');
  if (sess && String(adminPin || '') !== String(settings_().adminPin || '')) { failed_('pin:' + sess.ae); }
  var want = settings_().adminPin || '';
  if (!want) throw new Error('No adminPin set in the Settings tab');
  if (String(adminPin || '') !== want) throw new Error('Wrong admin PIN');
}

/* ---------- Sign-in: every person is themselves ----------
 * Each device holds a session token, issued only after the person proves who they are:
 *   • a 6-digit code emailed to the address on their AE row, or
 *   • a one-time code a manager generated for them (Settings → Team), or
 *   • for managers: their name + the admin PIN.
 * Every write is then checked against the session's own name, so nobody can log visits as someone else.
 */
var SESSION_DAYS = 60;
function aeRow_(name) {
  var a = readTab_('AEs').filter(function (x) { return String(x.name) === String(name) && String(x.active || 'Y').toUpperCase() !== 'N'; })[0];
  if (!a) throw new Error('That name is not on the team list.');
  return a;
}
function roleOf_(a) { return String(a.role || '').toLowerCase() === 'manager' ? 'manager' : 'ae'; }
function maskEmail_(e) { e = String(e || ''); var m = /^(.)([^@]*)@(.*)$/.exec(e); return m ? m[1] + m[2].replace(/./g, '•').slice(0, 6) + '@' + m[3] : ''; }
function newToken_() { return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 16); }
function issueSession_(ae, role, device, method) {
  try { purgeSessions_(); } catch (e) {}
  var tok = newToken_(), now = new Date().toISOString();
  var sh = sheet_('Sessions'), hdr = headers_(sh), row = { token: tok, ae: ae, role: role, created: now, lastSeen: now, device: String(device || '').slice(0, 120), method: method };
  sh.appendRow(hdr.map(function (c) { return row[c] === undefined ? '' : row[c]; }));
  CacheService.getScriptCache().put('sess:' + tok, JSON.stringify({ ae: ae, role: role }), 21600);
  return { token: tok, ae: ae, role: role };
}
function session_(token) {
  token = String(token || ''); if (token.length < 20) return null;
  var cache = CacheService.getScriptCache(), hit = cache.get('sess:' + token);
  if (hit) return JSON.parse(hit);
  var rows = readTab_('Sessions').filter(function (r) { return r.token === token; });
  if (!rows.length) return null;
  var r = rows[0];
  if (new Date().getTime() - new Date(r.created).getTime() > SESSION_DAYS * 864e5) return null;
  var a; try { a = aeRow_(r.ae); } catch (e) { return null; }          // removed AEs lose access at once
  var s = { ae: r.ae, role: roleOf_(a) };
  cache.put('sess:' + token, JSON.stringify(s), 21600);
  return s;
}
function requireSession_(b) {
  var s = session_(b.token);
  if (!s) throw new Error('SIGNIN: please sign in');
  return s;
}
function purgeSessions_() {
  var sh = sheet_('Sessions'), last = sh.getLastRow(); if (last < 2) return;
  var hdr = headers_(sh), ci = hdr.indexOf('created'), vals = sh.getRange(2, ci + 1, last - 1, 1).getValues(), cut = new Date().getTime() - SESSION_DAYS * 864e5;
  for (var i = vals.length - 1; i >= 0; i--) { var t = new Date(vals[i][0]).getTime(); if (t && t < cut) sh.deleteRow(i + 2); }
}
function revokeSessions_(ae, exceptToken) {
  var sh = sheet_('Sessions'), last = sh.getLastRow(); if (last < 2) return 0;
  var hdr = headers_(sh), ti = hdr.indexOf('token'), ai = hdr.indexOf('ae'), vals = sh.getRange(2, 1, last - 1, hdr.length).getValues(), n = 0, cache = CacheService.getScriptCache();
  for (var i = vals.length - 1; i >= 0; i--) if (vals[i][ai] === ae && vals[i][ti] !== exceptToken) { cache.remove('sess:' + vals[i][ti]); sh.deleteRow(i + 2); n++; }
  return n;
}
function authAction_(b) {
  var cache = CacheService.getScriptCache(), props = PropertiesService.getScriptProperties();
  switch (b.action) {
    case 'roster':   // names for the sign-in screen; nothing private
      return { status: 'ok', roster: readTab_('AEs').filter(function (a) { return String(a.active || 'Y').toUpperCase() !== 'N'; }).map(function (a) { return { name: a.name, city: a.city, email: maskEmail_(a.email), photo: a.photo || '', manager: roleOf_(a) === 'manager' }; }) };
    case 'sendCode': {
      var a = aeRow_(b.name);
      if (!/@/.test(String(a.email || ''))) throw new Error('No email on file for ' + a.name + '. Ask your manager for a sign-in code.');
      if (cache.get('sent:' + a.name)) throw new Error('A code was sent less than a minute ago — check your inbox (and spam).');
      var code = String(Math.floor(100000 + Math.random() * 900000));
      cache.put('code:' + a.name, JSON.stringify({ code: code, tries: 0 }), 600); cache.put('sent:' + a.name, '1', 60);
      MailApp.sendEmail({ to: a.email, subject: 'Your TMC Go sign-in code: ' + code, htmlBody: '<div style="font:15px Arial">Hi ' + String(a.name).split(' ')[0] + ',<br><br>Your TMC Go sign-in code is<br><div style="font:700 30px Arial;letter-spacing:6px;margin:12px 0;color:#1E2867">' + code + '</div>It works for 10 minutes. If you didn\'t ask for it, ignore this email.<br><br>— TMC Go</div>' });
      return { status: 'ok', to: maskEmail_(a.email) };
    }
    case 'verifyCode': {
      var a2 = aeRow_(b.name), code2 = String(b.code || '').replace(/\D/g, '');
      if (tooMany_('code:' + a2.name)) throw new Error('Too many tries — wait 15 minutes or ask your manager for a new code.');
      var held = cache.get('code:' + a2.name), ok = false;
      if (held) { var h = JSON.parse(held); if (h.tries >= 5) { cache.remove('code:' + a2.name); throw new Error('Too many tries — ask for a new code.'); } ok = h.code === code2; if (!ok) { h.tries++; cache.put('code:' + a2.name, JSON.stringify(h), 600); } }
      if (!ok) {   // a one-time code from a manager also works here
        var inv = props.getProperty('invite:' + code2);
        if (inv) { var iv = JSON.parse(inv); if (iv.ae === a2.name && new Date().getTime() < iv.exp) { ok = true; props.deleteProperty('invite:' + code2); } }
      }
      if (!ok) { failed_('code:' + a2.name); throw new Error('That code is not right or has expired.'); }
      cache.remove('code:' + a2.name); cleared_('code:' + a2.name);
      lockedWrite_(function () {});
      return { status: 'ok', session: issueSession_(a2.name, roleOf_(a2), b.device, held ? 'email' : 'invite') };
    }
    case 'managerSignIn': {
      var a3 = aeRow_(b.name);
      if (weakAdminPin_()) throw new Error('WEAKPIN: the admin PIN is still the default or too simple. Set a stronger one (at least 6 digits) in the Settings tab of the sheet first.');
      if (tooMany_('pin:' + a3.name)) throw new Error('Too many tries — wait 15 minutes.');
      try { checkAdmin_(b.adminPin); } catch (e) { failed_('pin:' + a3.name); throw e; }
      cleared_('pin:' + a3.name);
      // the first manager to sign in with the admin PIN becomes a manager on the team list
      if (roleOf_(a3) !== 'manager') setAeField_(a3.name, 'role', 'manager');
      return { status: 'ok', session: issueSession_(a3.name, 'manager', b.device, 'adminPin') };
    }
    case 'signOut': { var s = session_(b.token); if (s) { cache.remove('sess:' + b.token); var sh = sheet_('Sessions'), r = findRowBy_(sh, 'token', b.token); if (r !== -1) sh.deleteRow(r); } return { status: 'ok' }; }
    case 'whoami':  { var s2 = session_(b.token); if (!s2) throw new Error('SIGNIN: please sign in'); return { status: 'ok', session: s2 }; }
  }
  return null;
}
function lockedWrite_(fn) { fn(); }
// Slows down guessing: 8 wrong attempts per name (or per signed-in person) lock that door for 15 minutes.
function tooMany_(key) { var c = CacheService.getScriptCache(); return Number(c.get('fail:' + key) || 0) >= 8; }
function failed_(key) { var c = CacheService.getScriptCache(), n = Number(c.get('fail:' + key) || 0) + 1; c.put('fail:' + key, String(n), 900); }
function cleared_(key) { CacheService.getScriptCache().remove('fail:' + key); }
function weakAdminPin_() { var p = String(settings_().adminPin || ''); return !p || p.length < 6 || /^(\d)\1+$/.test(p) || /^(0123|1234|12345|123456|654321)$/.test(p); }
function findRowBy_(sh, col, val) {
  var hdr = headers_(sh), ci = hdr.indexOf(col), last = sh.getLastRow(); if (ci === -1 || last < 2) return -1;
  var v = sh.getRange(2, ci + 1, last - 1, 1).getValues();
  for (var i = 0; i < v.length; i++) if (String(v[i][0]) === String(val)) return i + 2;
  return -1;
}
function setAeField_(name, field, value) {
  var sh = sheet_('AEs'), hdr = headers_(sh), r = findRowBy_(sh, 'name', name), ci = hdr.indexOf(field);
  if (r === -1 || ci === -1) throw new Error('AE not found');
  sh.getRange(r, ci + 1).setValue(value);
}
function setAeActive_(name, active) {
  var sh = sheet_('AEs'), hdr = headers_(sh), last = sh.getLastRow();
  var iName = hdr.indexOf('name'), iActive = hdr.indexOf('active');
  if (last < 2) throw new Error('No AEs');
  var names = sh.getRange(2, 1, last - 1, hdr.length).getValues();
  for (var i = 0; i < names.length; i++) if (String(names[i][iName]) === String(name)) { sh.getRange(i + 2, iActive + 1).setValue(active); return { name: name, active: active }; }
  throw new Error('AE not found: ' + name);
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function newId_(prefix) {
  return prefix + Utilities.getUuid().slice(0, 8);
}
function findRow_(sh, id) {
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}
function upsert_(name, row, prefix) {
  var sh = sheet_(name), hdr = headers_(sh);
  if (!row.id) row.id = newId_(prefix);
  var r = findRow_(sh, row.id);
  var existing = r === -1 ? [] : sh.getRange(r, 1, 1, hdr.length).getValues()[0];
  var vals = hdr.map(function (c, i) {
    if (row[c] === undefined) return existing[i] === undefined ? '' : existing[i]; // keep columns the client didn't send
    return row[c] === null ? '' : row[c];
  });
  if (r === -1) sh.appendRow(vals);
  else sh.getRange(r, 1, 1, hdr.length).setValues([vals]);
  return row;
}
function setSetting_(key, value) {
  var sh = sheet_('Settings'), last = sh.getLastRow();
  if (last >= 2) {
    var keys = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < keys.length; i++) if (String(keys[i][0]) === key) { sh.getRange(i + 2, 2).setValue(value); return; }
  }
  sh.appendRow([key, value]);
}
function remove_(name, id) {
  var sh = sheet_(name);
  var r = findRow_(sh, id);
  if (r !== -1) sh.deleteRow(r);
  return { id: id };
}

/* ---------- public API (fetch) ---------- */
function doGet(e) {
  ensureSetup_();
  var p = (e && e.parameter) || {};
  if (p.ping) {
    var t0 = new Date().getTime(), info = { ok: true, time: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm') };
    try { info.tabs = SpreadsheetApp.getActiveSpreadsheet().getSheets().map(function (x) { return x.getName() + ':' + x.getLastRow(); }); } catch (e) { info.ok = false; info.error = String(e.message || e); }
    info.ms = new Date().getTime() - t0;
    return json_(info);
  }
  if (!p.action) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'ok', message: 'TMC Go backend is running. Open the app page and paste this link into it.' })).setMimeType(ContentService.MimeType.JSON);
  }
  try {
    return json_(handle_({ action: p.action, pin: p.pin, token: p.token }));
  } catch (err) {
    return json_({ status: 'error', message: String(err.message || err) });
  }
}
function doPost(e) {
  ensureSetup_();
  try {
    var body = JSON.parse(e.postData.contents || '{}');
    return json_(handle_(body));
  } catch (err) {
    return json_({ status: 'error', message: String(err.message || err) });
  }
}

/* ---------- public API (google.script.run, when served via HtmlService) ---------- */
function api(body) {
  ensureSetup_();
  try { return handle_(body || {}); }
  catch (err) { return { status: 'error', message: String(err.message || err) }; }
}


/* ---------- Google Maps: geocoding + routing (built-in Maps service) ---------- */
var CITY_BOUNDS = { // SW lat,lng → NE lat,lng, used to bias geocoding toward the AE's city
  'Lahore':     [31.25, 74.05, 31.75, 74.60],
  'Karachi':    [24.75, 66.80, 25.15, 67.45],
  'Islamabad':  [33.40, 72.75, 33.85, 73.35],
  'Rawalpindi': [33.40, 72.75, 33.85, 73.35],
  'Faisalabad': [31.30, 72.90, 31.55, 73.25],
  'Multan':     [30.10, 71.35, 30.30, 71.60],
  'Peshawar':   [33.90, 71.40, 34.10, 71.70]
};
function publicSettings_() {
  var s = settings_(), out = { autoKm: s.autoKm || 'Y', purposes: s.purposes || '', offices: {}, categoriesSync: s.categoriesSync || '', ai: !!(s.aiKey), maps: s.mapsKey ? 'platform' : 'builtin' };
  Object.keys(s).forEach(function (k) {
    var m = /^office\.([^.]+)$/.exec(k);
    if (m) out.offices[m[1]] = { address: s[k], latlng: s[k + '.latlng'] || '' };
  });
  return out;
}
function mapsKey_() { return String(settings_().mapsKey || '').trim(); }
function geocode_(query, city) {
  var key = mapsKey_();
  if (key) {
    var b0 = CITY_BOUNDS[city];
    var url = 'https://maps.googleapis.com/maps/api/geocode/json?address=' + encodeURIComponent(query) +
      '&region=pk&language=en&components=country:PK' +
      (b0 ? '&bounds=' + b0[0] + ',' + b0[1] + '|' + b0[2] + ',' + b0[3] : '') +
      '&key=' + encodeURIComponent(key);
    var j = http_(url, { method: 'get' });
    if (j.status === 'OK' && j.results && j.results.length) {
      var r0 = j.results[0];
      return { lat: r0.geometry.location.lat, lng: r0.geometry.location.lng, address: r0.formatted_address };
    }
    if (j.status !== 'ZERO_RESULTS') throw new Error('Google Maps geocoding error: ' + j.status + (j.error_message ? ' — ' + j.error_message : ''));
    return null;
  }
  var g = Maps.newGeocoder().setRegion('pk').setLanguage('en');
  var b = CITY_BOUNDS[city];
  if (b) g.setBounds(b[0], b[1], b[2], b[3]);
  var r = g.geocode(query);
  if (r.status !== 'OK' || !r.results || !r.results.length) return null;
  var res = r.results[0];
  return { lat: res.geometry.location.lat, lng: res.geometry.location.lng, address: res.formatted_address };
}
function office_(city) {
  city = String(city || '').split('/')[0].trim();
  var s = settings_();
  var addr = s['office.' + city];
  if (!addr) throw new Error('No office address for ' + city + ' — add office.' + city + ' in Settings (⚙ in the app).');
  var cached = s['office.' + city + '.latlng'];
  if (cached && /^-?[\d.]+,-?[\d.]+$/.test(cached)) {
    var p = cached.split(',');
    return { name: city + ' office', lat: Number(p[0]), lng: Number(p[1]), address: addr };
  }
  var g = geocode_(addr, city);
  if (!g) throw new Error('Google Maps could not find the ' + city + ' office address: ' + addr);
  setSetting_('office.' + city + '.latlng', g.lat + ',' + g.lng);
  return { name: city + ' office', lat: g.lat, lng: g.lng, address: g.address };
}
function haversine_(a, b) {
  var R = 6371, dLat = (b.lat - a.lat) * Math.PI / 180, dLng = (b.lng - a.lng) * Math.PI / 180;
  var x = Math.sin(dLat / 2) * Math.sin(dLat / 2) + Math.cos(a.lat * Math.PI / 180) * Math.cos(b.lat * Math.PI / 180) * Math.sin(dLng / 2) * Math.sin(dLng / 2);
  return 2 * R * Math.atan2(Math.sqrt(x), Math.sqrt(1 - x));
}
function normName_(s) { return String(s || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim(); }

// Find (or geocode and cache) the coordinates of one account. Returns {lat,lng,address} or null.
function locateAccount_(ae, name, city, save) {
  var sh = sheet_('Accounts'), hdr = headers_(sh), last = sh.getLastRow();
  var rows = last < 2 ? [] : sh.getRange(2, 1, last - 1, hdr.length).getValues();
  var iAe = hdr.indexOf('ae'), iName = hdr.indexOf('account'), iCity = hdr.indexOf('city'), iAddr = hdr.indexOf('address'), iLat = hdr.indexOf('lat'), iLng = hdr.indexOf('lng'), iLoc = hdr.indexOf('located');
  var key = normName_(name), hit = -1;
  for (var i = 0; i < rows.length; i++) if (normName_(rows[i][iName]) === key && String(rows[i][iAe]) === String(ae)) { hit = i; break; }
  if (hit === -1) for (var j = 0; j < rows.length; j++) if (normName_(rows[j][iName]) === key) { hit = j; break; }
  var row = hit === -1 ? null : rows[hit];
  if (row && row[iLat] !== '' && row[iLng] !== '' && !isNaN(Number(row[iLat]))) {
    return { lat: Number(row[iLat]), lng: Number(row[iLng]), address: String(row[iLoc] || row[iAddr] || ''), cached: true };
  }
  var c = (row && row[iCity]) ? String(row[iCity]) : city;
  var q = (row && row[iAddr]) ? String(row[iAddr]) : (name + ', ' + c + ', Pakistan');
  var g = geocode_(q, c);
  if (!g && row && row[iAddr]) g = geocode_(name + ', ' + c + ', Pakistan', c); // address failed → try the name
  if (!g) return null;
  if (save && row) {
    sh.getRange(hit + 2, iLat + 1).setValue(g.lat);
    sh.getRange(hit + 2, iLng + 1).setValue(g.lng);
    sh.getRange(hit + 2, iLoc + 1).setValue(g.address);
  }
  return { lat: g.lat, lng: g.lng, address: g.address, cached: false };
}

// Office → nearest account → next nearest … → office. Driving distance from Google Maps.
function routeForVisit_(row, save) {
  var city = String(row.city || '').split('/')[0].trim();
  // Stops come from the per-visit details when present (each may carry the office the AE chose for that visit);
  // otherwise from the accounts list. A client visited twice in a day is routed once, at the first position.
  var det = null; try { det = JSON.parse(row.details || ''); } catch (e) { det = null; }
  var items = (det && det.length) ? det.map(function (d) { return { name: String(d.client || '').trim(), lat: Number(d.lat), lng: Number(d.lng), address: d.addr || '' }; })
                                 : String(row.accounts || '').split(';').map(function (x) { return { name: x.trim(), lat: NaN, lng: NaN, address: '' }; });
  var uniq = [], seen = {};
  items.forEach(function (it) { if (!it.name) return; var k = normName_(it.name) + (isNaN(it.lat) ? '' : '@' + it.lat.toFixed(4) + ',' + it.lng.toFixed(4)); if (!seen[k]) { seen[k] = true; uniq.push(it); } });
  if (!uniq.length) return { totalKm: 0, legs: [], text: '', missing: [], office: city + ' office' };
  if (uniq.length > 23) throw new Error('Google Maps routes allow at most 23 stops in one day.');
  var office = office_(city);
  var pts = [], missing = [];
  uniq.forEach(function (it) {
    if (!isNaN(it.lat) && !isNaN(it.lng) && (it.lat || it.lng)) { pts.push({ name: it.name, lat: it.lat, lng: it.lng, address: it.address }); return; }
    var p = locateAccount_(row.ae, it.name, city, save);
    if (p) pts.push({ name: it.name, lat: p.lat, lng: p.lng, address: p.address }); else missing.push(it.name);
  });
  if (!pts.length) throw new Error('Could not locate any of these accounts on Google Maps: ' + missing.join(', ') + '. Add their addresses on the Accounts tab, or enter km manually.');
  // The AE decides the sequence: stops are routed exactly in the order they were entered.
  var order = pts;
  var legInfo = routeLegs_(office, order), stops = [office].concat(order, [office]), out = [], total = 0, mins = 0;
  for (var k = 0; k < legInfo.length; k++) {
    var km = Math.round(legInfo[k].meters / 100) / 10; total += km; mins += Math.round(legInfo[k].seconds / 60);
    out.push({ from: stops[k].name, to: stops[k + 1].name, km: km, mins: Math.round(legInfo[k].seconds / 60), altMin: Math.round(legInfo[k].min / 100) / 10, altMax: Math.round(legInfo[k].max / 100) / 10 });
  }
  total = Math.round(total * 10) / 10;
  var text = office.name + ' → ' + out.map(function (l) { return l.to + ' (' + l.km + ' km' + (l.mins ? ', ' + l.mins + ' min' : '') + ')'; }).join(' → ');
  return { totalKm: total, totalMins: mins, legs: out, text: text, missing: missing, office: office.name, engine: mapsKey_() ? 'platform' : 'builtin', stops: order.map(function (p) { return { name: p.name, address: p.address }; }) };
}

// Driving distance of each leg, office → stops → office, in metres.
// With a Maps Platform key: Routes API, traffic-aware for today's visits. Without: the built-in Maps service.
// Each leg is routed on its own, exactly as Google Maps does when you look up A → B: Google's recommended route
// (its own balance of time, distance and traffic — not simply the shortest), so the app shows the same number as the
// Maps app. Alternatives are asked for and their km range is kept so a figure can be sanity-checked.
// With a Maps Platform key: Routes API, traffic-aware for now. Without: the built-in Maps service.
function routeLegs_(office, order) {
  var stops = [office].concat(order, [office]), key = mapsKey_(), legs = [];
  for (var i = 0; i < stops.length - 1; i++) {
    var a = stops[i], b = stops[i + 1];
    if (Math.abs(a.lat - b.lat) < 1e-6 && Math.abs(a.lng - b.lng) < 1e-6) { legs.push({ meters: 0, seconds: 0, min: 0, max: 0 }); continue; }
    legs.push(key ? legPlatform_(a, b, key) : legBuiltin_(a, b));
  }
  return legs;
}
function legPlatform_(a, b, key) {
  var pt = function (p) { return { location: { latLng: { latitude: p.lat, longitude: p.lng } } }; };
  var body = { origin: pt(a), destination: pt(b), travelMode: 'DRIVE', routingPreference: 'TRAFFIC_AWARE', computeAlternativeRoutes: true, languageCode: 'en', units: 'METRIC' };
  var j = http_('https://routes.googleapis.com/directions/v2:computeRoutes', { method: 'post', contentType: 'application/json',
    headers: { 'X-Goog-Api-Key': key, 'X-Goog-FieldMask': 'routes.distanceMeters,routes.duration' }, payload: JSON.stringify(body) });
  if (!j.routes || !j.routes.length) throw new Error('Google Maps could not build the route' + (j.error && j.error.message ? ': ' + j.error.message : '.'));
  var all = j.routes.map(function (r) { return Number(r.distanceMeters || 0); });
  var secs = parseInt(String(j.routes[0].duration || '0').replace('s', ''), 10) || 0;
  return { meters: all[0], seconds: secs, min: Math.min.apply(null, all), max: Math.max.apply(null, all) };
}
function legBuiltin_(a, b) {
  var dir = Maps.newDirectionFinder().setMode(Maps.DirectionFinder.Mode.DRIVING).setAlternatives(true)
    .setOrigin(a.lat, a.lng).setDestination(b.lat, b.lng).getDirections();
  if (dir.status !== 'OK' || !dir.routes || !dir.routes.length) throw new Error('Google Maps could not build the route (' + dir.status + ').');
  var all = dir.routes.map(function (r) { return r.legs.reduce(function (s, l) { return s + Number(l.distance.value || 0); }, 0); });
  var secs = dir.routes[0].legs.reduce(function (s, l) { return s + Number(l.duration.value || 0); }, 0);
  return { meters: all[0], seconds: secs, min: Math.min.apply(null, all), max: Math.max.apply(null, all) };
}

// Geocode accounts in bulk (ids, or every account without coordinates when ids is empty). Bounded per call.
function locateAccounts_(ids, limit) {
  var sh = sheet_('Accounts'), hdr = headers_(sh), last = sh.getLastRow();
  var rows = last < 2 ? [] : sh.getRange(2, 1, last - 1, hdr.length).getValues();
  var iId = hdr.indexOf('id'), iName = hdr.indexOf('account'), iCity = hdr.indexOf('city'), iAddr = hdr.indexOf('address'), iLat = hdr.indexOf('lat'), iLng = hdr.indexOf('lng'), iLoc = hdr.indexOf('located');
  var want = ids && ids.length ? {} : null;
  (ids || []).forEach(function (x) { want[x] = true; });
  var done = 0, results = [], max = Math.min(Number(limit || 40), 60);
  for (var i = 0; i < rows.length && done < max; i++) {
    var r = rows[i];
    if (want ? !want[r[iId]] : ((r[iLat] !== '' && r[iLat] !== null && !isNaN(Number(r[iLat]))) || String(r[iLoc]) === 'NOT FOUND')) continue; // bulk: skip located + already-failed (retry those by id)
    var c = String(r[iCity] || ''), q = r[iAddr] ? String(r[iAddr]) : (r[iName] + ', ' + c + ', Pakistan');
    var g = geocode_(q, c);
    if (!g && r[iAddr]) g = geocode_(r[iName] + ', ' + c + ', Pakistan', c);
    done++;
    if (g) {
      sh.getRange(i + 2, iLat + 1).setValue(g.lat); sh.getRange(i + 2, iLng + 1).setValue(g.lng); sh.getRange(i + 2, iLoc + 1).setValue(g.address);
      results.push({ id: r[iId], account: r[iName], lat: g.lat, lng: g.lng, located: g.address });
    } else {
      sh.getRange(i + 2, iLoc + 1).setValue('NOT FOUND');
      results.push({ id: r[iId], account: r[iName], located: 'NOT FOUND' });
    }
    Utilities.sleep(120);
  }
  return results;
}

// Candidate locations for a client name. With a Maps Platform key: Places text search (branches, plants, head
// offices as separate hits). Without: the built-in geocoder's result list. Never requires a typed address.
function findPlaces_(query, city) {
  query = String(query || '').trim(); city = String(city || '').split('/')[0].trim();
  if (!query) return [];
  var key = mapsKey_(), out = [], seen = {};
  var push = function (name, address, lat, lng) { var k = String(lat).slice(0, 7) + ',' + String(lng).slice(0, 7); if (seen[k]) return; seen[k] = true; out.push({ name: name, address: address, lat: lat, lng: lng }); };
  if (key) {
    var b = CITY_BOUNDS[city];
    var body = { textQuery: query + (city ? ' ' + city : ''), regionCode: 'PK', languageCode: 'en', maxResultCount: 6 };
    if (b) body.locationBias = { rectangle: { low: { latitude: b[0], longitude: b[1] }, high: { latitude: b[2], longitude: b[3] } } };
    var j = http_('https://places.googleapis.com/v1/places:searchText', { method: 'post', contentType: 'application/json', headers: { 'X-Goog-Api-Key': key, 'X-Goog-FieldMask': 'places.displayName,places.formattedAddress,places.location' }, payload: JSON.stringify(body) });
    (j.places || []).forEach(function (p) { if (p.location) push((p.displayName && p.displayName.text) || query, p.formattedAddress || '', p.location.latitude, p.location.longitude); });
    if (out.length) return out.slice(0, 6);
  }
  var g = Maps.newGeocoder().setRegion('pk').setLanguage('en');
  var bb = CITY_BOUNDS[city]; if (bb) g.setBounds(bb[0], bb[1], bb[2], bb[3]);
  var tries = [query + (city ? ', ' + city : '') + ', Pakistan', query + ', Pakistan'];
  for (var t = 0; t < tries.length && out.length < 4; t++) {
    var r = g.geocode(tries[t]);
    if (r.status === 'OK') (r.results || []).slice(0, 5).forEach(function (res) { push(query, res.formatted_address, res.geometry.location.lat, res.geometry.location.lng); });
  }
  return out.slice(0, 6);
}

// Save exact coordinates for one account (an AE standing at the client taps "pin from my location").
function pinAccount_(ae, name, lat, lng, city) {
  lat = Number(lat); lng = Number(lng);
  if (!name || isNaN(lat) || isNaN(lng)) throw new Error('Need the account and a location.');
  if (lat < -90 || lat > 90 || lng < -180 || lng > 180) throw new Error('That location does not look right.');
  var sh = sheet_('Accounts'), hdr = headers_(sh), last = sh.getLastRow();
  var rows = last < 2 ? [] : sh.getRange(2, 1, last - 1, hdr.length).getValues();
  var iAe = hdr.indexOf('ae'), iName = hdr.indexOf('account'), iLat = hdr.indexOf('lat'), iLng = hdr.indexOf('lng'), iLoc = hdr.indexOf('located');
  var key = normName_(name), hit = -1;
  for (var i = 0; i < rows.length; i++) if (normName_(rows[i][iName]) === key && String(rows[i][iAe]) === String(ae)) { hit = i; break; }
  if (hit === -1) for (var j = 0; j < rows.length; j++) if (normName_(rows[j][iName]) === key) { hit = j; break; }
  var stamp = 'Pinned on site ' + Utilities.formatDate(new Date(), TZ, 'd MMM yyyy');
  if (hit === -1) { var row = upsert_('Accounts', { ae: ae, account: name, tier: '', city: city || '', netNew: 'Y', addedOn: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'), notes: 'Added from visit log', lat: lat, lng: lng, located: stamp }, 'A'); return row; }
  sh.getRange(hit + 2, iLat + 1).setValue(lat); sh.getRange(hit + 2, iLng + 1).setValue(lng); sh.getRange(hit + 2, iLoc + 1).setValue(stamp);
  var out = {}; hdr.forEach(function (c, k) { out[c] = rows[hit][k]; }); out.lat = lat; out.lng = lng; out.located = stamp;
  return out;
}

/* ---------- Categories sync + aliases ---------- */
// Categories tab (AE | Account | Category | City) → Accounts tab. Only runs when the tab has rows.
function syncCategories_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), cat = ss.getSheetByName(CATEGORIES_TAB);
  if (!cat || cat.getLastRow() < 2) return;
  var ch = headers_(cat);
  var iAe = findCol_(ch, 'ae'), iAcc = findCol_(ch, 'account'), iCat = findCol_(ch, 'category'), iCity = findCol_(ch, 'city');
  if (iAe === -1 || iAcc === -1 || iCat === -1) return;
  var rows = cat.getRange(2, 1, cat.getLastRow() - 1, ch.length).getValues();
  var aes = readTab_('AEs'), unknown = {}, resolvedCount = 0;
  var wanted = {}; // ae|normname -> {tier, city, account, ae}
  rows.forEach(function (r) {
    var rawAe = String(r[iAe] || '').trim(), acc = String(r[iAcc] || '').trim(), tier = String(r[iCat] || '').trim().toUpperCase();
    var city = iCity === -1 ? '' : String(r[iCity] || '').trim();
    if (!acc || /^total\b/i.test(acc)) return;
    if (['A', 'B', 'C'].indexOf(tier) === -1) tier = tier === 'PARTNER' ? 'Partner' : '';
    var ae = resolveAe_(rawAe, city, aes);
    if (!ae && rawAe && !/^tbh$/i.test(rawAe)) { unknown[rawAe] = true; ae = rawAe; } else if (ae) resolvedCount++;
    wanted[ae + '|' + normName_(acc)] = { ae: ae, account: acc, tier: tier, city: city };
  });
  try { setSetting_('categoriesSync', JSON.stringify({ at: new Date().toISOString(), rows: rows.length, unknownAEs: Object.keys(unknown) })); } catch (e2) {}
  var sh = sheet_('Accounts'), hdr = headers_(sh), last = sh.getLastRow();
  var acc = last < 2 ? [] : sh.getRange(2, 1, last - 1, hdr.length).getValues();
  var jAe = hdr.indexOf('ae'), jAcc = hdr.indexOf('account'), jTier = hdr.indexOf('tier'), jCity = hdr.indexOf('city'), jNew = hdr.indexOf('netNew');
  // Work in memory, then write whole columns at once — one cell at a time is far too slow for a book this size.
  var tierCol = [], cityCol = [], changed = false, seen = {};
  for (var i = 0; i < acc.length; i++) {
    var tier = String(acc[i][jTier] || ''), city = String(acc[i][jCity] || '');
    var key = String(acc[i][jAe]) + '|' + normName_(acc[i][jAcc]), w = wanted[key];
    if (w) {
      seen[key] = true;
      if (tier !== w.tier) { tier = w.tier; changed = true; }
      if (w.city && city !== w.city) { city = w.city; changed = true; }
    } else if (['A', 'B', 'C'].indexOf(tier) !== -1 && String(acc[i][jNew]).toUpperCase() !== 'Y') {
      tier = ''; changed = true; // no longer in the manager's list → no category
    }
    tierCol.push([tier]); cityCol.push([city]);
  }
  if (changed && tierCol.length) {
    sh.getRange(2, jTier + 1, tierCol.length, 1).setValues(tierCol);
    sh.getRange(2, jCity + 1, cityCol.length, 1).setValues(cityCol);
  }
  var add = [];
  Object.keys(wanted).forEach(function (key) {
    if (seen[key]) return;
    var w = wanted[key], row = { id: newId_('A'), ae: w.ae, account: w.account, tier: w.tier, city: w.city, netNew: 'N', addedOn: '', notes: 'From Categories tab' };
    add.push(hdr.map(function (c) { return row[c] === undefined ? '' : row[c]; }));
  });
  if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, hdr.length).setValues(add);
}
// Map a free-form AE name (as typed in the Categories tab) to a name in the AEs tab; city breaks ties.
function dice_(a, b) {
  var A = {}, B = {}, n = 0, sa = 0, sb = 0; a = String(a); b = String(b);
  for (var i = 0; i < a.length - 1; i++) { A[a.substr(i, 2)] = true; }
  for (var j = 0; j < b.length - 1; j++) { B[b.substr(j, 2)] = true; }
  for (var k in A) { sa++; if (B[k]) n++; } for (var m in B) sb++;
  return sa && sb ? 2 * n / (sa + sb) : 0;
}
function resolveAe_(raw, city, aes) {
  raw = String(raw || '').trim(); if (!raw || /^tbh$/i.test(raw) || /^total\b/i.test(raw)) return '';
  var rl = raw.toLowerCase();
  for (var e = 0; e < aes.length; e++) if (String(aes[e].name).toLowerCase() === rl) return aes[e].name;
  var rt = rl.split(/[^a-z]+/).filter(Boolean), best = [], bs = 0;
  aes.forEach(function (a) {
    var nt = String(a.name).toLowerCase().split(/[^a-z]+/).filter(Boolean), hit = 0, firstHit = false;
    rt.forEach(function (t, idx) {
      var ok = nt.some(function (n) { return n === t || (t.length >= 3 && n.indexOf(t) === 0) || (t.length === 1 && n.charAt(0) === t) || (t.length >= 5 && dice_(t, n) >= 0.65); });
      if (ok) { hit++; if (idx === 0) firstHit = true; }
    });
    var score = hit / rt.length;
    if (!firstHit) score = 0;
    if (score > bs) { bs = score; best = [a]; } else if (score === bs && score > 0) best.push(a);
  });
  if (bs < 0.5 || !best.length) return '';
  if (best.length > 1 && city) {
    var c = String(city).toLowerCase();
    var primary = best.filter(function (a) { return String(a.city || '').toLowerCase().split('/')[0].trim() === c; });
    if (primary.length === 1) return primary[0].name;
    var same = best.filter(function (a) { return String(a.city || '').toLowerCase().indexOf(c) !== -1; });
    if (same.length === 1) return same[0].name;
  }
  return best.length === 1 ? best[0].name : '';
}
function upsertAlias_(alias, account, ae) {
  alias = String(alias || '').trim(); account = String(account || '').trim(); ae = String(ae || '').trim();
  if (!alias || !account || normName_(alias) === normName_(account)) return null;
  var sh = sheet_('Aliases'), last = sh.getLastRow();
  var rows = last < 2 ? [] : sh.getRange(2, 1, last - 1, 3).getValues();
  for (var i = 0; i < rows.length; i++) if (normName_(rows[i][0]) === normName_(alias) && String(rows[i][2]) === ae) { if (String(rows[i][1]) !== account) sh.getRange(i + 2, 2).setValue(account); return { alias: alias, account: account, ae: ae }; }
  sh.appendRow([alias, account, ae]);
  return { alias: alias, account: account, ae: ae };
}


/* ---------- TMC Assistant (AI layer) ---------- */
var AI_SYSTEM = [
  "You are Go — the assistant inside TMC Go, the field-visits app used by TMC's account executives (AEs) and sales managers. You are funny, friendly and empathetic: a sharp colleague who has sat in on every visit, notices when someone has had a long day, celebrates wins, and never lectures. Jokes are welcome (original, about field-sales life in Pakistan — Shaheed-e-Millat and Ferozepur Road traffic, DHA Business Hub lifts, Friday go-lives, tenders at 11:59 pm, chai between visits, the PO that is always 'next week', SAP acronyms); never about a real person, always clean, never the same one twice, and never when someone is reporting a problem.",
  "",
  "ABOUT TMC (TMC Private Limited, formerly TallyMarks Consulting; tmcltd.com): founded 2010 in Lahore as a BI specialist; SAP partner since 2011, Gold since 2018; Global SAP Gold Partner and the largest SAP System Integrator in Pakistan for six consecutive years (2020–2025); Global Winner at the 2026 SAP Partner Awards (Customer Transformation: Corporate), the first Pakistani company to win a global SAP award; SAP Quality Awards (MENA); acquired Siemens Pakistan's SAP ERP reselling business on 1 April 2023; 750+ SAP consultants; Qlik Master Reseller; TallyMarks Cloud since 2018; offices in Lahore (DHA Business Hub), Karachi (Nasreen Building, Shaheed-e-Millat Road), Islamabad (SNC Center, Blue Area). Sells SAP S/4HANA, RISE/GROW with SAP, Business One, SuccessFactors, Ariba, BTP, Analytics Cloud, DMC, Qlik, Tableau, Salesforce, cloud, cyber security, HR solutions, PeopleSoft, EFFI, and TMC's AI products Satori and Verity.",
  "",
  "THE CHORES YOU TAKE OFF AN AE'S PLATE (do them, don't describe them):",
  "- LOG VISITS from a sentence or a conversation ('aaj Ghani Glass aur Dawn Foods gaya, SAP demo'): match names to the ACCOUNTS in the data even if mangled (Dawn bread → Dawn Foods, Daimond Phone → Diamond Foam, M&P → Muller & Phipps), keep the order the AE visited them, take the purpose if given (match to PURPOSES where it fits; don't nag for it), kal/yesterday = yesterday. Then emit the log action. If they say save/done/haan/yes, set save:true; otherwise leave save:false so they can review.",
  "- PLAN: tomorrow's or the week's visits from the UNTOUCHED accounts (A before B before C), grouped sensibly by area/city, with the one-line reason for each.",
  "- WRITE: follow-up emails and WhatsApp messages after a visit, meeting summaries, visit notes, proposal outlines, talking points, objection handling — in TMC's voice, in the user's language. Put the finished text in the draft action so it can be copied or sent in one tap.",
  "- SUMMARISE: a day, week, month, AE or client, for the AE or for their manager; call out wins, risks and what to do next.",
  "- PREP: a meeting with a named client from what the data shows (last visits, purposes, category, opportunities).",
  "- REMEMBER: when the AE mentions something in play at a client (a demo booked, a tender coming, a renewal), offer the opportunity action so it is recorded.",
  "- NAVIGATE: change the dashboard period when asked.",
  "- ANSWER anything about coverage, visits, kilometres, net-new, who is behind, trends — from the DATA only; never invent a visit, client or figure, and say plainly when something is not in the data.",
  "",
  "HOW YOU TALK: in the language the user uses (English, Urdu script, or Roman Urdu — never mix scripts). Short and specific; bold the key numbers; bullets only for lists; no tables; under ~150 words unless a draft or plan needs more. Names exactly as in the data.",
  "",
  "ACTIONS (append at most one fenced json block after your sentence; the app executes it):",
  '{"log":[{"date":"YYYY-MM-DD","client":"<exact account name, or as said if unknown>","purpose":"<short or empty>"}],"save":false}',
  '{"range":{"preset":"today|yesterday|week|last7|month|lastMonth|quarter"}} or {"range":{"from":"YYYY-MM-DD","to":"YYYY-MM-DD"}}',
  '{"draft":{"title":"Follow-up email|WhatsApp message|Visit notes|…","subject":"<for emails>","text":"<the full text>"}}',
  '{"opportunity":{"client":"<exact account name>","text":"<what is in play>","stage":"Open|Demo|Proposal|Closing"}}',
  "Only emit an action when it is clearly what the user wants.",
  "",
  "VOICE: a message starting with [voice:ur] was spoken in Urdu — reply in Urdu script, 2–3 short spoken sentences, no lists or markdown, English only for names and numbers. [voice:ro] → Roman Urdu. [voice:en] → spoken English, 2–3 sentences. In voice, confirm what you captured out loud before saving."
].join('\n');

function aiConfig_() {
  var s = settings_();
  return { provider: 'gemini', key: String(s.aiKey || '').trim(), model: String(s.aiModel || '').trim() || 'gemini-3.5-flash' };
}
function normKey_(s) { return normName_(s); }
// Compact, live digest of the sheet for the model: team table, per-AE detail, recent days.
function aiContext_(range, city, me) {
  // (opportunities and the user's full untouched book are appended at the end)
  var from = (range && range.from) || Utilities.formatDate(new Date(), TZ, 'yyyy-MM-01');
  var to = (range && range.to) || Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  var aes = readTab_('AEs').filter(function (a) { return String(a.active || 'Y').toUpperCase() !== 'N'; });
  var accounts = readTab_('Accounts'), aliases = readTab_('Aliases');
  var visits = readTab_('Visits').filter(function (v) { return v.date >= from && v.date <= to; });
  var aliasMap = {}; aliases.forEach(function (x) { aliasMap[normKey_(x.alias)] = x.account; });
  var byAe = {}; accounts.forEach(function (a) { (byAe[a.ae] = byAe[a.ae] || []).push(a); });
  function findAcc(ae, name) {
    var k = normKey_(name), al = aliasMap[k]; if (al) k = normKey_(al);
    var list = byAe[ae] || [];
    for (var i = 0; i < list.length; i++) if (normKey_(list[i].account) === k) return list[i];
    return null;
  }
  function ownerOf(name) { var k = normKey_(name), al = aliasMap[k]; if (al) k = normKey_(al); for (var i = 0; i < accounts.length; i++) if (accounts[i].ae && normKey_(accounts[i].account) === k) return accounts[i].ae; return ''; }
  var lines = ['LIVE DATA (period ' + from + ' to ' + to + (city ? ', city ' + city : '') + '; today ' + Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd') + (me ? '; the user is AE ' + me : '; the user is a manager') + ')', ''];
  var team = { asg: 0, vis: 0, visits: 0, km: 0, newAcc: 0 };
  var aeLines = [], detail = [];
  aes.forEach(function (a) {
    if (city && String(a.city || '').indexOf(city) === -1) return;
    var mine = visits.filter(function (v) { return v.ae === a.name; });
    var counts = {}; mine.forEach(function (v) { String(v.accounts || '').split(';').map(function (x) { return x.trim(); }).filter(Boolean).forEach(function (n) { counts[n] = (counts[n] || 0) + 1; }); });
    var book = byAe[a.name] || [], tiers = { A: [0, 0], B: [0, 0], C: [0, 0] }, notV = [], visitedOwn = {};
    Object.keys(counts).forEach(function (n) { var r = findAcc(a.name, n); if (r) visitedOwn[r.id] = (visitedOwn[r.id] || 0) + counts[n]; });
    book.forEach(function (b) { if (['A', 'B', 'C'].indexOf(b.tier) === -1) return; tiers[b.tier][0]++; if (visitedOwn[b.id]) tiers[b.tier][1]++; else notV.push(b.account + ' (' + b.tier + ')'); });
    var asg = tiers.A[0] + tiers.B[0] + tiers.C[0], vis = tiers.A[1] + tiers.B[1] + tiers.C[1];
    var nv = Object.keys(counts).reduce(function (s, n) { return s + counts[n]; }, 0), km = mine.reduce(function (s, v) { return s + Number(v.km || 0); }, 0);
    var newAcc = Object.keys(counts).filter(function (n) { var r = findAcc(a.name, n); return !r || ['A', 'B', 'C'].indexOf(r.tier) === -1; });
    var other = Object.keys(counts).filter(function (n) { return !findAcc(a.name, n) && ownerOf(n); });
    team.asg += asg; team.vis += vis; team.visits += nv; team.km += km; team.newAcc += newAcc.length;
    var measured = String(a.coverageMeasured || 'Y').toUpperCase() !== 'N' && asg > 0;
    aeLines.push(a.name + ' (' + (a.city || '') + '): ' + (measured ? 'coverage ' + Math.round(vis / asg * 100) + '% (' + vis + '/' + asg + '; A ' + tiers.A[1] + '/' + tiers.A[0] + ', B ' + tiers.B[1] + '/' + tiers.B[0] + ', C ' + tiers.C[1] + '/' + tiers.C[0] + ')' : 'no assigned book') + '; visits ' + nv + '; distinct accounts ' + Object.keys(counts).length + '; net-new/unlisted ' + newAcc.length + '; other AEs\' accounts visited ' + other.length + '; km ' + Math.round(km * 10) / 10 + (a.notes ? '; note: ' + a.notes : ''));
    if (measured && notV.length) detail.push(a.name + ' not yet visited (' + notV.length + '): ' + notV.slice(0, 14).join(', ') + (notV.length > 14 ? ', …' : ''));
    if (newAcc.length) detail.push(a.name + ' net-new/unlisted clients: ' + newAcc.slice(0, 10).join(', ') + (newAcc.length > 10 ? ', …' : ''));
  });
  lines.push('TEAM: coverage ' + (team.asg ? Math.round(team.vis / team.asg * 100) : 0) + '% (' + team.vis + '/' + team.asg + ' assigned accounts visited), visits ' + team.visits + ', km ' + Math.round(team.km) + ', net-new accounts ' + team.newAcc + '.');
  lines.push('', 'PER AE:'); lines = lines.concat(aeLines);
  lines.push('', 'DETAIL:'); lines = lines.concat(detail);
  var recent = visits.slice().sort(function (x, y) { return String(y.date).localeCompare(String(x.date)) || String(y.ts).localeCompare(String(x.ts)); }).slice(0, 45);
  lines.push('', 'RECENT DAYS (newest first; client — purpose; km for the day):');
  recent.forEach(function (v) {
    var det; try { det = JSON.parse(v.details || ''); } catch (e) { det = null; }
    var items = det && det.length ? det.map(function (d) { return d.client + (d.purpose ? ' — ' + d.purpose : ''); }) : String(v.accounts || '').split(';').map(function (x) { return x.trim(); }).filter(Boolean);
    lines.push(v.date + ' ' + v.ae + ': ' + (items.length ? items.join('; ') : (v.purpose || 'no visits')) + ' | ' + (v.kmMode === 'pending' ? 'km pending' : Math.round(Number(v.km || 0) * 10) / 10 + ' km'));
  });
  // Opportunities in play (what is being sold where) and the user's full untouched book, so planning and prep are concrete.
  try {
    var opps = readTab_('Opportunities').filter(function (o) { return !me || !o.ae || String(o.ae).indexOf(me) !== -1 || String(o.ae).indexOf(String(me).split(' ')[0]) !== -1; }).slice(0, 60);
    if (opps.length) { lines.push('', 'OPPORTUNITIES IN PLAY (AE: client — what):'); opps.forEach(function (o) { lines.push((o.ae || '') + ': ' + o.client + ' — ' + o.opportunity + (o.stage ? ' [' + o.stage + ']' : '')); }); }
  } catch (e) {}
  if (me) {
    var mine = accounts.filter(function (a) { return a.ae === me && ['A', 'B', 'C'].indexOf(String(a.tier)) !== -1; });
    var visitedKeys = {};
    visits.filter(function (v) { return v.ae === me; }).forEach(function (v) { String(v.accounts || '').split(';').forEach(function (n) { if (n.trim()) visitedKeys[normKey_(n)] = true; }); });
    var untouched = mine.filter(function (a) { return !visitedKeys[normKey_(a.account)]; }).sort(function (x, y) { return String(x.tier).localeCompare(String(y.tier)); });
    lines.push('', me + "'s UNTOUCHED ACCOUNTS this period (" + untouched.length + '):', untouched.map(function (a) { return a.account + ' [' + a.tier + ']' + (a.city ? ' ' + a.city : ''); }).join('; ') || 'none');
  }
  var st = settings_(); var offs = Object.keys(st).filter(function (k) { return /^office\.[^.]+$/.test(k) && st[k]; }).map(function (k) { return k.slice(7) + ': ' + st[k]; });
  if (offs.length) lines.push('', 'OFFICES: ' + offs.join(' | '));
  var text = lines.join('\n');
  if (text.length > 60000) text = text.slice(0, 60000) + '\n…(truncated)';
  return text;
}
function askSafe_(b) {
  try { return { reply: ask_(b) }; }
  catch (e) {
    var m = String(e.message || e);
    if (/Authorization|permission|ScriptError/i.test(m)) return { error: 'Go is not approved to reach Google yet. A manager needs to open Apps Script, run checkSetup once and click Allow.' };
    if (/API key not valid|API_KEY_INVALID|403/i.test(m)) return { error: 'The Gemini key in the sheet is not valid — a manager can create one at aistudio.google.com and paste it in Settings → aiKey.' };
    if (/not found|404|NOT_FOUND/i.test(m)) return { error: 'The Gemini model set in the sheet is not available. A manager can clear the aiModel cell to use the default.' };
    if (/no-ai|No AI key/i.test(m)) return { error: 'A manager still needs to add the Gemini key in the sheet (Settings → aiKey).' };
    if (/429|quota|RESOURCE_EXHAUSTED/i.test(m)) return { error: 'Gemini is busy right now — give it a minute and try again.' };
    return { error: 'Go could not answer just now. Try again in a moment.' };
  }
}
function ask_(b) {
  var cfg = aiConfig_();
  if (!cfg.key) throw new Error('The assistant is not set up yet — a manager pastes a Gemini API key into aiKey in the Settings tab of the sheet.');
  var system = AI_SYSTEM + '\n\n' + aiContext_(b.range, b.city, b.me);
  var msgs = (b.messages || []).filter(function (m) { return m && m.content; }).slice(-12).map(function (m) { return { role: m.role === 'assistant' ? 'assistant' : 'user', content: String(m.content).slice(0, 4000) }; });
  if (!msgs.length) throw new Error('Nothing to ask.');
  var text = callGemini_(cfg, system, msgs);
  return String(text || '').trim() || 'I could not produce an answer — try again.';
}
function http_(url, opts) {
  opts.muteHttpExceptions = true;
  var res = UrlFetchApp.fetch(url, opts), code = res.getResponseCode(), body = res.getContentText();
  if (code >= 300) { var msg = body; try { var j = JSON.parse(body); msg = (j.error && (j.error.message || j.error.type)) || body; } catch (e) {} throw new Error('AI provider error ' + code + ': ' + String(msg).slice(0, 300)); }
  return JSON.parse(body);
}
function callGemini_(cfg, system, msgs) {
  var body = { system_instruction: { parts: [{ text: system }] }, contents: msgs.map(function (m) { return { role: m.role === 'assistant' ? 'model' : 'user', parts: [{ text: m.content }] }; }), generationConfig: { temperature: 0.6, maxOutputTokens: 1024 } };
  var j = http_('https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(cfg.model) + ':generateContent', { method: 'post', contentType: 'application/json', headers: { 'x-goog-api-key': cfg.key }, payload: JSON.stringify(body) });
  var c = j.candidates && j.candidates[0];
  return c && c.content && c.content.parts ? c.content.parts.map(function (p) { return p.text || ''; }).join('') : '';
}
/* ---------- Self-check: is Go (Gemini) and are Maps actually able to run? ---------- */
// Run checkSetup() once from the Apps Script editor after pasting a new Code.gs: it triggers the one-time
// permission prompts (external requests for Gemini, Maps) that a deployment cannot ask for by itself.
function checkSetup() {
  var r = selfCheck_();
  Logger.log(JSON.stringify(r, null, 2));
  return r;
}
function selfCheck_() {
  var out = { gemini: {}, maps: {}, sheet: {} };
  try { out.sheet = { ok: true, tabs: SpreadsheetApp.getActiveSpreadsheet().getSheets().length }; } catch (e) { out.sheet = { ok: false, error: String(e.message || e) }; }
  var cfg = aiConfig_();
  if (!cfg.key) out.gemini = { ok: false, error: 'No key: put a Gemini API key in the Settings tab, row aiKey.' };
  else {
    try {
      var txt = callGemini_(cfg, 'Reply with exactly: ready', [{ role: 'user', content: 'ping' }]);
      out.gemini = { ok: true, model: cfg.model, reply: String(txt).slice(0, 40) };
    } catch (e) {
      var m = String(e.message || e);
      out.gemini = { ok: false, error: /Authorization|permission|ScriptError/i.test(m) ? 'The script needs a one-time approval for internet access: open Apps Script, pick checkSetup in the function list, press Run, and click Allow. Then deploy a new version.' : /API key not valid|API_KEY_INVALID|403/i.test(m) ? 'Google says the key is not valid. Create one at aistudio.google.com/app/apikey and paste it in the Settings tab (aiKey).' : /not found|404|NOT_FOUND/i.test(m) ? 'The model "' + cfg.model + '" is not available on this key. Clear the aiModel cell (default) or set it to a current Gemini model.' : m };
    }
  }
  try { out.mail = { ok: true, remainingToday: MailApp.getRemainingDailyQuota() }; } catch (e) { out.mail = { ok: false, error: 'Sign-in emails need a one-time approval: run checkSetup in Apps Script and click Allow.' }; }
  out.adminPin = weakAdminPin_() ? { ok: false, error: 'The admin PIN is the default or too simple — change it in the Settings tab (at least 6 digits). Until then nobody can sign in as a manager with it.' } : { ok: true };
  try {
    var city = Object.keys(settings_()).filter(function (k) { return /^office\.[^.]+$/.test(k); }).map(function (k) { return k.slice(7); })[0] || 'Lahore';
    var o = office_(city); out.maps = { ok: true, engine: mapsKey_() ? 'platform' : 'builtin', office: city, pinned: o.lat.toFixed(4) + ',' + o.lng.toFixed(4) };
  } catch (e) {
    var mm = String(e.message || e);
    out.maps = { ok: false, error: /Authorization|permission|ScriptError/i.test(mm) ? 'Maps needs the same one-time approval: run checkSetup in Apps Script and click Allow.' : mm };
  }
  return out;
}

/* ---------- dispatcher ---------- */
function handle_(b) {
  checkPin_(b.pin);
  var authRes = authAction_(b); if (authRes) return authRes;
  var sess = requireSession_(b);
  var isMgr = sess.role === 'manager';
  // an AE can only act as themselves; managers (or the admin PIN) may act for others
  var mine = function (ae) { if (!ae || ae === sess.ae || isMgr) return; try { checkAdmin_(b.adminPin); } catch (e) { throw new Error('NOTYOU: you can only change your own visits'); } };
  // Only sheet WRITES are serialised, and only for the moment of writing. Reading, Gemini, Google Maps lookups and
  // route calculations never hold the lock — they are what made other people wait until "Lock timeout".
  var lock = null;
  var acquire = function () { lock = LockService.getScriptLock(); if (!lock.tryLock(30000)) { lock = null; throw new Error('BUSY: the sheet is busy saving for someone else'); } };
  var NOLOCK = { all: 1, ask: 1, selfCheck: 1, findPlaces: 1, findPlace: 1, computeRoute: 1, checkAdmin: 1, locateAccounts: 1, upsertVisit: 1, recalcKm: 1 };
  if (!NOLOCK[b.action]) acquire();
  try {
    switch (b.action) {
      case 'all':
        try {
          var cache = CacheService.getScriptCache(), cat0 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CATEGORIES_TAB);
          var stamp = cat0 ? String(cat0.getLastRow()) + 'x' + String(cat0.getLastColumn()) + '@' + String(SpreadsheetApp.getActiveSpreadsheet().getLastUpdated ? '' : '') : '0';
          if (cache.get('catsync') !== stamp) { var sl = LockService.getScriptLock(); if (sl.tryLock(0)) { try { syncCategories_(); cache.put('catsync', stamp, 3600); } finally { sl.releaseLock(); } } }
        } catch (syncErr) { /* a bad Categories row must not block loading */ }
        return { status: 'ok', data: {
          aes: readTab_('AEs'), accounts: readTab_('Accounts'),
          visits: readTab_('Visits'), opportunities: readTab_('Opportunities'),
          aliases: readTab_('Aliases'), settings: publicSettings_()
        }, session: sess };
      case 'upsertAlias':       return { status: 'ok', row: upsertAlias_(b.alias, b.account, b.ae) };
      case 'ask': {             b.me = sess.ae; var ar = askSafe_(b); if (ar.error) throw new Error(ar.error); return { status: 'ok', reply: ar.reply }; }
      case 'upsertVisit': {
        var vrow = b.row || {};
        if (!vrow.ae) vrow.ae = sess.ae;
        if (vrow.id) { var exV = readTab_('Visits').filter(function (x) { return x.id === vrow.id; })[0]; if (exV) mine(exV.ae); }
        mine(vrow.ae);
        if (b.adminPin !== undefined) checkAdmin_(b.adminPin, sess); // dashboard (manager) edits identify themselves
        if (vrow.note !== undefined && b.adminPin === undefined) { // only managers write the day note
          var old = findRow_(sheet_('Visits'), vrow.id); if (old !== -1) { var vh0 = headers_(sheet_('Visits')), ni = vh0.indexOf('note'); vrow.note = ni === -1 ? '' : sheet_('Visits').getRange(old, ni + 1).getValue(); } else vrow.note = '';
        }
        if (b.autoKm) {
          // Never lose a visit because Maps is down or over quota: save it as 'pending' and let the app retry.
          var rt = null, mapsErr = null;
          try { rt = routeForVisit_(vrow, true); } catch (e) { mapsErr = e; }      // slow part: no lock held
          if (rt) { vrow.km = rt.totalKm; vrow.route = rt.text; vrow.kmMode = 'auto'; }
          else { vrow.km = ''; vrow.route = ''; vrow.kmMode = 'pending'; }
          acquire();                                                             // quick part: the write
          return rt ? { status: 'ok', row: upsert_('Visits', vrow, 'V'), route: rt } : { status: 'ok', row: upsert_('Visits', vrow, 'V'), pending: String(mapsErr.message || mapsErr) };
        }
        if (!vrow.kmMode) vrow.kmMode = 'manual';
        acquire();
        return { status: 'ok', row: upsert_('Visits', vrow, 'V') };
      }
      case 'recalcKm': {
        var vs = sheet_('Visits'), vh = headers_(vs), vr = findRow_(vs, b.id);
        if (vr === -1) throw new Error('Visit not found');
        var vals = vs.getRange(vr, 1, 1, vh.length).getValues()[0], cur = {};
        vh.forEach(function (c, i) { var v = vals[i]; if (v instanceof Date) v = Utilities.formatDate(v, TZ, 'yyyy-MM-dd'); cur[c] = v; });
        mine(cur.ae);
        var rt2 = routeForVisit_(cur, true);                                      // slow part: no lock held
        cur.km = rt2.totalKm; cur.route = rt2.text; cur.kmMode = 'auto';
        acquire();
        return { status: 'ok', row: upsert_('Visits', cur, 'V'), route: rt2 };
      }
      case 'computeRoute':      return { status: 'ok', route: routeForVisit_(b.row || {}, true) };
      case 'locateAccounts':    return { status: 'ok', results: locateAccounts_(b.ids, b.limit) };
      case 'pinAccount':        mine(b.ae); return { status: 'ok', row: pinAccount_(b.ae || sess.ae, b.account, b.lat, b.lng, b.city) };
      case 'findPlaces':        return { status: 'ok', places: findPlaces_(b.query, b.city) };
      case 'findPlace': {       // look up an address or place name, so an AE can set a client's location by hand
        var hit = geocode_(String(b.query || '').trim(), String(b.city || '').trim());
        if (!hit) throw new Error('Google Maps could not find that. Try a fuller address, or paste a Google Maps link.');
        return { status: 'ok', place: hit };
      }
      case 'removeAE':  checkAdmin_(b.adminPin, sess); revokeSessions_(b.name); return { status: 'ok', row: setAeActive_(b.name, 'N') };
      case 'inviteCode': {      // a one-time sign-in code for an AE, valid 24 h — share it by WhatsApp / in person
        checkAdmin_(b.adminPin, sess); aeRow_(b.name);
        var code = String(Math.floor(100000 + Math.random() * 900000));
        PropertiesService.getScriptProperties().setProperty('invite:' + code, JSON.stringify({ ae: b.name, exp: new Date().getTime() + 864e5 }));
        return { status: 'ok', code: code, expires: '24 hours' };
      }
      case 'revokeSessions': checkAdmin_(b.adminPin, sess); return { status: 'ok', revoked: revokeSessions_(b.name) };
      case 'setRole':   checkAdmin_(b.adminPin, sess); setAeField_(b.name, 'role', b.role === 'manager' ? 'manager' : ''); revokeSessions_(b.name); return { status: 'ok' };
      case 'setEmail':  checkAdmin_(b.adminPin, sess); setAeField_(b.name, 'email', String(b.email || '').trim()); return { status: 'ok' };
      case 'sessions':  checkAdmin_(b.adminPin, sess); return { status: 'ok', sessions: readTab_('Sessions').map(function (r) { return { ae: r.ae, role: r.role, created: r.created, lastSeen: r.lastSeen, device: r.device, method: r.method }; }) };
      case 'restoreAE': checkAdmin_(b.adminPin, sess); return { status: 'ok', row: setAeActive_(b.name, 'Y') };
      case 'checkAdmin': checkAdmin_(b.adminPin, sess); return { status: 'ok' };
      case 'selfCheck':  checkAdmin_(b.adminPin, sess); return { status: 'ok', check: selfCheck_() };
      case 'setSettings': {
        checkAdmin_(b.adminPin, sess);
        if (b.settings && b.settings.adminPin !== undefined) delete b.settings.adminPin; // never set over the API
                var out = {};
        Object.keys(b.settings || {}).forEach(function (k) { setSetting_(k, b.settings[k]); if (/^office\./.test(k)) setSetting_(k + '.latlng', ''); });
        Object.keys(b.settings || {}).forEach(function (k) {
          if (/^office\.[^.]+$/.test(k) && b.settings[k]) { try { out[k] = office_(k.slice(7)); } catch (err) { out[k] = { error: String(err.message || err) }; } }
        });
        return { status: 'ok', offices: out, settings: publicSettings_() };
      }
      case 'deleteVisit':       { var dv = readTab_('Visits').filter(function (x) { return x.id === b.id; })[0]; if (dv) mine(dv.ae); return { status: 'ok', row: remove_('Visits', b.id) }; }
      case 'upsertAccount':     mine((b.row || {}).ae); return { status: 'ok', row: upsert_('Accounts', b.row, 'A') };
      case 'deleteAccount':     return { status: 'ok', row: remove_('Accounts', b.id) };
      case 'upsertOpportunity': mine((b.row || {}).ae); return { status: 'ok', row: upsert_('Opportunities', b.row, 'O') };
      case 'deleteOpportunity': return { status: 'ok', row: remove_('Opportunities', b.id) };
      case 'upsertAE': {
        var sh = sheet_('AEs'), cols = TABS.AEs, row = b.row || {};
        if (!row.name) throw new Error('AE name missing');
        if (row.photo && !/^data:image\/(png|jpeg|webp);base64,[A-Za-z0-9+\/=]+$/.test(String(row.photo))) throw new Error('That photo could not be saved — please choose a picture again.');
        if (String(row.photo || '').length > 60000) throw new Error('That photo is too large — please choose a smaller picture.');
        ['title', 'phone', 'notes'].forEach(function (k) { if (row[k] !== undefined) row[k] = String(row[k]).replace(/[<>]/g, '').slice(0, k === 'notes' ? 300 : 60); });
        if (row.name !== sess.ae) checkAdmin_(b.adminPin, sess);           // your own profile, or a manager
        var last = sh.getLastRow(), r = -1, cur = null;
        if (last >= 2) {
          var rows = sh.getRange(2, 1, last - 1, cols.length).getValues();
          for (var i = 0; i < rows.length; i++) if (String(rows[i][0]) === String(row.name)) { r = i + 2; cur = rows[i]; }
        }
        if (!cur) checkAdmin_(b.adminPin, sess); // new AEs are added by managers only; existing rows (notes, city) use the team PIN
        // the active flag is only changed through removeAE / restoreAE (admin PIN); dashboard notes are manager-only
        row.active = cur ? cur[cols.indexOf('active')] : 'Y';
        if (cur && row.role !== undefined && String(row.role || '') !== String(cur[cols.indexOf('role')] || '')) checkAdmin_(b.adminPin, sess);   // only managers change roles
        if (cur && row.email !== undefined && String(row.email || '') !== String(cur[cols.indexOf('email')] || '')) checkAdmin_(b.adminPin, sess); // the sign-in email belongs to managers
        if (cur && row.notes !== undefined && String(row.notes || '') !== String(cur[cols.indexOf('notes')] || '')) checkAdmin_(b.adminPin, sess);
        var vals = cols.map(function (c) { return row[c] === undefined ? (cur ? cur[cols.indexOf(c)] : '') : row[c]; });
        if (r === -1) sh.appendRow(vals); else sh.getRange(r, 1, 1, cols.length).setValues([vals]);
        var saved = {}; cols.forEach(function (c, j) { saved[c] = vals[j]; });
        return { status: 'ok', row: saved };
      }
      case 'bulkVisits': {
        (b.rows || []).forEach(function (x) { mine(x.ae); });
        var rows = (b.rows || []).map(function (x) { return upsert_('Visits', x, 'V'); });
        return { status: 'ok', rows: rows };
      }
      default: throw new Error('Unknown action: ' + b.action);
    }
  } finally {
    if (lock) lock.releaseLock();
  }
}
