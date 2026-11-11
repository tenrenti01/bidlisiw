/**
 * Bidlisiw · PSHS-CVC Batch 2007 — database tools (container-bound Apps Script)
 *
 *  Menu  Bidlisiw ▸ Link photos      matches the Drive folder's photos to Students / Pictures and writes photo_url
 *  Web app = the DATA API for the website on GitHub Pages (https://tenrenti01.github.io/bidlisiw/).
 *    The page itself (index.html) now lives in the GitHub repo tenrenti01/bidlisiw — no Index.html file is needed here.
 *    Deploy ▸ New deployment ▸ Web app · Execute as: Me · Who has access: Anyone
 *    then paste the /exec URL into config.js in the repo.
 *    The website calls it with POST {fn, args} (or GET ?fn=…&args=[…]) and gets back {ok, v} or {ok:false, e}.
 *  What the website shows
 *    · Teachers      (main view, loads first) every teacher: edit first name / last name / nickname,
 *                    and the subjects they taught with the year level (one teacher, many subjects)
 *                    → sheets "Teachers" and "Teacher Subjects"
 *                    a batchmate can leave ONE testimonial per teacher → sheet "Teacher Testimonials"
 *                    (the app never shows these texts to anyone; it only knows who has already written)
 *    · Testimonials  (second tab) pick who you are → your message to the batch, if none yet → write to batchmates
 *                    new ones are marked year 2026 in the Testimonials sheet (the app never shows that)
 *    · Profiles      (third tab) view every batchmate, fill in / edit the basic facts, set the sections
 *  Devices sheet: a browser that has written as someone stays that person. To unlock one, delete its row.
 */

const PHOTO_FOLDER_ID = '1qX6vA0lUF9Z3a14XClGhCYglAXFsQCI-';
const SS = () => SpreadsheetApp.getActive();

/* ------------------------------------------------------------------ menu */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Bidlisiw')
    .addItem('Link photos (Drive → Students, Pictures)', 'linkPhotos')
    .addItem('Make photos viewable without sign-in (faster)', 'sharePhotos')
    .addItem('Recount testimonials', 'recountTestimonials')
    .addItem('API URL & website…', 'showWebAppUrl')
    .addToUi();
}
function showWebAppUrl() {
  const url = ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().alert(url ? 'API URL (paste this into config.js on GitHub):\n' + url + '\n\nWebsite: https://tenrenti01.github.io/bidlisiw/' : 'Not deployed yet.\nDeploy ▸ New deployment ▸ Web app (Execute as: Me · Who has access: Anyone), then run this again.');
}

/* ------------------------------------------------------------------ sheet helpers */
function sheet_(name) {
  const sh = SS().getSheetByName(name);
  if (!sh) throw new Error('Sheet not found: ' + name);
  return sh;
}
/** {headers:[...], rows:[[...]], col:{header:index0}} for a sheet */
function table_(name) {
  const sh = sheet_(name), vals = sh.getDataRange().getValues();
  const headers = vals[0].map(h => String(h).trim());
  const col = {}; headers.forEach((h, i) => col[h] = i);
  return { sh, headers, rows: vals.slice(1), col };
}
/** make sure the named columns exist (appended at the right if missing); returns the header → index map */
function ensureColumns_(name, wanted) {
  const sh = sheet_(name);
  let headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
  wanted.forEach(w => {
    if (headers.indexOf(w) < 0) {
      sh.getRange(1, headers.length + 1).setValue(w).setFontWeight('bold').setBackground('#5e0f1b').setFontColor('#f6dfa0');
      headers.push(w);
    }
  });
  const col = {}; headers.forEach((h, i) => col[h] = i);
  return col;
}
function baseName_(p) { p = String(p || ''); const i = p.lastIndexOf('/'); return (i >= 0 ? p.slice(i + 1) : p).trim(); }
function thumbUrl_(id, w) { return 'https://drive.google.com/thumbnail?id=' + id + '&sz=w' + (w || 600); }
function nowStamp_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'); }

/* ------------------------------------------------------------------ 1 · link photos */
function linkPhotos() {
  const folder = DriveApp.getFolderById(PHOTO_FOLDER_ID);
  const byName = {};
  const it = folder.getFiles();
  while (it.hasNext()) { const f = it.next(); byName[f.getName().toLowerCase()] = f.getId(); }

  // Students: photo_url + drive_file_id
  const sc = ensureColumns_('Students', ['photo_url', 'drive_file_id']);
  const st = table_('Students');
  let matched = 0, missing = [];
  const urlOut = [], idOut = [];
  st.rows.forEach(r => {
    const sid = String(r[st.col.student_id] || '').trim();
    if (!sid) { urlOut.push(['']); idOut.push(['']); return; }
    const want = baseName_(r[st.col.photo_file]).toLowerCase();
    let id = byName[want] || byName[sid.toLowerCase() + '.jpg'] || byName[sid.toLowerCase() + '.png'] || byName[sid.toLowerCase() + '.jpeg'];
    if (!id) {  // tolerate a different extension / case: match on the stem
      const stem = want.replace(/\.[a-z0-9]+$/, '');
      const k = Object.keys(byName).find(n => n.replace(/\.[a-z0-9]+$/, '') === stem);
      if (k) id = byName[k];
    }
    if (id) { matched++; urlOut.push([thumbUrl_(id)]); idOut.push([id]); }
    else { if (want) missing.push(sid + ' (' + want + ')'); urlOut.push(['']); idOut.push(['']); }
  });
  if (st.rows.length) {
    st.sh.getRange(2, sc.photo_url + 1, urlOut.length, 1).setValues(urlOut);
    st.sh.getRange(2, sc.drive_file_id + 1, idOut.length, 1).setValues(idOut);
  }

  // Pictures: same two columns, by file name
  const pc = ensureColumns_('Pictures', ['photo_url', 'drive_file_id']);
  const pt = table_('Pictures');
  if (pt.rows.length) {
    const pu = [], pi = [];
    pt.rows.forEach(r => {
      const id = byName[baseName_(r[pt.col.file]).toLowerCase()];
      pu.push([id ? thumbUrl_(id) : '']); pi.push([id || '']);
    });
    pt.sh.getRange(2, pc.photo_url + 1, pu.length, 1).setValues(pu);
    pt.sh.getRange(2, pc.drive_file_id + 1, pi.length, 1).setValues(pi);
  }
  bust_(); try { CACHE().removeAll(st.rows.map(r => 'p:' + String(r[st.col.student_id] || '') + '#n')); } catch (e) { }
  const msg = 'Linked ' + matched + ' of ' + st.rows.length + ' students to photos in the Drive folder (' + Object.keys(byName).length + ' files).'
    + (missing.length ? '\n\nNo photo found for: ' + missing.join(', ') : '');
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
  return msg;
}

/* ------------------------------------------------------------------ counts */
function recountTestimonials() {
  const st = table_('Students'), tt = table_('Testimonials');
  const recv = {}, wrote = {};
  tt.rows.forEach(r => { const a = String(r[tt.col.about_id] || ''), b = String(r[tt.col.by_id] || ''); if (a) recv[a] = (recv[a] || 0) + 1; if (b) wrote[b] = (wrote[b] || 0) + 1; });
  const out = st.rows.map(r => { const sid = String(r[st.col.student_id] || ''); return [recv[sid] || 0, wrote[sid] || 0]; });
  if (out.length && st.col.testimonials_received !== undefined && st.col.testimonials_written === st.col.testimonials_received + 1) {
    st.sh.getRange(2, st.col.testimonials_received + 1, out.length, 2).setValues(out);
  } else if (out.length) {
    st.sh.getRange(2, st.col.testimonials_received + 1, out.length, 1).setValues(out.map(x => [x[0]]));
    st.sh.getRange(2, st.col.testimonials_written + 1, out.length, 1).setValues(out.map(x => [x[1]]));
  }
}

/* ------------------------------------------------------------------ cache (faster loads; every write clears what it touches) */
const CACHE = () => CacheService.getScriptCache();
const BOOT_TTL = 90, PROFILE_TTL = 300;   /* seconds */
function cached_(key, ttl, make) {
  try {   /* big values are stored in ~90 KB pieces: key#n holds the count, key#0..n-1 the parts */
    const c = CACHE(), n = Number(c.get(key + '#n') || 0);
    if (n) { const parts = c.getAll(Array.from({ length: n }, (_, i) => key + '#' + i)), txt = Array.from({ length: n }, (_, i) => parts[key + '#' + i]); if (txt.every(x => x != null)) return JSON.parse(txt.join('')); }
  } catch (e) { }
  const v = make();
  try {
    const txt = JSON.stringify(v), SZ = 90000, n = Math.ceil(txt.length / SZ), put = {};
    for (let i = 0; i < n; i++) put[key + '#' + i] = txt.slice(i * SZ, (i + 1) * SZ);
    put[key + '#n'] = String(n); CACHE().putAll(put, ttl);
  } catch (e) { }
  return v;
}
function bust_(id) { try { CACHE().removeAll(['boot#n'].concat(id ? ['p:' + id + '#n'] : [])); } catch (e) { } }

/* ------------------------------------------------------------------ photos for visitors who aren't signed in to Google
   Drive only shows a private photo to people allowed to see it. The web app runs as the owner, so it can read the
   photo and hand it over as data instead: the app asks for this whenever a Drive link fails (e.g. incognito). */
const PHOTO_TTL = 21600;   /* 6 hours, the longest the cache allows */
function apiPhotos(ids) {
  ids = (ids || []).slice(0, 20);
  const st = table_('Students'), fid = {};
  st.rows.forEach(r => fid[String(r[st.col.student_id])] = st.col.drive_file_id === undefined ? '' : String(r[st.col.drive_file_id] || ''));
  const out = {}, c = CACHE();
  let hits = {}; try { hits = c.getAll(ids.map(id => 'img:' + fid[id])); } catch (e) { }
  ids.forEach(id => {
    const f = fid[id]; if (!f) { out[id] = ''; return; }
    if (hits['img:' + f]) { out[id] = hits['img:' + f]; return; }
    try {
      const blob = DriveApp.getFileById(f).getBlob();
      const url = 'data:' + (blob.getContentType() || 'image/jpeg') + ';base64,' + Utilities.base64Encode(blob.getBytes());
      out[id] = url;
      if (url.length < 99000) try { c.put('img:' + f, url, PHOTO_TTL); } catch (e) { }
    } catch (e) { out[id] = ''; }
  });
  return out;
}
/** optional: let anyone with the link view the portraits, so they load straight from Google's image servers (fastest).
    The links are only visible inside the web app. Run it again after adding photos by hand to the folder. */
function sharePhotos() {
  const it = DriveApp.getFolderById(PHOTO_FOLDER_ID).getFiles();
  let n = 0, bad = 0;
  while (it.hasNext()) { const f = it.next(); if (share_(f)) n++; else bad++; }
  const msg = 'Made ' + n + ' photo' + (n === 1 ? '' : 's') + ' viewable by anyone with the link.' + (bad ? '\n' + bad + ' could not be changed (your Drive settings may block link sharing); the web app still shows them its own way.' : '');
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
  return msg;
}
function share_(f) { try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); return true; } catch (e) { return false; } }

/* ------------------------------------------------------------------ web app = JSON API for the GitHub Pages site */
/** only these can be called from the website */
function apiTable_() {
  return { apiTeachers, apiTeacherPhotos, apiSaveTeacher, apiAddTeacherSubject, apiUpdateTeacherSubject, apiDeleteTeacherSubject, apiAddTeacherTestimonial,
    apiBootstrap, apiGetProfile, apiPhotos, apiSaveProfile, apiAddDreams, apiAddSongs, apiUploadPhoto, apiSaveSections,
    apiWhoAmI, apiSaveMessage, apiAddTestimonial };
}
function doGet(e) {
  const p = (e && e.parameter) || {};
  let args = []; try { args = p.args ? JSON.parse(p.args) : []; } catch (x) { }
  return reply_(p.fn, args);
}
function doPost(e) {
  let body = {}; try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (x) { }
  return reply_(body.fn, body.args || []);
}
function reply_(fn, args) {
  let out;
  if (!fn) out = { ok: true, v: 'Bidlisiw API is running.' };
  else {
    const f = apiTable_()[fn];
    if (!f) out = { ok: false, e: 'Unknown function: ' + fn };
    else try { out = { ok: true, v: f.apply(null, [].concat(args)) }; } catch (err) { out = { ok: false, e: String((err && err.message) || err) }; }
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/** everything the app needs at start: the roster, the section lists, and known year sections */
function apiBootstrap() { return cached_('boot', BOOT_TTL, buildBootstrap_); }
function buildBootstrap_() {
  const st = table_('Students');
  /* the facts the Missing sheet checks, with the labels shown in the app */
  const FACTS = [['first_name', 'first name'], ['middle_name', 'middle name'], ['last_name', 'last name'], ['birthday', 'birthday'], ['hometown', 'hometown'], ['home_province', 'home province'], ['current_location', 'current location'], ['self_quote', 'self-quote'], ['photo_file', 'photo'], ['section', 'section (IV)']];
  const students = st.rows.filter(r => r[st.col.student_id]).map(r => ({
    id: String(r[st.col.student_id]),
    name: String(r[st.col.full_name] || ''),
    display: String(r[st.col.display_name] || r[st.col.first_name] || r[st.col.student_id]),
    photo: st.col.photo_url !== undefined ? String(r[st.col.photo_url] || '') : '',
    section: String(r[st.col.section] || ''),
    missing: FACTS.filter(([k]) => st.col[k] === undefined || String(r[st.col[k]] || '').trim() === '').map(([, label]) => label)
  }));
  const ls = table_('Lists');
  const sections = {};
  ls.rows.forEach(r => {
    const y = Number(r[ls.col.year_level]); if (!y) return;
    sections[y] = ls.headers.filter(h => /^section_/.test(h)).map(h => String(r[ls.col[h]] || '').trim()).filter(Boolean);
  });
  const ys = table_('Year Sections');
  const yearSections = {};
  ys.rows.forEach(r => {
    const sid = String(r[ys.col.student_id] || ''), y = Number(r[ys.col.year_level]); if (!sid || !y) return;
    (yearSections[sid] = yearSections[sid] || {})[y] = String(r[ys.col.section] || '');
  });
  /* every batchmate's full profile, read in one pass, so opening a person needs no extra trip to the server */
  const T = { st, dr: table_('Dreams'), so: table_('Songs'), co: table_('Contacts'), tt: table_('Testimonials'), ys };
  const profiles = {};
  students.forEach(x => profiles[x.id] = buildProfile_(x.id, T));
  /* who has already written to whom this time round: the app greys those out (the texts themselves are not sent) */
  const given = {}, given2007 = {};
  T.tt.rows.forEach(r => { const by = String(r[T.tt.col.by_id] || ''), about = String(r[T.tt.col.about_id] || ''); if (!by || !about) return;
    const g = isNew_(T.tt, r) ? given : given2007; if ((g[by] || []).indexOf(about) < 0) (g[by] = g[by] || []).push(about); });
  /* who already has a message to the batch on file */
  const hasMessage = {}; st.rows.forEach(r => { const id = String(r[st.col.student_id] || ''); if (id) hasMessage[id] = st.col.message_to_batch !== undefined && String(r[st.col.message_to_batch] || '').trim() !== ''; });
  return { students, sections, yearSections, profiles, given, given2007, hasMessage, at: nowStamp_() };
}

const EDITABLE = ['first_name', 'middle_name', 'last_name', 'nicknames', 'birthday', 'hometown', 'home_province', 'current_location', 'self_quote', 'message_to_batch', 'email_current'];
/* what was written in 2007 is kept as it was: these can be filled in when blank, never changed once on file */
const ADD_ONLY = ['self_quote', 'message_to_batch'];

function apiGetProfile(id) { return cached_('p:' + id, PROFILE_TTL, () => buildProfile_(id)); }
function buildProfile_(id, T) {
  T = T || {};
  const st = T.st || table_('Students');
  const i = st.rows.findIndex(r => String(r[st.col.student_id]) === id);
  if (i < 0) throw new Error('Unknown student: ' + id);
  const r = st.rows[i], get = k => st.col[k] === undefined ? '' : String(r[st.col[k]] || '');
  const p = { id, row: i + 2 };
  ['full_name', 'display_name', 'zodiac', 'photo_file', 'photo_url', 'section', 'testimonials_received', 'testimonials_written'].concat(EDITABLE).forEach(k => p[k] = get(k));
  const dr = T.dr || table_('Dreams'); p.dreams = dr.rows.filter(x => String(x[dr.col.student_id]) === id).map(x => String(x[dr.col.dream] || ''));
  const so = T.so || table_('Songs'); p.songs = so.rows.filter(x => String(x[so.col.student_id]) === id).map(x => String(x[so.col.song_raw] || ''));
  const co = T.co || table_('Contacts'); p.contacts = co.rows.filter(x => String(x[co.col.student_id]) === id).map(x => ({ type: String(x[co.col.type] || ''), value: String(x[co.col.value] || '') }));
  const tt = T.tt || table_('Testimonials');
  const names = T.names || (T.names = {}); if (!Object.keys(names).length) st.rows.forEach(x => names[String(x[st.col.student_id])] = String(x[st.col.display_name] || x[st.col.student_id]));
  p.testimonials = tt.rows.filter(x => String(x[tt.col.about_id]) === id && !isNew_(tt, x)).map(x => ({
    id: String(x[tt.col.testimonial_id] || ''), by: String(x[tt.col.by_id] || ''), byName: names[String(x[tt.col.by_id])] || String(x[tt.col.by_name_raw] || 'Batchmate'),
    text: String(x[tt.col.text] || ''), note: String(x[tt.col.personal_note] || '')
  }));
  p.wrote = tt.rows.filter(x => String(x[tt.col.by_id]) === id && !isNew_(tt, x)).map(x => names[String(x[tt.col.about_id])] || String(x[tt.col.about_id]));
  const ys = T.ys || table_('Year Sections'); p.years = {};
  ys.rows.filter(x => String(x[ys.col.student_id]) === id).forEach(x => p.years[Number(x[ys.col.year_level])] = String(x[ys.col.section] || ''));
  return p;
}

function zodiac_(bday) {
  const m = String(bday || '').match(/^([A-Za-z]+)\s+(\d{1,2})/); if (!m) return '';
  const months = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'];
  const mo = months.indexOf(m[1].toLowerCase()) + 1, d = Number(m[2]); if (!mo) return '';
  const z = [[1, 19, 'Capricorn'], [2, 18, 'Aquarius'], [3, 20, 'Pisces'], [4, 19, 'Aries'], [5, 20, 'Taurus'], [6, 20, 'Gemini'], [7, 22, 'Cancer'], [8, 22, 'Leo'], [9, 22, 'Virgo'], [10, 22, 'Libra'], [11, 21, 'Scorpio'], [12, 21, 'Sagittarius'], [12, 31, 'Capricorn']];
  for (const [zm, zd, name] of z) if (mo === zm && d <= zd) return name;
  return mo === 12 ? 'Capricorn' : z.find(x => x[0] === mo + 1)[2];
}

/** save the editable facts of one student; only the keys sent are written */
function apiSaveProfile(id, fields) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const st = table_('Students');
    const i = st.rows.findIndex(r => String(r[st.col.student_id]) === id);
    if (i < 0) throw new Error('Unknown student: ' + id);
    const row = i + 2;
    Object.keys(fields || {}).forEach(k => {
      if (EDITABLE.indexOf(k) < 0 || st.col[k] === undefined) return;
      if (ADD_ONLY.indexOf(k) >= 0 && String(st.rows[i][st.col[k]] || '').trim()) return;   /* already on file: keep the 2007 text */
      st.sh.getRange(row, st.col[k] + 1).setValue(String(fields[k] || '').trim());
    });
    if (fields && fields.birthday !== undefined && st.col.zodiac !== undefined) st.sh.getRange(row, st.col.zodiac + 1).setValue(zodiac_(fields.birthday));
    if (fields && fields.first_name !== undefined && st.col.display_name !== undefined) {
      const cur = String(st.rows[i][st.col.display_name] || '');
      if (!cur) st.sh.getRange(row, st.col.display_name + 1).setValue(String(fields.first_name || '').split(' ')[0]);
    }
    bust_(id);
    return apiGetProfile(id);
  } finally { lock.releaseLock(); }
}

/** dreams can be added only when none are on file: the 2007 lists are never changed or extended */
function apiAddDreams(id, lines) {
  lines = [].concat(lines || []).map(x => String(x || '').trim()).filter(Boolean);
  if (!id || !lines.length) throw new Error('Write at least one dream.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const dr = table_('Dreams');
    if (dr.rows.some(r => String(r[dr.col.student_id]) === id)) throw new Error('Dreams are already on file for this batchmate; they stay as written.');
    lines = lines.filter((x, i) => lines.findIndex(y => y.toLowerCase() === x.toLowerCase()) === i);
    lines.forEach((text, i) => { const rec = {}; dr.headers.forEach(h => rec[h] = '');
      rec.dream_id = id + '-D' + (i + 1); rec.student_id = id; rec.order = i + 1; rec.dream = text; if ('source' in rec) rec.source = 'web app · ' + nowStamp_();
      dr.sh.appendRow(dr.headers.map(h => rec[h])); });
    bust_(id);
    return lines;
  } finally { lock.releaseLock(); }
}
/** the playlist can be added only when none is on file (up to 5 songs, as in 2007) */
function apiAddSongs(id, lines) {
  lines = [].concat(lines || []).map(x => String(x || '').trim()).filter(Boolean);
  if (!id || !lines.length) throw new Error('Write at least one song.');
  if (lines.length > 5) throw new Error('Five songs at most, as in 2007.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const so = table_('Songs');
    if (so.rows.some(r => String(r[so.col.student_id]) === id)) throw new Error('A playlist is already on file for this batchmate; it stays as it was.');
    lines = lines.filter((x, i) => lines.findIndex(y => y.toLowerCase() === x.toLowerCase()) === i);
    lines.forEach((text, i) => { const rec = {}; so.headers.forEach(h => rec[h] = '');
      const parts = text.split(/\s+[—–-]\s+/);
      rec.song_id = id + '-S' + (i + 1); rec.student_id = id; rec.rank = i + 1; rec.song_raw = text; rec.title = parts[0]; rec.artist = parts.slice(1).join(' — ');
      so.sh.appendRow(so.headers.map(h => rec[h])); });
    bust_(id);
    return lines;
  } finally { lock.releaseLock(); }
}

/** a portrait can be uploaded only for a batchmate who has none on file; it goes into the Drive photo folder as <id>.jpg */
function apiUploadPhoto(id, dataUrl) {
  const m = String(dataUrl || '').match(/^data:(image\/(?:jpeg|png|webp));base64,(.+)$/);
  if (!id || !m) throw new Error('Choose an image file (JPG or PNG).');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sc = ensureColumns_('Students', ['photo_url', 'drive_file_id']);
    const st = table_('Students');
    const i = st.rows.findIndex(r => String(r[st.col.student_id]) === id);
    if (i < 0) throw new Error('Unknown student: ' + id);
    const r = st.rows[i];
    if (String(r[st.col.drive_file_id] || '').trim() || String(r[st.col.photo_url] || '').trim()) throw new Error('A portrait is already on file for this batchmate.');
    const bytes = Utilities.base64Decode(m[2]);
    if (bytes.length > 4 * 1024 * 1024) throw new Error('Image too large (4 MB max).');
    const ext = m[1] === 'image/png' ? 'png' : m[1] === 'image/webp' ? 'webp' : 'jpg';
    const blob = Utilities.newBlob(bytes, m[1], id + '.' + ext);
    const file = DriveApp.getFolderById(PHOTO_FOLDER_ID).createFile(blob);
    try { const others = DriveApp.getFolderById(PHOTO_FOLDER_ID).getFiles(); while (others.hasNext()) { const o = others.next(); if (o.getId() !== file.getId()) { if (o.getSharingAccess() === DriveApp.Access.ANYONE_WITH_LINK) share_(file); break; } } } catch (e) { }
    const row = i + 2, url = thumbUrl_(file.getId());
    st.sh.getRange(row, sc.photo_file + 1).setValue('img/' + id + '.' + ext);
    st.sh.getRange(row, sc.photo_url + 1).setValue(url);
    st.sh.getRange(row, sc.drive_file_id + 1).setValue(file.getId());
    const pc = ensureColumns_('Pictures', ['photo_url', 'drive_file_id']);
    const pt = table_('Pictures'); const rec = {}; pt.headers.forEach(h => rec[h] = '');
    rec.picture_id = id + '-P1'; rec.student_id = id; rec.kind = 'portrait, uploaded ' + nowStamp_(); rec.file = 'img/' + id + '.' + ext; rec.source = 'web app'; rec.photo_url = url; rec.drive_file_id = file.getId();
    pt.sh.appendRow(pt.headers.map(h => rec[h] === undefined ? '' : rec[h]));
    bust_(id);
    return { url, fileId: file.getId() };
  } finally { lock.releaseLock(); }
}

/* ------------------------------------------------------------------ testimonials written now (2026) and the browser that wrote them */
const NEW_YEAR = '2026';
/** a testimonial row written through the web app (year column), as opposed to the 2007 files */
function isNew_(tt, r) { return tt.col.year !== undefined && String(r[tt.col.year] || '').trim() === NEW_YEAR; }
/** the Devices sheet: device_id / fingerprint → student_id. A browser that has written as someone is that person from then on.
    Two keys, because the app runs inside Google's sandboxed frame where saved browser data is often wiped or blocked:
    device_id is what the browser managed to save, fingerprint is what the device looks like (screen, language, time zone…). */
function devices_() {
  let sh = SS().getSheetByName('Devices');
  if (!sh) { sh = SS().insertSheet('Devices'); sh.getRange(1, 1, 1, 5).setValues([['device_id', 'fingerprint', 'student_id', 'claimed_at', 'user_agent']]).setFontWeight('bold').setBackground('#5e0f1b').setFontColor('#f6dfa0'); sh.setFrozenRows(1); }
  const col = ensureColumns_('Devices', ['device_id', 'fingerprint', 'student_id', 'claimed_at', 'user_agent']);
  return table_('Devices');
}
function deviceRow_(dv, device, fp) {
  device = String(device || '').trim(); fp = String(fp || '').trim();
  return dv.rows.find(r => (device && String(r[dv.col.device_id]) === device) || (fp && String(r[dv.col.fingerprint]) === fp)) || null;
}
/** who this browser is, or null when it has never written (or its row was deleted to unlock it) */
function apiWhoAmI(device, fp) {
  if (!device && !fp) return null;
  const dv = devices_(), row = deviceRow_(dv, device, fp);
  return row ? { id: String(row[dv.col.student_id]) } : null;
}
/** bind a browser to a batchmate on its first write; a browser already bound to someone else is refused */
function claim_(device, fp, id, ua) {
  device = String(device || '').trim(); fp = String(fp || '').trim(); if ((!device && !fp) || !id) return;
  const dv = devices_(), row = deviceRow_(dv, device, fp);
  if (row) {
    if (String(row[dv.col.student_id]) !== id) throw new Error('This phone or browser is already registered to another batchmate. Each device can write as one person only; ask Vincent to reset it.');
    if (device && !String(row[dv.col.device_id] || '')) dv.sh.getRange(dv.rows.indexOf(row) + 2, dv.col.device_id + 1).setValue(device);
    if (fp && !String(row[dv.col.fingerprint] || '')) dv.sh.getRange(dv.rows.indexOf(row) + 2, dv.col.fingerprint + 1).setValue(fp);
    return;
  }
  const rec = {}; dv.headers.forEach(h => rec[h] = ''); rec.device_id = device; rec.fingerprint = fp; rec.student_id = id; rec.claimed_at = nowStamp_(); rec.user_agent = String(ua || '').slice(0, 200);
  dv.sh.appendRow(dv.headers.map(h => rec[h] === undefined ? '' : rec[h]));
}
/** fill in the message to the batch (only when none is on file yet), then register the browser */
function apiSaveMessage(id, text, device, ua, fp) {
  text = String(text || '').trim();
  if (!id || !text) throw new Error('Write your message to the batch first.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    claim_(device, fp, id, ua);
    const sc = ensureColumns_('Students', ['message_to_batch']);
    const st = table_('Students'); const i = st.rows.findIndex(r => String(r[st.col.student_id]) === id);
    if (i < 0) throw new Error('Unknown student: ' + id);
    if (String(st.rows[i][sc.message_to_batch] || '').trim()) throw new Error('A message to the batch is already on file for you.');
    st.sh.getRange(i + 2, sc.message_to_batch + 1).setValue(text);
    bust_(id);
    return { ok: true };
  } finally { lock.releaseLock(); }
}
/** append a testimonial written now: marked year 2026, one per writer/batchmate pair; the browser is registered to the writer */
function apiAddTestimonial(aboutId, byId, text, device, ua, fp) {
  text = String(text || '').trim();
  if (!byId) throw new Error('Pick who you are first.');
  if (!aboutId || !text) throw new Error('Pick who it is about and write the testimonial.');
  if (aboutId === byId) throw new Error('A testimonial is written about someone else.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    claim_(device, fp, byId, ua);
    const tc = ensureColumns_('Testimonials', ['year']);
    const st = table_('Students'), tt = table_('Testimonials');
    const names = {}; st.rows.forEach(x => names[String(x[st.col.student_id])] = String(x[st.col.display_name] || x[st.col.student_id]));
    if (!names[aboutId]) throw new Error('Unknown student: ' + aboutId);
    if (tt.rows.some(r => String(r[tt.col.by_id]) === byId && String(r[tt.col.about_id]) === aboutId)) throw new Error('You have already written to ' + names[aboutId] + (tt.rows.some(r => !isNew_(tt, r) && String(r[tt.col.by_id]) === byId && String(r[tt.col.about_id]) === aboutId) ? ', back in 2007.' : '.'));
    let max = 0; tt.rows.forEach(r => { const m = String(r[tt.col.testimonial_id] || '').match(/^T(\d+)$/); if (m) max = Math.max(max, Number(m[1])); });
    const tid = 'T' + String(max + 1).padStart(4, '0');
    const rec = {}; tt.headers.forEach(h => rec[h] = '');
    rec.testimonial_id = tid; rec.about_id = aboutId; rec.about_name = names[aboutId];
    rec.by_id = byId; rec.by_name_raw = names[byId] || byId;
    rec.text = text; rec.personal_note = ''; rec.source_file = 'web app · ' + nowStamp_(); rec.year = NEW_YEAR;
    tt.sh.appendRow(tt.headers.map(h => rec[h] === undefined ? '' : rec[h]));
    recountTestimonials(); bust_(aboutId); bust_(byId);
    return { id: tid, about: names[aboutId], by: rec.by_name_raw };
  } finally { lock.releaseLock(); }
}

/** write the chosen sections for one student; keys 1–4, only the ones sent are written (partial is fine) */
function apiSaveSections(id, choices) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const ys = table_('Year Sections'), st = table_('Students');
    const ls = table_('Lists'); const allowed = {};
    ls.rows.forEach(r => { const y = Number(r[ls.col.year_level]); if (y) allowed[y] = ls.headers.filter(h => /^section_/.test(h)).map(h => String(r[ls.col[h]] || '').trim()).filter(Boolean); });
    let written = 0;
    Object.keys(choices || {}).forEach(k => {
      const y = Number(k), v = String(choices[k] || '').trim();
      if (!y || !v) return;
      if (!allowed[y] || allowed[y].indexOf(v) < 0) throw new Error(v + ' is not a year ' + y + ' section.');
      let idx = ys.rows.findIndex(r => String(r[ys.col.student_id]) === id && Number(r[ys.col.year_level]) === y);
      if (idx >= 0) ys.sh.getRange(idx + 2, ys.col.section + 1).setValue(v);
      else {  // no row yet for this student/year: add one (full_name/year columns are formulas in the template; copy them down)
        const row = ys.sh.getLastRow() + 1;
        ys.sh.getRange(row, ys.col.student_id + 1).setValue(id);
        ys.sh.getRange(row, ys.col.year_level + 1).setValue(y);
        ys.sh.getRange(row, ys.col.section + 1).setValue(v);
        if (ys.col.full_name !== undefined) ys.sh.getRange(row, ys.col.full_name + 1).setFormula('=IFERROR(INDEX(Students!$B:$B,MATCH(A' + row + ',Students!$A:$A,0)),"")');
        if (ys.col.year !== undefined) ys.sh.getRange(row, ys.col.year + 1).setFormula('=IFERROR(INDEX(Lists!$B:$B,MATCH(C' + row + ',Lists!$A:$A,0)),"")');
      }
      written++;
      if (y === 4 && st.col.section !== undefined) {
        const i = st.rows.findIndex(r => String(r[st.col.student_id]) === id);
        if (i >= 0) st.sh.getRange(i + 2, st.col.section + 1).setValue(v);
      }
    });
    bust_(id);
    return { written };
  } finally { lock.releaseLock(); }
}

/* ================================================================== TEACHERS (main view)
   Teachers         teacher_code · first_name · last_name · nickname · display_name (a formula, never written here) · photo_file · photo_link · photo_url · drive_file_id · notes · gender (F → Ma'am, M → Sir)
   Teacher Subjects teacher_code · teacher_name (a formula) · subject · year_level (1–4) · entry_id · updated_at
                    one row per subject taught: a teacher can have many rows, at different year levels */
const T_SHEET = 'Teachers', TS_SHEET = 'Teacher Subjects', T_TTL = 120;
const T_EDITABLE = ['first_name', 'last_name', 'nickname', 'gender'];   /* gender: F → Ma'am, M → Sir */
function bustT_() { try { CACHE().removeAll(['tboot#n']); } catch (e) { } }

/** everything the Teachers view needs, in one small trip: it is the first thing the app asks for */
function apiTeachers() { return cached_('tboot', T_TTL, buildTeachers_); }
function buildTeachers_() {
  const t = table_(T_SHEET), s = subjectsTable_();
  const subs = {};
  s.rows.forEach((r, i) => {
    const code = String(r[s.col.teacher_code] || '').trim(), subject = String(r[s.col.subject] || '').trim();
    if (!code || !subject) return;
    (subs[code] = subs[code] || []).push({ id: String(r[s.col.entry_id] || ''), subject, year: Number(r[s.col.year_level]) || '' });
  });
  const g = (r, k) => t.col[k] === undefined ? '' : String(r[t.col[k]] == null ? '' : r[t.col[k]]).trim();
  const teachers = t.rows.filter(r => g(r, 'teacher_code')).map(r => {
    const code = g(r, 'teacher_code');
    return { code, first: g(r, 'first_name'), last: g(r, 'last_name'), nick: g(r, 'nickname'), gender: g(r, 'gender').toUpperCase().slice(0, 1), photo: g(r, 'photo_url'),
      subjects: (subs[code] || []).sort((a, b) => (a.year || 9) - (b.year || 9) || a.subject.localeCompare(b.subject)) };
  });
  /* who has already written to which teacher (names only, never the texts): the app greys them out */
  const tt = teacherTestis_(), wrote = {};
  tt.rows.forEach(r => { const c = String(r[tt.col.teacher_code] || ''), sid = String(r[tt.col.student_id] || ''); if (c && sid) (wrote[c] = wrote[c] || []).push(sid); });
  teachers.forEach(x => x.wrote = wrote[x.code] || []);
  /* a light list of the batch for "who is writing?", so the Teachers view never waits for the big batch download */
  const st = table_('Students');
  const students = st.rows.filter(r => r[st.col.student_id]).map(r => ({ id: String(r[st.col.student_id]), display: String(r[st.col.display_name] || r[st.col.first_name] || r[st.col.student_id]), name: String(r[st.col.full_name] || ''), photo: st.col.photo_url !== undefined ? String(r[st.col.photo_url] || '') : '' }));
  return { teachers, students, at: nowStamp_() };
}
/** the Teacher Testimonials sheet (made on first use) */
const TT_SHEET = 'Teacher Testimonials', TT_COLS = ['entry_id', 'teacher_code', 'teacher_name', 'student_id', 'student_name', 'text', 'written_at', 'device_id'];
function teacherTestis_() {
  let sh = SS().getSheetByName(TT_SHEET);
  if (!sh) { sh = SS().insertSheet(TT_SHEET); sh.getRange(1, 1, 1, TT_COLS.length).setValues([TT_COLS]).setFontWeight('bold').setBackground('#5e0f1b').setFontColor('#f6dfa0'); sh.setFrozenRows(1); }
  ensureColumns_(TT_SHEET, TT_COLS);
  return table_(TT_SHEET);
}
/** a batchmate's testimonial for a teacher: one per batchmate per teacher; the browser is registered to the writer as with the batch testimonials */
function apiAddTeacherTestimonial(code, studentId, text, device, ua, fp) {
  text = String(text || '').trim();
  if (!studentId) throw new Error('Pick who is writing first.');
  if (!text) throw new Error('Write the testimonial first.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const t = table_(T_SHEET), ti = t.rows.findIndex(r => String(r[t.col.teacher_code]) === code);
    if (ti < 0) throw new Error('Unknown teacher: ' + code);
    const st = table_('Students'), si = st.rows.findIndex(r => String(r[st.col.student_id]) === studentId);
    if (si < 0) throw new Error('Unknown batchmate: ' + studentId);
    claim_(device, fp, studentId, ua);
    const tt = teacherTestis_();
    const tr = t.rows[ti], tname = [tr[t.col.first_name], tr[t.col.last_name]].filter(Boolean).join(' ');
    if (tt.rows.some(r => String(r[tt.col.teacher_code]) === code && String(r[tt.col.student_id]) === studentId)) throw new Error('You have already left a testimonial for ' + tname + '. It can be written only once.');
    let max = 0; tt.rows.forEach(r => { const m = String(r[tt.col.entry_id] || '').match(/^TT(\d+)$/); if (m) max = Math.max(max, +m[1]); });
    const rec = {}; tt.headers.forEach(h => rec[h] = '');
    rec.entry_id = 'TT' + String(max + 1).padStart(4, '0'); rec.teacher_code = code; rec.teacher_name = tname;
    rec.student_id = studentId; rec.student_name = String(st.rows[si][st.col.full_name] || studentId); rec.text = text; rec.written_at = nowStamp_(); rec.device_id = String(device || fp || '');
    tt.sh.appendRow(tt.headers.map(h => rec[h] === undefined ? '' : rec[h]));
    SpreadsheetApp.flush(); bustT_();
    return { ok: true, teacher: tname };
  } finally { lock.releaseLock(); }
}
/** the Teacher Subjects sheet, after making sure every subject row has an entry_id (rows typed by hand get one too) */
function subjectsTable_() {
  ensureColumns_(TS_SHEET, ['teacher_code', 'teacher_name', 'subject', 'year_level', 'entry_id', 'updated_at']);
  const s = table_(TS_SHEET);
  let max = 0; s.rows.forEach(r => { const m = String(r[s.col.entry_id] || '').match(/^TS(\d+)$/); if (m) max = Math.max(max, +m[1]); });
  s.rows.forEach((r, i) => {
    if (String(r[s.col.teacher_code] || '').trim() && String(r[s.col.subject] || '').trim() && !String(r[s.col.entry_id] || '').trim()) {
      const id = 'TS' + String(++max).padStart(4, '0'); r[s.col.entry_id] = id; s.sh.getRange(i + 2, s.col.entry_id + 1).setValue(id);
    }
  });
  s.nextId = () => 'TS' + String(++max).padStart(4, '0');
  return s;
}
function teacherOut_(code) { const d = buildTeachers_(); return d.teachers.find(x => x.code === code) || null; }

/** save a teacher's names; only first_name, last_name and nickname are written (display_name is the sheet's formula) */
function apiSaveTeacher(code, fields) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    ensureColumns_(T_SHEET, T_EDITABLE);
    const t = table_(T_SHEET);
    const i = t.rows.findIndex(r => String(r[t.col.teacher_code]) === code);
    if (i < 0) throw new Error('Unknown teacher: ' + code);
    if (fields && fields.gender !== undefined && ['', 'F', 'M'].indexOf(String(fields.gender).toUpperCase()) < 0) throw new Error('Gender is F or M.');
    Object.keys(fields || {}).forEach(k => { if (T_EDITABLE.indexOf(k) >= 0) t.sh.getRange(i + 2, t.col[k] + 1).setValue(k === 'gender' ? String(fields[k] || '').toUpperCase() : String(fields[k] || '').trim()); });
    SpreadsheetApp.flush(); bustT_();
    return teacherOut_(code);
  } finally { lock.releaseLock(); }
}
function cleanYear_(y) { y = Number(y); if ([1, 2, 3, 4].indexOf(y) < 0) throw new Error('Pick the year level (I to IV).'); return y; }
/** add one subject taught (a teacher can have many, at different year levels) */
function apiAddTeacherSubject(code, subject, year) {
  subject = String(subject || '').trim(); if (!subject) throw new Error('Type the subject.'); year = cleanYear_(year);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const t = table_(T_SHEET); if (!t.rows.some(r => String(r[t.col.teacher_code]) === code)) throw new Error('Unknown teacher: ' + code);
    const s = subjectsTable_();
    if (s.rows.some(r => String(r[s.col.teacher_code]) === code && String(r[s.col.subject]).trim().toLowerCase() === subject.toLowerCase() && Number(r[s.col.year_level]) === year))
      throw new Error('Already listed: ' + subject + ', year ' + year + '.');
    /* next free row by the teacher_code column (teacher_name is a formula that fills its whole column, so it is never written) */
    const codes = s.sh.getRange(1, s.col.teacher_code + 1, s.sh.getMaxRows(), 1).getValues();
    let row = codes.length; while (row > 1 && !String(codes[row - 1][0]).trim()) row--; row++;
    if (row > s.sh.getMaxRows()) s.sh.insertRowsAfter(s.sh.getMaxRows(), 20);
    const id = s.nextId();
    s.sh.getRange(row, s.col.teacher_code + 1).setValue(code);
    s.sh.getRange(row, s.col.subject + 1).setValue(subject);
    s.sh.getRange(row, s.col.year_level + 1).setValue(year);
    s.sh.getRange(row, s.col.entry_id + 1).setValue(id);
    s.sh.getRange(row, s.col.updated_at + 1).setValue(nowStamp_());
    SpreadsheetApp.flush(); bustT_();
    return teacherOut_(code);
  } finally { lock.releaseLock(); }
}
/** change one subject row (its subject and/or year level) */
function apiUpdateTeacherSubject(entryId, subject, year) {
  subject = String(subject || '').trim(); if (!subject) throw new Error('Type the subject.'); year = cleanYear_(year);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const s = subjectsTable_(); const i = s.rows.findIndex(r => String(r[s.col.entry_id]) === entryId);
    if (i < 0) throw new Error('That subject row is gone (someone may have removed it). Reload to see the latest.');
    s.sh.getRange(i + 2, s.col.subject + 1).setValue(subject);
    s.sh.getRange(i + 2, s.col.year_level + 1).setValue(year);
    s.sh.getRange(i + 2, s.col.updated_at + 1).setValue(nowStamp_());
    SpreadsheetApp.flush(); bustT_();
    return teacherOut_(String(s.rows[i][s.col.teacher_code]));
  } finally { lock.releaseLock(); }
}
/** remove one subject row */
function apiDeleteTeacherSubject(entryId) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const s = subjectsTable_(); const i = s.rows.findIndex(r => String(r[s.col.entry_id]) === entryId);
    if (i < 0) throw new Error('That subject row is already gone.');
    const code = String(s.rows[i][s.col.teacher_code]);
    s.sh.deleteRow(i + 2);
    SpreadsheetApp.flush(); bustT_();
    return teacherOut_(code);
  } finally { lock.releaseLock(); }
}
/** teacher portraits as data, for visitors whose browser can't open the Drive link (same idea as apiPhotos).
    Photoshop files can't be shown in a browser, so Drive's own preview picture of the file is sent instead. */
function apiTeacherPhotos(codes) {
  codes = (codes || []).slice(0, 20);
  const t = table_(T_SHEET), fid = {};
  t.rows.forEach(r => fid[String(r[t.col.teacher_code])] = t.col.drive_file_id === undefined ? '' : String(r[t.col.drive_file_id] || ''));
  const out = {}, c = CACHE();
  let hits = {}; try { hits = c.getAll(codes.map(k => 'img:' + fid[k])); } catch (e) { }
  codes.forEach(k => {
    const f = fid[k]; if (!f) { out[k] = ''; return; }
    if (hits['img:' + f]) { out[k] = hits['img:' + f]; return; }
    try {
      const file = DriveApp.getFileById(f), type = file.getMimeType();
      const blob = /^image\/(jpeg|png|gif|webp)$/.test(type) ? file.getBlob() : file.getThumbnail();
      const url = 'data:' + (blob.getContentType() || 'image/jpeg') + ';base64,' + Utilities.base64Encode(blob.getBytes());
      out[k] = url; if (url.length < 99000) try { c.put('img:' + f, url, PHOTO_TTL); } catch (e) { }
    } catch (e) { out[k] = ''; }
  });
  return out;
}

