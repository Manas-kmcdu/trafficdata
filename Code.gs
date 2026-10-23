/**
 * POV by Unrest: shared backend v2.
 * Stores: opinions + reactions, poll votes, food outlets + ratings,
 * and CONFIG (polls + topics live in sheet, no rebuild to add one).
 *
 * Deploy: Deploy > New deployment > Web app. Execute as Me. Access Anyone.
 * After any edit: Deploy > Manage deployments > edit > New version.
 *
 * Sheets (auto-created):
 *   opinions  id|topicId|text|campus|status|createdAt|clientId|reports|gender
 *   reactions opinionId|clientId|at
 *   votes     pollId|optionId|collegeId|campus|gender|clientId|at
 *   outlets   id|campus|name|status|createdAt|clientId
 *   ratings   outletId|clientId|stars|at
 *   polls     id|opensAt|categoryId|spicy|question|options   (options split by |)
 *   topics    id|title|blurb|moderated|opensAt|closesAt       (dates: ms or date cell)
 *
 * Moderation: opinions on moderated topics start "review". Set "live" to publish,
 * "hidden" to remove. Same for outlets.status.
 * Privacy: no names/emails. clientId = random per-browser id (anti-dupe only).
 */

const MIN_LEN = 20, MAX_LEN = 280, MAX_PER_TOPIC = 3, REPORT_HIDE_AT = 3;
const MIN_RESULTS = 30;                 // hide counts until this many votes in a slice
const MAX_OUTLETS_PER_CAMPUS = 3;
const FOOD_MODERATED = false;           // true: new outlets start "review"
const GENDERS = ['m', 'f', 'o'];
const CAMPUSES = ['north', 'south', 'off'];

const DAY = 864e5;
const HEAD = {
  opinions:  ['id', 'topicId', 'text', 'campus', 'status', 'createdAt', 'clientId', 'reports', 'gender'],
  reactions: ['opinionId', 'clientId', 'at'],
  votes:     ['pollId', 'optionId', 'collegeId', 'campus', 'gender', 'clientId', 'at'],
  outlets:   ['id', 'campus', 'name', 'status', 'createdAt', 'clientId'],
  ratings:   ['outletId', 'clientId', 'stars', 'at'],
  polls:     ['id', 'opensAt', 'categoryId', 'spicy', 'question', 'options'],
  topics:    ['id', 'title', 'blurb', 'moderated', 'opensAt', 'closesAt'],
};
function seed_(name) {
  const n = Date.now();
  if (name === 'topics') return [
    ['other-gender', 'My Message to the Other Gender', "Say what you'd want them to hear. Specific, constructive and kind. Reviewed before anything goes public.", true, n, n + 7 * DAY],
    ['unpopular-college', 'Unpopular Opinions About College Life', "The take you'd only say in a group chat of three. Anonymous, and no names.", false, n + 7 * DAY, n + 14 * DAY],
    ['twenties', 'Things Nobody Warns You About in Your 20s', 'What you wish someone had told you before it happened.', false, n + 14 * DAY, n + 21 * DAY],
    ['friends-understood', 'Things You Wish Your College Friends Understood', "The thing you haven't said out loud yet.", false, n + 21 * DAY, n + 28 * DAY],
    ['society-politics', 'Society Politics Nobody Talks About', 'How things really work inside societies. No names, no colleges named.', false, n + 28 * DAY, n + 35 * DAY],
  ];
  if (name === 'polls') return [
    ['dating-hard', '2026-10-01', 'dating', true, 'How hard is dating in college?', 'Easier than people say.|Manageable.|Pretty hard.|Nearly impossible.'],
    ['societies-worth', '2026-10-30', 'college', false, 'Are college societies worth it?', 'Absolutely.|Depends on the society.|Not really.|A waste of time.'],
    ['course-vs-college', '2026-11-06', 'college', false, 'Course over college, or the other way around?', 'Course, always.|Leaning course.|Leaning college.|College, always.'],
  ];
  return null;
}

// ---------- helpers ----------
function sh_(name) {
  const ss = SpreadsheetApp.getActive(), head = HEAD[name];
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.appendRow(head);
    s.setFrozenRows(1);
  } else if (s.getLastColumn() < head.length) {
    s.getRange(1, 1, 1, head.length).setValues([head]);   // migrate old sheet
  }
  // Seed config sheets even when they already exist but contain only a header.
  if (s.getLastRow() <= 1) (seed_(name) || []).forEach(r => s.appendRow(r));
  return s;
}
function rows_(name) { return sh_(name).getDataRange().getValues().slice(1); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function norm_(s) { return String(s).toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim(); }
function ms_(v) {
  if (v instanceof Date) return v.getTime();
  if (v === '' || v === null) return 0;
  const n = Number(v); if (!isNaN(n)) return n;
  const d = Date.parse(v); return isNaN(d) ? 0 : d;
}
function bool_(v) { return v === true || String(v).toLowerCase() === 'true'; }
function topic_(id) {
  const r = rows_('topics').filter(x => x[0] === id)[0];
  return r ? { id: r[0], title: r[1], blurb: r[2], moderated: bool_(r[3]), opensAt: ms_(r[4]), closesAt: ms_(r[5]) } : null;
}

// ---------- GET ----------
function doGet(e) {
  const p = e.parameter || {}, a = p.action;
  if (a === 'list')    return json_(list_(String(p.topic || ''), String(p.client || '')));
  if (a === 'topics')  return json_(topics_());
  if (a === 'polls')   return json_(polls_());
  if (a === 'results') return json_(results_(p));
  if (a === 'food')    return json_(food_(String(p.client || '')));
  return json_({ ok: true });
}

function topics_() {
  return { ok: true, topics: rows_('topics').filter(r => r[0]).map(r => ({
    id: r[0], title: r[1], blurb: r[2], moderated: bool_(r[3]), opensAt: ms_(r[4]), closesAt: ms_(r[5]) })) };
}
function polls_() {
  return { ok: true, polls: rows_('polls').filter(r => r[0]).map(r => ({
    id: r[0], opensAt: ms_(r[1]), categoryId: r[2], spicy: bool_(r[3]), question: r[4],
    options: String(r[5]).split('|').map(s => s.trim()).filter(Boolean) })) };
}

function list_(topicId, clientId) {
  const count = {}, mine = {};
  rows_('reactions').forEach(r => {
    count[r[0]] = (count[r[0]] || 0) + 1;
    if (clientId && r[1] === clientId) mine[r[0]] = true;
  });
  return { ok: true, opinions: rows_('opinions').filter(r => r[1] === topicId && r[4] === 'live').map(r => ({
    id: r[0], text: r[2], campus: r[3], at: ms_(r[5]), count: count[r[0]] || 0, mine: !!mine[r[0]], gender: r[8] || '' })) };
}

// results: ?action=results&poll=ID&scope=all|campus|college|gender&campus=&college=&gender=
function results_(p) {
  const pollId = String(p.poll || ''), scope = String(p.scope || 'all');
  let v = rows_('votes').filter(r => r[0] === pollId);
  if (scope === 'campus')  v = v.filter(r => r[3] === p.campus);
  if (scope === 'college') v = v.filter(r => r[2] === p.college);
  if (scope === 'gender')  v = v.filter(r => r[4] === p.gender);
  return slice_(v);
}
function slice_(v) {
  const total = v.length, counts = {};
  if (total >= MIN_RESULTS) v.forEach(r => counts[r[1]] = (counts[r[1]] || 0) + 1);  // hide small slices
  return { ok: true, total: total, counts: counts };
}

function food_(client) {
  const rt = rows_('ratings'), hist = {}, mine = {};
  rt.forEach(r => {
    (hist[r[0]] = hist[r[0]] || [0, 0, 0, 0, 0])[Number(r[2]) - 1]++;
    if (client && r[1] === client) mine[r[0]] = Number(r[2]);
  });
  return { ok: true, outlets: rows_('outlets').filter(r => r[3] === 'live').map(r => ({
    id: r[0], campus: r[1], name: r[2], createdAt: ms_(r[4]),
    hist: hist[r[0]] || [0, 0, 0, 0, 0], mine: mine[r[0]] || 0, addedByMe: !!client && r[5] === client })) };
}

// ---------- POST ----------
function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (_) { return json_({ ok: false, message: 'Bad request.' }); }
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const fn = { submit: submit_, react: react_, report: report_, vote: vote_, addOutlet: addOutlet_, rate: rate_ }[b.action];
    return json_(fn ? fn(b) : { ok: false, message: 'Unknown action.' });
  } finally { lock.releaseLock(); }
}

function submit_(b) {
  const t = topic_(b.topicId), now = Date.now();
  if (!t || now < t.opensAt || now >= t.closesAt) return { ok: false, message: "That topic isn't open." };
  const text = String(b.text || '').trim().replace(/\s+/g, ' ');
  if (text.length < MIN_LEN || text.length > MAX_LEN) return { ok: false, message: 'Opinion must be ' + MIN_LEN + '-' + MAX_LEN + ' characters.' };
  if (CAMPUSES.indexOf(b.campus) < 0) return { ok: false, message: 'Missing campus.' };
  if (GENDERS.indexOf(b.gender) < 0) return { ok: false, message: 'Choose your gender.' };
  const gender = b.gender;
  const clientId = String(b.clientId || '');
  if (!clientId) return { ok: false, message: 'Missing client.' };
  const s = sh_('opinions');
  const rows = rows_('opinions').filter(r => r[1] === b.topicId);
  if (rows.filter(r => r[6] === clientId).length >= MAX_PER_TOPIC) return { ok: false, message: "You've shared " + MAX_PER_TOPIC + ' takes on this topic.' };
  if (rows.some(r => norm_(r[2]) === norm_(text))) return { ok: false, message: 'Someone already said exactly that.' };
  const id = Utilities.getUuid().slice(0, 8), status = t.moderated ? 'review' : 'live';
  s.appendRow([id, b.topicId, /^[=+\-@]/.test(text) ? "'" + text : text, b.campus, status, now, clientId, 0, gender]);
  return { ok: true, id: id, status: status };
}

function react_(b) {
  const o = String(b.opinionId || ''), c = String(b.clientId || '');
  if (!o || !c) return { ok: false };
  const s = sh_('reactions'), rows = s.getDataRange().getValues();
  let f = -1;
  for (let i = 1; i < rows.length; i++) if (rows[i][0] === o && rows[i][1] === c) { f = i + 1; break; }
  if (b.on && f < 0) s.appendRow([o, c, Date.now()]);
  if (!b.on && f > 0) s.deleteRow(f);
  return { ok: true };
}

function report_(b) {
  const s = sh_('opinions'), rows = s.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) if (rows[i][0] === b.opinionId) {
    const n = Number(rows[i][7] || 0) + 1;
    s.getRange(i + 1, 8).setValue(n);
    if (n >= REPORT_HIDE_AT && rows[i][4] === 'live') s.getRange(i + 1, 5).setValue('hidden');
    break;
  }
  return { ok: true };
}

// vote: {action:'vote', pollId, optionId, collegeId, campus, gender, clientId}
function vote_(b) {
  const pollId = String(b.pollId || ''), optionId = String(b.optionId || ''), clientId = String(b.clientId || '');
  const poll = polls_().polls.filter(p => p.id === pollId)[0];
  if (!poll || Date.now() < poll.opensAt) return { ok: false, message: "That poll isn't open." };
  if (!/^[a-z]$/.test(optionId) || optionId.charCodeAt(0) - 97 >= poll.options.length) return { ok: false, message: 'Bad option.' };
  if (!clientId) return { ok: false, message: 'Missing client.' };
  if (!/^[a-z0-9-]{1,30}$/.test(String(b.collegeId || ''))) return { ok: false, message: 'Pick your college first.' };
  if (CAMPUSES.indexOf(b.campus) < 0) return { ok: false, message: 'Missing campus.' };
  if (GENDERS.indexOf(b.gender) < 0) return { ok: false, message: 'Choose your gender.' };
  const gender = b.gender;
  const dupe = rows_('votes').some(r => r[0] === pollId && r[5] === clientId);
  if (!dupe) sh_('votes').appendRow([pollId, optionId, b.collegeId, b.campus, gender, clientId, Date.now()]);
  const res = results_({ poll: pollId, scope: 'all' });
  res.already = dupe;
  return res;
}

// addOutlet: {action:'addOutlet', campus, name, userCampus, clientId}
function addOutlet_(b) {
  const clientId = String(b.clientId || ''), campus = b.campus;
  if (!clientId || (campus !== 'north' && campus !== 'south')) return { ok: false, message: 'Bad request.' };
  if (b.userCampus !== campus && b.userCampus !== 'off') return { ok: false, message: 'Only ' + campus + ' campus students can add ' + campus + ' outlets.' };
  const name = String(b.name || '').trim().replace(/\s+/g, ' ');
  if (name.length < 2 || name.length > 40) return { ok: false, message: 'Name must be 2-40 characters.' };
  if (/https?:|www\.|\.com|@/i.test(name) || !/^[\p{L}\p{N} '&.,()\-]+$/u.test(name)) return { ok: false, message: 'Letters and numbers only. No links or handles.' };
  const all = rows_('outlets').filter(r => r[1] === campus && r[3] !== 'hidden');
  if (all.filter(r => r[5] === clientId).length >= MAX_OUTLETS_PER_CAMPUS) return { ok: false, message: "You've added " + MAX_OUTLETS_PER_CAMPUS + ' outlets here already. Rate a few instead.' };
  const key = s => norm_(s).split(' ').sort().join(' ');
  const hit = all.filter(r => key(r[2]) === key(name))[0];
  if (hit) return { ok: false, message: '"' + name + '" is already listed as "' + hit[2] + '". Rate that one instead.', matchId: hit[0] };
  const id = 'u-' + campus[0] + '-' + Utilities.getUuid().slice(0, 6);
  sh_('outlets').appendRow([id, campus, name, FOOD_MODERATED ? 'review' : 'live', Date.now(), clientId]);
  return { ok: true, id: id, status: FOOD_MODERATED ? 'review' : 'live' };
}

// rate: {action:'rate', outletId, stars, userCampus, clientId}
function rate_(b) {
  const clientId = String(b.clientId || ''), stars = Number(b.stars);
  if (!clientId || !(stars >= 1 && stars <= 5) || stars % 1) return { ok: false, message: 'Pick 1 to 5 stars.' };
  const o = rows_('outlets').filter(r => r[0] === b.outletId && r[3] === 'live')[0];
  if (!o) return { ok: false, message: "That outlet isn't here anymore." };
  if (b.userCampus !== o[1] && b.userCampus !== 'off') return { ok: false, message: 'Only ' + o[1] + ' campus students can rate this.' };
  if (o[5] === clientId) return { ok: false, message: "You can't rate an outlet you added." };
  const s = sh_('ratings'), rows = s.getDataRange().getValues();
  let f = -1;
  for (let i = 1; i < rows.length; i++) if (rows[i][0] === b.outletId && rows[i][1] === clientId) { f = i + 1; break; }
  if (f > 0) s.getRange(f, 3, 1, 2).setValues([[stars, Date.now()]]);
  else s.appendRow([b.outletId, clientId, stars, Date.now()]);
  return { ok: true };
}