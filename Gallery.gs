/**
 * בנק הזמנות — שלב 2: גלריה, לבבות, תגובות, צפיות ופרופילים.
 * צילומי המסך נשמרים בתיקיית Drive פרטית ומוגשים ללקוח דרך השרת (getThumbs),
 * כך שאין צורך לשתף את הקבצים ואין תלות בעוגיות של צד שלישי.
 */

var GALLERY_LIMITS = { title: 60, description: 300, comment: 200, suggestion: 300, bio: 120 };
var MAX_THUMB_CHARS = 300000; // אורך מקסימלי של תמונת data URL שמגיעה מהלקוח

// ───────────── עזרים ─────────────

/** מזהה ציבורי לא הפיך של משתמשת (כדי לא לחשוף מיילים ללקוח) */
function uid_(email) {
  var d = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, 'bank-hz:' + String(email).toLowerCase().trim());
  return d.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('').slice(0, 12);
}

function genId_(prefix) { return prefix + new Date().getTime().toString(36) + Math.random().toString(36).slice(2, 6); }

function findBy_(t, key, val, msg) {
  var r = t.rows.filter(function (x) { return String(x[key]) === String(val); })[0];
  if (!r) throw new Error(msg || 'לא נמצא. רענני את הדף.');
  return r;
}

function visibleProject_(p, me) {
  return me.role === 'teacher' || !me.grade || str_(p.grade) === me.grade;
}

function deleteRowsWhere_(name, pred) {
  var t = table_(name);
  for (var i = t.rows.length - 1; i >= 0; i--) if (pred(t.rows[i])) t.sh.deleteRow(t.rows[i]._row);
}

/** סופר לבבות/תגובות/צפיות בבת אחת */
function aggregate_() {
  var agg = { hearts: {}, comments: {}, views: {}, heartedBy: {}, givenBy: {}, receivedBy: {}, commentsBy: {}, authorOf: {} };
  table_('Projects').rows.forEach(function (p) { agg.authorOf[String(p.id)] = String(p.authorEmail).toLowerCase(); });
  table_('Hearts').rows.forEach(function (r) {
    var pid = String(r.projectId), em = String(r.fromEmail).toLowerCase();
    agg.hearts[pid] = (agg.hearts[pid] || 0) + 1;
    agg.heartedBy[pid + '|' + em] = true;
    agg.givenBy[em] = (agg.givenBy[em] || 0) + 1;
    var au = agg.authorOf[pid];
    if (au) agg.receivedBy[au] = (agg.receivedBy[au] || 0) + 1;
  });
  table_('Comments').rows.forEach(function (r) {
    var pid = String(r.projectId), em = String(r.fromEmail).toLowerCase();
    if (str_(r.text)) agg.comments[pid] = (agg.comments[pid] || 0) + 1; // מונה ציבורי: רק משובים גלויים
    agg.commentsBy[em] = (agg.commentsBy[em] || 0) + 1;
  });
  table_('Views').rows.forEach(function (r) { var pid = String(r.projectId); agg.views[pid] = (agg.views[pid] || 0) + 1; });
  return agg;
}

function pubProject_(p, me, agg) {
  var id = String(p.id), authorEmail = String(p.authorEmail).toLowerCase();
  var mine = authorEmail === me.email;
  return {
    id: id, createdAt: str_(p.createdAt), title: str_(p.title), description: str_(p.description), link: str_(p.link),
    authorName: str_(p.authorName), authorUid: uid_(authorEmail), grade: str_(p.grade),
    hasThumb: !!str_(p.thumbFileId), pitch: parsePitch_(p.pitchPrep),
    hearts: agg.hearts[id] || 0, comments: agg.comments[id] || 0, views: agg.views[id] || 0,
    heartedByMe: !!agg.heartedBy[id + '|' + me.email],
    mine: mine, canDelete: mine || me.role === 'teacher'
  };
}

function galleryFor_(me) {
  var agg = aggregate_();
  var list = table_('Projects').rows
    .filter(function (p) { return visibleProject_(p, me); })
    .map(function (p) { return pubProject_(p, me, agg); })
    .sort(function (a, b) { return a.createdAt < b.createdAt ? 1 : -1; });
  return { projects: list, stats: { given: agg.givenBy[me.email] || 0, received: agg.receivedBy[me.email] || 0 } };
}

var PITCH_FIELDS = ['hook', 'client', 'demo', 'feedback', 'learned', 'question'];

function parsePitch_(raw) {
  try {
    var o = JSON.parse(str_(raw) || '{}'), out = {};
    PITCH_FIELDS.forEach(function (k) { out[k] = String(o[k] || ''); });
    return out;
  } catch (e) { return {}; }
}

/** משובים לתוצר. "הצעות לשיפור" מוצגות רק ליוצרת התוצר, לכותבת ולמורה. */
function commentsFor_(pid, me, proj) {
  var isAuthor = !!proj && String(proj.authorEmail).toLowerCase() === me.email;
  var out = [];
  table_('Comments').rows
    .filter(function (c) { return String(c.projectId) === String(pid); })
    .sort(function (a, b) { return str_(a.createdAt) < str_(b.createdAt) ? -1 : 1; })
    .forEach(function (c) {
      var writer = String(c.fromEmail).toLowerCase() === me.email;
      var seesPrivate = isAuthor || writer || me.role === 'teacher';
      var item = { id: String(c.id), fromName: str_(c.fromName), fromUid: uid_(c.fromEmail), text: str_(c.text), createdAt: str_(c.createdAt),
                   canDelete: writer || me.role === 'teacher' };
      if (seesPrivate && str_(c.suggestion)) item.suggestion = str_(c.suggestion);
      if (item.text || item.suggestion) out.push(item);
    });
  return out;
}

// ───────────── קריאה ─────────────

function getGallery() {
  return galleryFor_(requireUser_());
}

/** פתיחת תוצר: מחזיר תגובות ורושם צפייה (פעם אחת לכל משתמשת, לא של היוצרת עצמה) */
function getProject(id) {
  return withLock_(function () {
    var me = requireUser_();
    var p = findBy_(table_('Projects'), 'id', id, 'התוצר לא נמצא (אולי נמחק).');
    if (!visibleProject_(p, me)) throw new Error('התוצר שייך לקבוצה אחרת.');
    if (String(p.authorEmail).toLowerCase() !== me.email) {
      var v = table_('Views');
      var seen = v.rows.some(function (r) { return String(r.projectId) === String(id) && String(r.email).toLowerCase() === me.email; });
      if (!seen) appendRow_(v, { projectId: id, email: me.email, createdAt: new Date().toISOString() });
    }
    return { project: pubProject_(p, me, aggregate_()), comments: commentsFor_(id, me, p) };
  });
}

/** תמונות ממוזערות כ-data URL, עד 10 בקריאה. נשמרות ב-Cache ל-6 שעות. */
function getThumbs(ids) {
  var me = requireUser_();
  var out = {};
  var cache = CacheService.getScriptCache();
  var t = table_('Projects');
  (ids || []).slice(0, 10).forEach(function (id) {
    try {
      var p = findBy_(t, 'id', id);
      var fid = str_(p.thumbFileId);
      if (!fid || !visibleProject_(p, me)) return;
      var key = 'th_' + fid;
      var data = cache.get(key);
      if (!data) {
        data = 'data:image/jpeg;base64,' + Utilities.base64Encode(DriveApp.getFileById(fid).getBlob().getBytes());
        try { cache.put(key, data, 21600); } catch (e) { /* גדול מדי למטמון — לא נורא */ }
      }
      out[id] = data;
    } catch (e) { /* תמונה חסרה — מדלגים */ }
  });
  return out;
}

function getProfile(uid) {
  var me = requireUser_();
  var u = table_('Users').rows.filter(function (r) { return uid_(r.email) === uid; })[0];
  if (!u) throw new Error('הפרופיל לא נמצא.');
  var ug = normGrade_(u.grade);
  if (me.role !== 'teacher' && me.grade && ug && ug !== me.grade) throw new Error('הפרופיל שייך לקבוצה אחרת.');
  var email = String(u.email).toLowerCase();
  var agg = aggregate_();
  var projects = table_('Projects').rows
    .filter(function (p) { return String(p.authorEmail).toLowerCase() === email && visibleProject_(p, me); })
    .map(function (p) { return pubProject_(p, me, agg); })
    .sort(function (a, b) { return a.createdAt < b.createdAt ? 1 : -1; });
  return {
    uid: uid, name: String(u.firstName).trim(), grade: ug, bio: str_(u.bio), mine: email === me.email,
    stats: { projects: projects.length, heartsReceived: agg.receivedBy[email] || 0, heartsGiven: agg.givenBy[email] || 0, commentsGiven: agg.commentsBy[email] || 0 },
    projects: projects
  };
}

// ───────────── כתיבה ─────────────

/** data: {title, description, link, thumb (data URL של JPEG, לא חובה), orderId (לא חובה)} */
function createProject(data) {
  return withLock_(function () {
    var me = requireUser_();
    data = data || {};
    var link = cleanLink_(data.link);
    if (!link) throw new Error('חסר: קישור לתוצר');
    var p = {
      id: genId_('p'), createdAt: new Date().toISOString(),
      authorEmail: me.email, authorName: me.firstName, grade: me.grade,
      title: clean_(data.title, GALLERY_LIMITS.title, 'כותרת', true),
      description: clean_(data.description, GALLERY_LIMITS.description, 'תיאור', true),
      link: link, thumbFileId: '', orderId: ''
    };
    if (data.thumb) p.thumbFileId = saveThumb_(data.thumb, p.id);
    // קישור להזמנה: אם מי שמפרסמת קשורה להזמנה, הקישור נשמר גם שם
    if (data.orderId) {
      var ot = table_('Orders');
      var o = ot.rows.filter(function (r) { return String(r.id) === String(data.orderId); })[0];
      if (o) {
        var emails = splitList_(o.adoptedEmails).map(function (e) { return e.toLowerCase(); });
        if (emails.indexOf(me.email) >= 0 || String(o.bringerEmail).toLowerCase() === me.email) {
          p.orderId = String(o.id);
          if (!str_(o.productLink)) { o.productLink = link; writeRow_(ot, o); }
        }
      }
    }
    appendRow_(table_('Projects'), p);
    return galleryFor_(me);
  });
}

function saveThumb_(dataUrl, pid) {
  var s = String(dataUrl);
  if (s.length > MAX_THUMB_CHARS) throw new Error('התמונה גדולה מדי.');
  var m = s.match(/^data:image\/jpeg;base64,([A-Za-z0-9+\/=]+)$/);
  if (!m) throw new Error('התמונה חייבת להיות בפורמט JPEG.');
  var bytes = Utilities.base64Decode(m[1]);
  if (bytes.length < 4 || (bytes[0] & 0xff) !== 0xff || (bytes[1] & 0xff) !== 0xd8) throw new Error('קובץ התמונה לא תקין.');
  try {
    var file = DriveApp.getFolderById(DRIVE_FOLDER_ID).createFile(Utilities.newBlob(bytes, 'image/jpeg', pid + '.jpg'));
    return file.getId();
  } catch (e) {
    throw new Error('לא הצלחנו לשמור את צילום המסך. אפשר לפרסם בלי תמונה ולבקש מהמורה לבדוק את התיקייה.');
  }
}

function deleteProject(id) {
  return withLock_(function () {
    var me = requireUser_();
    var p = findBy_(table_('Projects'), 'id', id, 'התוצר לא נמצא.');
    if (String(p.authorEmail).toLowerCase() !== me.email && me.role !== 'teacher') throw new Error('אפשר למחוק רק תוצר שפרסמת.');
    if (str_(p.thumbFileId)) { try { DriveApp.getFileById(str_(p.thumbFileId)).setTrashed(true); } catch (e) { } }
    ['Hearts', 'Comments', 'Views'].forEach(function (n) { deleteRowsWhere_(n, function (r) { return String(r.projectId) === String(id); }); });
    var t = table_('Projects');
    t.sh.deleteRow(findBy_(t, 'id', id)._row);
    return galleryFor_(me);
  });
}

/** לב: לחיצה מוסיפה, לחיצה נוספת מסירה. אי אפשר לתת לב לתוצר של עצמך. */
function toggleHeart(id) {
  return withLock_(function () {
    var me = requireUser_();
    var p = findBy_(table_('Projects'), 'id', id, 'התוצר לא נמצא.');
    if (!visibleProject_(p, me)) throw new Error('התוצר שייך לקבוצה אחרת.');
    if (String(p.authorEmail).toLowerCase() === me.email) throw new Error('אי אפשר לתת לב לתוצר של עצמך.');
    var ht = table_('Hearts');
    var mine = ht.rows.filter(function (r) { return String(r.projectId) === String(id) && String(r.fromEmail).toLowerCase() === me.email; })[0];
    if (mine) ht.sh.deleteRow(mine._row);
    else appendRow_(ht, { projectId: id, fromEmail: me.email, createdAt: new Date().toISOString() });
    var count = ht.rows.filter(function (r) { return String(r.projectId) === String(id); }).length + (mine ? -1 : 1);
    return { hearted: !mine, hearts: count };
  });
}

/** משוב: liked = "מה אהבתי" (גלוי לכולן), suggestion = "הצעות לשיפור" (רק ליוצרת). לפחות אחד מהם. */
function addComment(id, liked, suggestion) {
  return withLock_(function () {
    var me = requireUser_();
    var p = findBy_(table_('Projects'), 'id', id, 'התוצר לא נמצא.');
    if (!visibleProject_(p, me)) throw new Error('התוצר שייך לקבוצה אחרת.');
    var text = clean_(liked, GALLERY_LIMITS.comment, 'מה אהבתי', false);
    var sug = clean_(suggestion, GALLERY_LIMITS.suggestion, 'הצעות לשיפור', false);
    if (!text && !sug) throw new Error('כתבי לפחות דבר אחד: מה אהבת או הצעה לשיפור.');
    appendRow_(table_('Comments'), {
      id: genId_('c'), projectId: id, fromEmail: me.email, fromName: me.firstName,
      text: text, suggestion: sug, createdAt: new Date().toISOString()
    });
    return commentsFor_(id, me, p);
  });
}

function deleteComment(commentId, projectId) {
  return withLock_(function () {
    var me = requireUser_();
    var t = table_('Comments');
    var c = findBy_(t, 'id', commentId, 'התגובה לא נמצאה.');
    if (String(c.fromEmail).toLowerCase() !== me.email && me.role !== 'teacher') throw new Error('אפשר למחוק רק תגובה שכתבת.');
    t.sh.deleteRow(c._row);
    return commentsFor_(projectId, me, findBy_(table_('Projects'), 'id', projectId));
  });
}

/** יוצרת עורכת את התוצר שלה: כותרת, תיאור, קישור, וצילום מסך חדש (לא חובה). */
function editMyProject(id, f) {
  return withLock_(function () {
    var me = requireUser_();
    f = f || {};
    var t = table_('Projects');
    var p = findBy_(t, 'id', id, 'התוצר לא נמצא.');
    if (String(p.authorEmail).toLowerCase() !== me.email && me.role !== 'teacher') throw new Error('אפשר לערוך רק תוצר שפרסמת.');
    if (f.title !== undefined) p.title = clean_(f.title, GALLERY_LIMITS.title, 'כותרת', true);
    if (f.description !== undefined) p.description = clean_(f.description, GALLERY_LIMITS.description, 'תיאור', true);
    if (f.link !== undefined) { var l = cleanLink_(f.link); if (!l) throw new Error('חסר: קישור לתוצר'); p.link = l; }
    if (f.thumb) {
      var old = str_(p.thumbFileId);
      p.thumbFileId = saveThumb_(f.thumb, p.id + '-' + new Date().getTime());
      if (old) { try { DriveApp.getFileById(old).setTrashed(true); } catch (e) { } }
    }
    writeRow_(t, p);
    return galleryFor_(me);
  });
}

/** דף הכנה לפיץ': שדות קצרים שנשמרים כ-JSON. o: {hook, demo, feedback, learned, question} */
function savePitchPrep(id, o) {
  return withLock_(function () {
    var me = requireUser_();
    o = o || {};
    var t = table_('Projects');
    var p = findBy_(t, 'id', id, 'התוצר לא נמצא.');
    if (String(p.authorEmail).toLowerCase() !== me.email && me.role !== 'teacher') throw new Error('רק היוצרת ממלאת את דף ההכנה.');
    var out = {};
    PITCH_FIELDS.forEach(function (k) { out[k] = clean_(o[k], 200, 'שדה בדף ההכנה', false); });
    p.pitchPrep = JSON.stringify(out);
    writeRow_(t, p);
    return { pitch: out };
  });
}

// ───────────── הצגה בכיתה ("מוצג עכשיו") ─────────────

/** מצב ההצגה {p: מזהה התוצר המוצג, q: סדר ההצגות}. נשמר במטמון כדי שבדיקה כל כמה שניות לא תקרא את הגיליון. */
function presState_() {
  var cache = CacheService.getScriptCache();
  var raw = cache.get('pres2');
  if (raw) { try { return JSON.parse(raw); } catch (e) { } }
  var cfg = readConfig_();
  var s = { p: cfg.presenting || '', q: cfg.queue };
  cache.put('pres2', JSON.stringify(s), 21600);
  return s;
}
function setPresState_(s) { CacheService.getScriptCache().put('pres2', JSON.stringify(s), 21600); }

/** מה מוצג עכשיו (או null), ומי הבאה בתור. */
function getPresenting() {
  var me = requireUser_();
  var s = presState_();
  if (!s.p) return null;
  var projects = table_('Projects').rows;
  function find(id) { return projects.filter(function (r) { return String(r.id) === String(id); })[0]; }
  var p = find(s.p);
  if (!p || !visibleProject_(p, me)) return null;
  var idx = s.q.indexOf(s.p), next = null;
  if (idx >= 0 && s.q[idx + 1]) {
    var n = find(s.q[idx + 1]);
    if (n && visibleProject_(n, me)) next = { authorName: str_(n.authorName), title: str_(n.title) };
  }
  return { id: s.p, title: str_(p.title), authorName: str_(p.authorName), grade: str_(p.grade), next: next };
}

/** מורה: קובעת איזה תוצר מוצג עכשיו. id ריק = סיום ההצגה. */
function adminSetPresenting(id) {
  return withLock_(function () {
    requireTeacher_();
    id = String(id || '');
    if (id) findBy_(table_('Projects'), 'id', id, 'התוצר לא נמצא.');
    setConfigValue_('presentingProjectId', id);
    setPresState_({ p: id, q: presState_().q });
    return { presenting: id };
  });
}

/** מורה: סדר ההצגות (מערך מזהי תוצרים). מחזיר את הסקירה המעודכנת. */
function adminSetQueue(ids) {
  return withLock_(function () {
    requireTeacher_();
    var seen = {}, out = [];
    var projects = table_('Projects');
    (ids || []).forEach(function (id) {
      id = String(id);
      if (seen[id]) return;
      findBy_(projects, 'id', id, 'אחד התוצרים בתור לא נמצא.');
      seen[id] = true; out.push(id);
    });
    if (out.length > 40) throw new Error('התור ארוך מדי (עד 40).');
    setConfigValue_('presentQueue', out.join(','));
    setPresState_({ p: presState_().p, q: out });
    return adminOverview_();
  });
}
// ───────────── סקירת מורה: דשבורד והשוואה בין קבוצות ─────────────

function adminOverview() {
  requireTeacher_();
  return adminOverview_();
}

function adminOverview_() {
  var users = table_('Users').rows, orders = table_('Orders').rows, projects = table_('Projects').rows;
  var agg = aggregate_();
  var groups = {};
  ['ז', 'ח'].forEach(function (g) {
    var byStatus = {};
    STATUSES.forEach(function (s) { byStatus[s] = 0; });
    groups[g] = { students: 0, orders: 0, adopted: 0, byStatus: byStatus, projects: 0, hearts: 0, comments: 0 };
  });
  users.forEach(function (u) {
    var g = normGrade_(u.grade);
    if (g && String(u.role).trim().toLowerCase() !== 'teacher') groups[g].students++;
  });
  orders.forEach(function (o) {
    var g = str_(o.grade); if (!groups[g]) return;
    groups[g].orders++;
    if (splitList_(o.adoptedBy).length) groups[g].adopted++;
    if (groups[g].byStatus[str_(o.status)] !== undefined) groups[g].byStatus[str_(o.status)]++;
  });
  projects.forEach(function (p) {
    var g = str_(p.grade); if (!groups[g]) return;
    var id = String(p.id);
    groups[g].projects++;
    groups[g].hearts += agg.hearts[id] || 0;
    groups[g].comments += agg.comments[id] || 0;
  });
  var rows = users.filter(function (u) { return String(u.role).trim().toLowerCase() !== 'teacher'; }).map(function (u) {
    var em = String(u.email).toLowerCase();
    var brought = orders.filter(function (o) { return String(o.bringerEmail).toLowerCase() === em; }).map(function (o) { return str_(o.title); });
    var adopted = orders.filter(function (o) { return splitList_(o.adoptedEmails).map(function (e) { return e.toLowerCase(); }).indexOf(em) >= 0; })
      .map(function (o) { return str_(o.title) + ' (' + str_(o.status) + ')'; });
    var mine = projects.filter(function (p) { return String(p.authorEmail).toLowerCase() === em; }).map(function (p) { return str_(p.title); });
    return { email: em, name: String(u.firstName).trim(), grade: normGrade_(u.grade), brought: brought, adopted: adopted, projects: mine,
             heartsGiven: agg.givenBy[em] || 0, heartsReceived: agg.receivedBy[em] || 0, feedbackGiven: agg.commentsBy[em] || 0 };
  }).sort(function (a, b) { return (a.grade + a.name) < (b.grade + b.name) ? -1 : 1; });
  var plist = projects.map(function (p) { return { id: String(p.id), title: str_(p.title), authorName: str_(p.authorName), grade: str_(p.grade) }; });
  return { groups: groups, students: rows, projects: plist, statuses: STATUSES, presenting: readConfig_().presenting, queue: readConfig_().queue };
}

function setBio(text) {
  return withLock_(function () {
    var me = requireUser_();
    var t = table_('Users');
    var u = t.rows.filter(function (r) { return String(r.email).toLowerCase().trim() === me.email; })[0];
    if (!u) throw new Error('המשתמשת לא נמצאה.');
    u.bio = clean_(text, GALLERY_LIMITS.bio, 'קצת עליי', false);
    writeRow_(t, u);
    return { bio: u.bio };
  });
}
