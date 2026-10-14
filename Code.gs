/**
 * בנק הזמנות — שלב 1
 * Web App של Google Apps Script. הנתונים ב-Google Sheets המחוברת לפרויקט.
 *
 * קבצים: Code.gs (שרת), Index.html, Styles.html, Script.html
 * הרצה ראשונה: להריץ את setup() ידנית מהעורך (ולאשר הרשאות).
 */

// ───────────── קבועים ─────────────

var SHEET_HEADERS = {
  // adoptedEmails — עמודה נוספת (טכנית) שמקבילה ל-adoptedBy, כדי לזהות אימוץ לפי מייל ולא לפי שם
  Orders: ['id', 'createdAt', 'title', 'bringerEmail', 'bringerName', 'userDescription', 'need',
           'interviewDetail', 'size', 'status', 'adoptedBy', 'notes', 'productLink', 'grade', 'adoptedEmails'],
  Users: ['email', 'firstName', 'grade', 'role', 'bio'],
  Config: ['key', 'value'],
  // שלב 2 — גלריה
  Projects: ['id', 'createdAt', 'authorEmail', 'authorName', 'grade', 'title', 'description', 'link', 'thumbFileId', 'orderId', 'pitchPrep'],
  Hearts: ['projectId', 'fromEmail', 'createdAt'],
  // text = "מה אהבתי" (גלוי לכולן), suggestion = "הצעות לשיפור" (רק ליוצרת ולמורה)
  Comments: ['id', 'projectId', 'fromEmail', 'fromName', 'text', 'createdAt', 'suggestion'],
  Views: ['projectId', 'email', 'createdAt']
};

// חלון האימוץ נפתח בנפרד לכל קבוצה (שכבה ז / שכבה ח)
var CONFIG_DEFAULTS = { 'adoptionOpen_ז': 'FALSE', 'adoptionOpen_ח': 'FALSE', maxAdoptionsPerOrder: '2', allowAdoptOwn: 'TRUE', presentingProjectId: '', presentQueue: '' };

/**
 * הגיליון: אם הפרויקט מחובר לגיליון — הוא. אחרת (פרויקט עצמאי, כמו Web App) —
 * הגיליון ששמור ב-Script Properties תחת SS_ID. setup() יוצר אותו אם אין.
 */
function ss_(createIfMissing) {
  var ss = SpreadsheetApp.getActive();
  if (ss) return ss;
  if (SPREADSHEET_ID) return SpreadsheetApp.openById(SPREADSHEET_ID);
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SS_ID');
  if (id) return SpreadsheetApp.openById(id);
  if (!createIfMissing) throw new Error('עוד לא הוגדר גיליון נתונים. יש להריץ setup() מהעורך.');
  ss = SpreadsheetApp.create('בנק הזמנות - נתונים');
  props.setProperty('SS_ID', ss.getId());
  return ss;
}

var SIZES = ['קנבס', 'אפפסקריפט'];
// שלבי ההתקדמות של הזמנה, לפי הסדר
var STATUSES = ['פתוחה', 'נבחרה', 'בבנייה', 'תוצר ראשוני', 'משוב עם לקוח', 'סבב שיפור', 'מוכן להשקה', 'נמסרה'];
var USER_STATUSES = STATUSES.slice(1); // מה שמותר למשתמשת לקבוע בעצמה (הכול חוץ מ"פתוחה")
var GRADES = ['ז', 'ח'];

// מגבלות אורך (תווים)
var LIMITS = { title: 60, userDescription: 120, need: 200, interviewDetail: 300, notes: 300, productLink: 500 };

// ───────────── כניסה לאפליקציה ─────────────

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('בנק הזמנות')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** טוען קובץ HTML (Styles / Script) להטמעה בתוך Index */
function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

// ───────────── הגדרה ראשונית ─────────────

/** יוצר גיליונות וכותרות אם חסרים. בטוח להריץ שוב ושוב. */
function setup() {
  var ss = ss_(true);
  Logger.log('גיליון הנתונים: ' + ss.getUrl());
  Object.keys(SHEET_HEADERS).forEach(function (name) {
    var headers = SHEET_HEADERS[name];
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      // טקסט רגיל לכל העמודות — מונע המרה אוטומטית של תאריכים/TRUE ושל תווי "=" לנוסחאות
      sh.getRange(1, 1, sh.getMaxRows(), headers.length).setNumberFormat('@');
      sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
      sh.setFrozenRows(1);
      sh.setRightToLeft(true);
    } else {
      // גיליון קיים: מוסיף עמודות חדשות שחסרות (למשל bio ב-Users)
      var cur = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
      headers.forEach(function (hd) {
        if (cur.indexOf(hd) >= 0) return;
        var col = sh.getLastColumn() + 1;
        sh.getRange(1, col, sh.getMaxRows(), 1).setNumberFormat('@');
        sh.getRange(1, col).setValue(hd).setFontWeight('bold');
        cur.push(hd);
      });
    }
  });
  // מיגרציה חד-פעמית: אימוץ הזמנה של עצמי מותר כברירת מחדל (המורה יכולה לכבות בלשונית ניהול)
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('MIGRATED_ADOPT_OWN')) {
    setConfigValue_('allowAdoptOwn', 'TRUE');
    props.setProperty('MIGRATED_ADOPT_OWN', '1');
  }
  // בדיקת גישה לתיקיית הצילומים (גם מבקשת את הרשאת Drive בהרצה הראשונה)
  Logger.log('תיקיית צילומים: ' + DriveApp.getFolderById(DRIVE_FOLDER_ID).getName());
  // ערכי ברירת מחדל ב-Config (רק מה שחסר)
  var cfg = table_('Config');
  var existing = {};
  cfg.rows.forEach(function (r) { existing[String(r.key)] = true; });
  Object.keys(CONFIG_DEFAULTS).forEach(function (k) {
    if (!existing[k]) appendRow_(cfg, { key: k, value: CONFIG_DEFAULTS[k] });
  });
  // אם אין משתמשות — מוסיף את בעל הפרויקט כמורה כדי שאפשר יהיה להיכנס
  var users = table_('Users');
  if (users.rows.length === 0) {
    var owner = (Session.getEffectiveUser().getEmail() || '').toLowerCase();
    if (owner) appendRow_(users, { email: owner, firstName: 'מורה', grade: '', role: 'teacher' });
  }
}

/** 6 הזמנות לדוגמה לבדיקה. בטוח להרצה חוזרת (לא מכפיל). */
function seedDemoData() {
  setup();
  var t = table_('Orders');
  var have = {};
  t.rows.forEach(function (r) { have[String(r.id)] = true; });
  var now = new Date().toISOString();
  var demo = [
    { id: 'demo-1', title: 'מחשבון טיפים למסעדה', bringerName: 'נועה', bringerEmail: 'demo.noa@example.com', userDescription: 'דוד שלי, מלצר',
      need: 'לחשב מהר כמה טיפ להשאיר ולחלק בין כמה אנשים', interviewDetail: 'הוא אמר שהוא תמיד מתבלבל כשמחלקים חשבון בין חמישה',
      size: 'קנבס', status: 'פתוחה', adoptedBy: '', adoptedEmails: '', notes: '', productLink: '', grade: 'ז' },
    { id: 'demo-2', title: 'לוח תורנויות למשפחה', bringerName: 'מיכל', bringerEmail: 'demo.michal@example.com', userDescription: 'שכנה שלי, אמא לארבעה',
      need: 'לדעת מי תורן היום בבית ולסמן שהוא סיים', interviewDetail: 'הילדים תמיד טוענים שהם כבר עשו',
      size: 'אפפסקריפט', status: 'פתוחה', adoptedBy: '', adoptedEmails: '', notes: 'כדאי שכל ילד יראה רק את התורנות שלו', productLink: '', grade: 'ח' },
    { id: 'demo-3', title: 'ממיר מתכונים לכמות אחרת', bringerName: 'שירה', bringerEmail: 'demo.shira@example.com', userDescription: 'סבתא שלי, אופה לשבת',
      need: 'להכפיל או לחצות מתכון בלי לטעות בחישובים', interviewDetail: 'סבתא כותבת הכול בעיפרון על פתקים ומוחקת',
      size: 'קנבס', status: 'נבחרה', adoptedBy: 'תהילה', adoptedEmails: 'demo.tehila@example.com', notes: '', productLink: '', grade: 'ז' },
    { id: 'demo-4', title: 'מעקב אימונים לקבוצת כדורסל', bringerName: 'אביגיל', bringerEmail: 'demo.avigail@example.com', userDescription: 'המאמן של אחי',
      need: 'לרשום מי הגיע לאימון ומי לא', interviewDetail: 'כל שבוע הוא מאבד את הדף עם הרשימה',
      size: 'אפפסקריפט', status: 'בבנייה', adoptedBy: 'רבקה, יעל', adoptedEmails: 'demo.rivka@example.com,demo.yael@example.com', notes: '', productLink: '', grade: 'ח' },
    { id: 'demo-5', title: 'מחולל שמות לחתלתולים', bringerName: 'הדס', bringerEmail: 'demo.hadas@example.com', userDescription: 'דודה שלי, מתנדבת בעמותת חתולים',
      need: 'לתת שמות מקוריים לחתלתולים שנולדים בעמותה', interviewDetail: 'נגמרו להם הרעיונות אחרי 40 חתולים',
      size: 'קנבס', status: 'נמסרה', adoptedBy: 'טל', adoptedEmails: 'demo.tal@example.com', notes: '', productLink: 'https://example.com', grade: 'ז' },
    { id: 'demo-6', title: 'מד זמן לשיעורי נהיגה', bringerName: 'רות', bringerEmail: 'demo.ruth@example.com', userDescription: 'אחות גדולה שלומדת נהיגה',
      need: 'לראות כמה שעות נהיגה כבר עשתה וכמה נשאר', interviewDetail: 'המורה לנהיגה לא תמיד זוכר כמה שיעורים היו',
      size: 'אפפסקריפט', status: 'פתוחה', adoptedBy: '', adoptedEmails: '', notes: '', productLink: '', grade: 'ח' }
  ];
  demo.forEach(function (d) {
    if (have[d.id]) return;
    d.createdAt = now;
    appendRow_(t, d);
  });
}

// ───────────── עזרים: גיליונות ─────────────

function sheet_(name) {
  var sh = ss_(false).getSheetByName(name);
  if (!sh) throw new Error('הגיליון ' + name + ' לא קיים. יש להריץ setup().');
  return sh;
}

/** קורא גיליון לאובייקטים לפי שורת הכותרות. לכל שורה _row = מספר השורה בגיליון. */
function table_(name) {
  var sh = sheet_(name);
  var vals = sh.getDataRange().getValues();
  var headers = vals[0].map(String);
  var rows = [];
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === '') continue; // שורה ריקה
    var o = { _row: i + 1 };
    for (var c = 0; c < headers.length; c++) o[headers[c]] = vals[i][c];
    rows.push(o);
  }
  return { sh: sh, headers: headers, rows: rows };
}

function rowArray_(t, obj) {
  return t.headers.map(function (h) {
    var v = obj[h];
    if (v === undefined || v === null) return '';
    if (v instanceof Date) return v.toISOString();
    return String(v);
  });
}

function writeRow_(t, obj) {
  var r = t.sh.getRange(obj._row, 1, 1, t.headers.length);
  r.setNumberFormat('@');
  r.setValues([rowArray_(t, obj)]);
}

function appendRow_(t, obj) {
  var row = t.sh.getLastRow() + 1;
  var r = t.sh.getRange(row, 1, 1, t.headers.length);
  r.setNumberFormat('@');
  r.setValues([rowArray_(t, obj)]);
}

function str_(v) { return v instanceof Date ? v.toISOString() : String(v == null ? '' : v); }

function splitList_(s) {
  return String(s || '').split(',').map(function (x) { return x.trim(); }).filter(Boolean);
}

// ───────────── הגדרות (Config) ─────────────

function readConfig_() {
  var cfg = {};
  Object.keys(CONFIG_DEFAULTS).forEach(function (k) { cfg[k] = CONFIG_DEFAULTS[k]; });
  table_('Config').rows.forEach(function (r) { cfg[String(r.key)] = r.value; });
  var max = parseInt(cfg.maxAdoptionsPerOrder, 10);
  return {
    open: { 'ז': isTrue_(cfg['adoptionOpen_ז']), 'ח': isTrue_(cfg['adoptionOpen_ח']) },
    maxAdoptionsPerOrder: (max >= 1 && max <= 10) ? max : 2,
    allowAdoptOwn: isTrue_(cfg.allowAdoptOwn),
    presenting: String(cfg.presentingProjectId || ''),
    queue: splitList_(cfg.presentQueue) // סדר ההצגות (מזהי תוצרים)
  };
}

function isTrue_(v) { return v === true || String(v).trim().toUpperCase() === 'TRUE'; }

function setConfigValue_(key, value) {
  var t = table_('Config');
  var row = t.rows.filter(function (r) { return String(r.key) === key; })[0];
  if (row) { row.value = String(value); writeRow_(t, row); }
  else appendRow_(t, { key: key, value: String(value) });
}

// ───────────── משתמשות והרשאות ─────────────

/** מחזיר את המשתמשת הנוכחית. registered=false אם לא ברשימת Users. null אם אין מייל. */
function me_() {
  var email = (Session.getActiveUser().getEmail() || '').toLowerCase().trim();
  if (!email) return null;
  var u = table_('Users').rows.filter(function (r) { return String(r.email).toLowerCase().trim() === email; })[0];
  if (!u) return { email: email, registered: false };
  return {
    email: email,
    registered: true,
    firstName: String(u.firstName).trim() || 'חברה',
    grade: normGrade_(u.grade),
    role: String(u.role).trim().toLowerCase() === 'teacher' ? 'teacher' : 'student'
  };
}

function normGrade_(g) {
  var s = String(g || '').trim().charAt(0);
  return GRADES.indexOf(s) >= 0 ? s : '';
}

function requireUser_() {
  var me = me_();
  if (!me || !me.registered) throw new Error('אין לך הרשאה. בקשי מהמורה להוסיף אותך.');
  return me;
}

function requireTeacher_() {
  var me = requireUser_();
  if (me.role !== 'teacher') throw new Error('הפעולה הזו שמורה למורה.');
  return me;
}

/** כל כתיבה עוברת כאן: נעילה אחת לכל הסקריפט, כדי שאימוצים מקבילים לא ידרסו זה את זה */
function withLock_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error('המערכת עסוקה כרגע, נסי שוב בעוד רגע.');
  try { return fn(); } finally { lock.releaseLock(); }
}

// ───────────── ולידציה ─────────────

/** מנקה טקסט חופשי, בודק אורך, וחוסם מספרי טלפון ומיילים (פרטיות) */
function clean_(v, max, label, required) {
  var s = String(v == null ? '' : v).replace(/[\u0000-\u001f\u007f]+/g, ' ').replace(/\s{2,}/g, ' ').trim();
  if (required && !s) throw new Error('חסר: ' + label);
  if (s.length > max) throw new Error(label + ' ארוך מדי (עד ' + max + ' תווים).');
  if (/(\d[\s\-.]*){7,}/.test(s) || s.indexOf('@') >= 0) {
    throw new Error('ב"' + label + '" נראה שיש מספר טלפון, כתובת או מייל. בלי פרטים מזהים בבקשה.');
  }
  return s;
}

function cleanLink_(v) {
  var s = String(v == null ? '' : v).trim();
  if (!s) return '';
  if (s.length > LIMITS.productLink) throw new Error('הקישור ארוך מדי.');
  if (!/^https?:\/\/[^\s]+$/i.test(s)) throw new Error('הקישור חייב להתחיל ב-https://');
  return s;
}

/** המלצת גודל: כן לכל ארבע השאלות → קנבס, אחרת אפפסקריפט */
function recommendSize_(answers) {
  if (!answers || answers.length !== 4) throw new Error('יש לענות על כל ארבע שאלות הגודל.');
  var allYes = answers.every(function (a) { return a === true; });
  answers.forEach(function (a) { if (a !== true && a !== false) throw new Error('יש לענות על כל ארבע שאלות הגודל.'); });
  return allYes ? 'קנבס' : 'אפפסקריפט';
}

// ───────────── מצב שנשלח ללקוח ─────────────

function toPublic_(r, me) {
  var emails = splitList_(r.adoptedEmails).map(function (e) { return e.toLowerCase(); });
  var o = {
    id: String(r.id),
    createdAt: str_(r.createdAt),
    title: str_(r.title),
    bringerName: str_(r.bringerName),
    userDescription: str_(r.userDescription),
    need: str_(r.need),
    interviewDetail: str_(r.interviewDetail),
    size: str_(r.size),
    status: str_(r.status),
    adoptedBy: splitList_(r.adoptedBy),
    notes: str_(r.notes),
    productLink: str_(r.productLink),
    grade: str_(r.grade),
    mine: str_(r.bringerEmail).toLowerCase() === me.email,
    adoptedByMe: emails.indexOf(me.email) >= 0
  };
  if (me.role === 'teacher') o.bringerEmail = str_(r.bringerEmail); // מיילים נחשפים רק למורה
  return o;
}

function state_(me) {
  // כל תלמידה רואה רק את הקבוצה שלה (ז או ח). מורה רואה הכול.
  var orders = table_('Orders').rows
    .filter(function (r) { return me.role === 'teacher' || !me.grade || str_(r.grade) === me.grade; })
    .map(function (r) { return toPublic_(r, me); })
    .sort(function (a, b) { return a.createdAt < b.createdAt ? 1 : -1; });
  return {
    status: 'ok',
    me: { firstName: me.firstName, grade: me.grade, role: me.role, uid: uid_(me.email) },
    config: readConfig_(),
    orders: orders,
    sheetUrl: me.role === 'teacher' ? ss_(false).getUrl() : ''
  };
}

/** טעינה ראשונית */
function getBootstrap() {
  var me = me_();
  if (!me) return { status: 'noemail' };
  if (!me.registered) return { status: 'unregistered', email: me.email };
  return state_(me);
}

/**
 * הרשמה עצמית: מי שלא ברשימה בוחרת שם פרטי ושכבה ונוספת כתלמידה.
 * תפקיד מורה אפשר לתת רק ידנית בגיליון Users (אין דרך לבקש אותו מהלקוח).
 */
function registerMe(data) {
  return withLock_(function () {
    var me = me_();
    if (!me) throw new Error('לא הצלחנו לזהות אותך. היכנסי עם המייל הארגוני.');
    if (me.registered) return state_(me);
    data = data || {};
    var name = String(data.firstName || '').replace(/\s+/g, ' ').trim();
    if (!/^[֐-׿a-zA-Z'\- ]{2,20}$/.test(name)) throw new Error('שם פרטי: 2 עד 20 אותיות, בלי מספרים.');
    if (GRADES.indexOf(data.grade) < 0) throw new Error('יש לבחור שכבה: ז או ח.');
    appendRow_(table_('Users'), { email: me.email, firstName: name, grade: data.grade, role: 'student' });
    return state_(me_());
  });
}

/** רענון (קריאה בלבד) */
function getState() {
  return state_(requireUser_());
}

/** עוטף כל פעולת כתיבה: נעילה → בדיקת משתמשת → הפעולה → מצב מעודכן */
function mutate_(fn) {
  return withLock_(function () {
    var me = requireUser_();
    fn(me, readConfig_());
    return state_(me);
  });
}

function findOrder_(t, id) {
  var o = t.rows.filter(function (r) { return String(r.id) === String(id); })[0];
  if (!o) throw new Error('ההזמנה לא נמצאה (אולי נמחקה). רענני את הדף.');
  return o;
}

// ───────────── פעולות של תלמידות ─────────────

/** data: {title, userDescription, need, interviewDetail, notes, answers:[bool x4]} */
function createOrder(data) {
  return mutate_(function (me) {
    data = data || {};
    var order = {
      id: 'o' + new Date().getTime().toString(36) + Math.random().toString(36).slice(2, 6),
      createdAt: new Date().toISOString(),
      title: clean_(data.title, LIMITS.title, 'כותרת', true),
      bringerEmail: me.email,
      bringerName: me.firstName,
      userDescription: clean_(data.userDescription, LIMITS.userDescription, 'מי המשתמש', true),
      need: clean_(data.need, LIMITS.need, 'הצורך', true),
      interviewDetail: clean_(data.interviewDetail, LIMITS.interviewDetail, 'פרט מהראיון', true),
      size: recommendSize_(data.answers),
      status: 'פתוחה',
      adoptedBy: '',
      notes: clean_(data.notes, LIMITS.notes, 'הערות', false),
      productLink: '',
      grade: me.grade,
      adoptedEmails: ''
    };
    appendRow_(table_('Orders'), order);
  });
}

/** אימוץ הזמנה — כל החוקים נאכפים כאן */
function adoptOrder(id) {
  return mutate_(function (me, cfg) {
    var t = table_('Orders');
    var o = findOrder_(t, id);
    var g = str_(o.grade);
    if (me.role !== 'teacher' && me.grade && g !== me.grade) throw new Error('ההזמנה הזו שייכת לקבוצה אחרת.');
    if (!cfg.open[g]) throw new Error('חלון האימוץ של קבוצה ' + (g || '?') + "' סגור כרגע.");
    var emails = splitList_(o.adoptedEmails).map(function (e) { return e.toLowerCase(); });
    var names = splitList_(o.adoptedBy);
    if (String(o.bringerEmail).toLowerCase() === me.email && !cfg.allowAdoptOwn) throw new Error('אי אפשר לאמץ הזמנה שהבאת בעצמך.');
    if (o.status !== 'פתוחה' && o.status !== 'נבחרה') throw new Error('ההזמנה הזו כבר לא זמינה לאימוץ.');
    if (emails.indexOf(me.email) >= 0) throw new Error('כבר אימצת את ההזמנה הזו.');
    if (emails.length >= cfg.maxAdoptionsPerOrder) throw new Error('ההזמנה כבר מלאה (' + cfg.maxAdoptionsPerOrder + ' אימוצים).');
    emails.push(me.email);
    names.push(me.firstName);
    o.adoptedEmails = emails.join(',');
    o.adoptedBy = names.join(', ');
    if (o.status === 'פתוחה') o.status = 'נבחרה'; // האימוץ הראשון
    writeRow_(t, o);
  });
}

/** ביטול אימוץ של עצמי, כל עוד ההזמנה לא בבנייה */
function cancelAdoption(id) {
  return mutate_(function (me) {
    var t = table_('Orders');
    var o = findOrder_(t, id);
    var emails = splitList_(o.adoptedEmails).map(function (e) { return e.toLowerCase(); });
    var names = splitList_(o.adoptedBy);
    var idx = emails.indexOf(me.email);
    if (idx < 0) throw new Error('לא אימצת את ההזמנה הזו.');
    if (o.status !== 'נבחרה') throw new Error('אי אפשר לבטל אימוץ אחרי שהבנייה התחילה. פני למורה.');
    emails.splice(idx, 1);
    names.splice(idx, 1);
    o.adoptedEmails = emails.join(',');
    o.adoptedBy = names.join(', ');
    if (emails.length === 0) o.status = 'פתוחה';
    writeRow_(t, o);
  });
}

/** עדכון סטטוס וקישור לתוצר ע"י מי שהביאה או מי שאימצה. data: {status?, productLink?} */
function updateMyOrder(id, data) {
  return mutate_(function (me) {
    data = data || {};
    var t = table_('Orders');
    var o = findOrder_(t, id);
    var emails = splitList_(o.adoptedEmails).map(function (e) { return e.toLowerCase(); });
    var involved = String(o.bringerEmail).toLowerCase() === me.email || emails.indexOf(me.email) >= 0;
    if (!involved && me.role !== 'teacher') throw new Error('אפשר לעדכן רק הזמנות שהבאת או אימצת.');
    if (data.status !== undefined && data.status !== o.status) {
      if (USER_STATUSES.indexOf(data.status) < 0) throw new Error('סטטוס לא חוקי.');
      if (emails.length === 0) throw new Error('אפשר לשנות סטטוס רק אחרי שמישהי אימצה את ההזמנה.');
      o.status = data.status;
    }
    if (data.productLink !== undefined) o.productLink = cleanLink_(data.productLink);
    writeRow_(t, o);
  });
}

/** יוצרת עורכת את ההזמנה שהביאה (כל עוד היא קיימת). f: {title, userDescription, need, interviewDetail, notes} */
function editMyOrder(id, f) {
  return mutate_(function (me) {
    f = f || {};
    var t = table_('Orders');
    var o = findOrder_(t, id);
    if (String(o.bringerEmail).toLowerCase() !== me.email && me.role !== 'teacher') throw new Error('אפשר לערוך רק הזמנה שהבאת.');
    if (f.title !== undefined) o.title = clean_(f.title, LIMITS.title, 'כותרת', true);
    if (f.userDescription !== undefined) o.userDescription = clean_(f.userDescription, LIMITS.userDescription, 'מי המשתמש', true);
    if (f.need !== undefined) o.need = clean_(f.need, LIMITS.need, 'הצורך', true);
    if (f.interviewDetail !== undefined) o.interviewDetail = clean_(f.interviewDetail, LIMITS.interviewDetail, 'פרט מהראיון', true);
    if (f.notes !== undefined) o.notes = clean_(f.notes, LIMITS.notes, 'הערות', false);
    writeRow_(t, o);
  });
}

function validName_(v) {
  var name = String(v || '').replace(/\s+/g, ' ').trim();
  if (!/^[֐-׿a-zA-Z'\- ]{2,20}$/.test(name)) throw new Error('שם פרטי: 2 עד 20 אותיות, בלי מספרים.');
  return name;
}

/** משנה שם משתמשת בכל מקום שבו הוא נשמר (הזמנות, אימוצים, תוצרים, תגובות) */
function renameUser_(email, name) {
  email = String(email).toLowerCase().trim();
  var ut = table_('Users');
  var u = ut.rows.filter(function (r) { return String(r.email).toLowerCase().trim() === email; })[0];
  if (!u) throw new Error('המשתמשת לא נמצאה.');
  u.firstName = name; writeRow_(ut, u);

  var ot = table_('Orders');
  ot.rows.forEach(function (o) {
    var changed = false;
    if (String(o.bringerEmail).toLowerCase() === email) { o.bringerName = name; changed = true; }
    var emails = splitList_(o.adoptedEmails).map(function (e) { return e.toLowerCase(); });
    var i = emails.indexOf(email), names = splitList_(o.adoptedBy);
    if (i >= 0 && names.length === emails.length) { names[i] = name; o.adoptedBy = names.join(', '); changed = true; }
    if (changed) writeRow_(ot, o);
  });
  var pt = table_('Projects');
  pt.rows.forEach(function (p) { if (String(p.authorEmail).toLowerCase() === email) { p.authorName = name; writeRow_(pt, p); } });
  var ct = table_('Comments');
  ct.rows.forEach(function (c) { if (String(c.fromEmail).toLowerCase() === email) { c.fromName = name; writeRow_(ct, c); } });
}

/** כל אחת יכולה לתקן את השם של עצמה */
function setMyName(name) {
  return mutate_(function (me) {
    var n = validName_(name);
    renameUser_(me.email, n);
    me.firstName = n; // כדי שהמצב שחוזר ללקוח יכלול את השם החדש
  });
}

/** מורה מתקנת שם של תלמידה. מחזיר את הסקירה המעודכנת. */
function adminSetName(email, name) {
  return withLock_(function () {
    requireTeacher_();
    renameUser_(email, validName_(name));
    return adminOverview_();
  });
}

// ───────────── פעולות מורה ─────────────

/** obj: {openGrade:'ז'|'ח', openValue:bool, maxAdoptionsPerOrder?, allowAdoptOwn?} */
function adminSetConfig(obj) {
  return withLock_(function () {
    var me = requireTeacher_();
    obj = obj || {};
    if (obj.openGrade !== undefined) {
      if (GRADES.indexOf(obj.openGrade) < 0) throw new Error('קבוצה לא חוקית.');
      setConfigValue_('adoptionOpen_' + obj.openGrade, obj.openValue ? 'TRUE' : 'FALSE');
    }
    if (obj.allowAdoptOwn !== undefined) setConfigValue_('allowAdoptOwn', obj.allowAdoptOwn ? 'TRUE' : 'FALSE');
    if (obj.maxAdoptionsPerOrder !== undefined) {
      var n = parseInt(obj.maxAdoptionsPerOrder, 10);
      if (!(n >= 1 && n <= 10)) throw new Error('מספר אימוצים מקסימלי חייב להיות בין 1 ל-10.');
      setConfigValue_('maxAdoptionsPerOrder', n);
    }
    return state_(me);
  });
}

/** עריכת הזמנה ע"י מורה. מעבר ל"פתוחה" מנקה את האימוצים. */
function adminUpdateOrder(id, f) {
  return withLock_(function () {
    var me = requireTeacher_();
    f = f || {};
    var t = table_('Orders');
    var o = findOrder_(t, id);
    if (f.title !== undefined) o.title = clean_(f.title, LIMITS.title, 'כותרת', true);
    if (f.userDescription !== undefined) o.userDescription = clean_(f.userDescription, LIMITS.userDescription, 'מי המשתמש', true);
    if (f.need !== undefined) o.need = clean_(f.need, LIMITS.need, 'הצורך', true);
    if (f.interviewDetail !== undefined) o.interviewDetail = clean_(f.interviewDetail, LIMITS.interviewDetail, 'פרט מהראיון', true);
    if (f.notes !== undefined) o.notes = clean_(f.notes, LIMITS.notes, 'הערות', false);
    if (f.productLink !== undefined) o.productLink = cleanLink_(f.productLink);
    if (f.size !== undefined) {
      if (SIZES.indexOf(f.size) < 0) throw new Error('גודל לא חוקי.');
      o.size = f.size;
    }
    if (f.grade !== undefined) {
      if (f.grade !== '' && GRADES.indexOf(f.grade) < 0) throw new Error('שכבה לא חוקית.');
      o.grade = f.grade;
    }
    if (f.status !== undefined) {
      if (STATUSES.indexOf(f.status) < 0) throw new Error('סטטוס לא חוקי.');
      o.status = f.status;
      if (f.status === 'פתוחה') { o.adoptedBy = ''; o.adoptedEmails = ''; }
    }
    writeRow_(t, o);
    return state_(me);
  });
}

function adminDeleteOrder(id) {
  return withLock_(function () {
    var me = requireTeacher_();
    var t = table_('Orders');
    var o = findOrder_(t, id);
    t.sh.deleteRow(o._row);
    return state_(me);
  });
}

/** ייצוא כל ההזמנות ל-CSV (הלקוח מוריד את הקובץ) */
function adminExport() {
  requireTeacher_();
  var t = table_('Orders');
  function cell(v) {
    var s = str_(v);
    if (/^[=+\-@\t\r]/.test(s)) s = "'" + s; // מניעת הזרקת נוסחאות כשפותחים באקסל
    return '"' + s.replace(/"/g, '""') + '"';
  }
  var lines = [t.headers.map(cell).join(',')];
  t.rows.forEach(function (r) {
    lines.push(t.headers.map(function (h) { return cell(r[h]); }).join(','));
  });
  return { filename: 'orders-' + new Date().toISOString().slice(0, 10) + '.csv', csv: lines.join('\r\n') };
}
