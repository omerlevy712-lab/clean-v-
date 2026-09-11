/**
 * איפוס אוטומטי של עמודה B בגיליון "פורמט"
 * -------------------------------------------------
 * כל מוצ"ש ב-23:59 (שעון ישראל) כל תא בעמודה B שערכו "v"
 * מוחלף ל-"---" (ברירת המחדל השחורה של הרשימה הנפתחת).
 *
 * התקנה: מריצים פעם אחת את הפונקציה setup()
 */

/** ===== הגדרות – אפשר לשנות כאן ולא בקוד עצמו ===== */
var SHEET_NAME   = 'פורמט';  // שם הלשונית
var COLUMN       = 2;         // עמודה B
var FIRST_ROW    = 2;         // שורת הנתונים הראשונה (מתחת לכותרת)
var VALUE_FROM   = 'v';       // הערך שמתאפס
var VALUE_TO     = '---';     // הערך שאליו מאפסים
var RUN_WEEKDAY  = 6;         // 0=ראשון ... 6=שבת
var RUN_HOUR     = 23;
var RUN_MINUTE   = 59;
var HANDLER      = 'weeklyReset';

/**
 * הפונקציה שמבצעת את האיפוס בפועל.
 * אפשר גם להריץ אותה ידנית כדי לבדוק שהכל עובד.
 */
function resetGreenV() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error('לא נמצאה לשונית בשם "' + SHEET_NAME + '"');
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < FIRST_ROW) return 0;

  var range  = sheet.getRange(FIRST_ROW, COLUMN, lastRow - FIRST_ROW + 1, 1);
  var values = range.getValues();
  var changed = 0;

  for (var i = 0; i < values.length; i++) {
    var cell = values[i][0];
    if (typeof cell === 'string' && cell.trim() === VALUE_FROM) {
      values[i][0] = VALUE_TO;
      changed++;
    }
  }

  if (changed > 0) {
    range.setValues(values);
    SpreadsheetApp.flush();
  }
  return changed;
}

/**
 * מה שהטריגר מפעיל: מאפס, ואז קובע מיד את ההרצה של השבוע הבא.
 * הקביעה מחדש נמצאת ב-finally כדי שהשרשרת לא תישבר גם אם היתה שגיאה.
 */
function weeklyReset() {
  try {
    resetGreenV();
  } finally {
    scheduleNextRun();
  }
}

/**
 * התקנה חד-פעמית: מוחק טריגרים ישנים וקובע את ההרצה הקרובה.
 */
function setup() {
  var next = scheduleNextRun();
  Logger.log('הותקן. ההרצה הבאה: ' + next);
  return next;
}

/**
 * קובע טריגר חד-פעמי למוצ"ש הקרוב ב-23:59 (לפי אזור הזמן של הפרויקט).
 * טריגר לשעה מדויקת מדויק בהרבה מטריגר שבועי לפי טווח שעות.
 */
function scheduleNextRun() {
  removeTriggers();

  var now  = new Date();
  var next = new Date(now.getTime());
  next.setHours(RUN_HOUR, RUN_MINUTE, 0, 0);

  var daysAhead = (RUN_WEEKDAY - next.getDay() + 7) % 7;
  // אם זו כבר שבת והשעה עברה (או ממש עומדת לעבור) – מדלגים לשבוע הבא
  if (daysAhead === 0 && next.getTime() <= now.getTime() + 60 * 1000) {
    daysAhead = 7;
  }
  next.setDate(next.getDate() + daysAhead);

  ScriptApp.newTrigger(HANDLER).timeBased().at(next).create();
  return next;
}

/**
 * מוחק את כל הטריגרים של הסקריפט הזה (כדי שלא יצטברו כפילויות).
 */
function removeTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === HANDLER) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/**
 * תפריט נוח בתוך הגיליון עצמו.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('איפוס משמרות')
    .addItem('אפס עכשיו את ה-v בעמודה B', 'menuResetNow')
    .addItem('הפעל/רענן תזמון שבועי', 'menuSetup')
    .addItem('בטל תזמון שבועי', 'menuRemove')
    .addToUi();
}

function menuResetNow() {
  var n = resetGreenV();
  SpreadsheetApp.getUi().alert('אופסו ' + n + ' תאים.');
}

function menuSetup() {
  var next = setup();
  SpreadsheetApp.getUi().alert('התזמון הופעל. ההרצה הבאה: ' + next);
}

function menuRemove() {
  removeTriggers();
  SpreadsheetApp.getUi().alert('התזמון בוטל.');
}
