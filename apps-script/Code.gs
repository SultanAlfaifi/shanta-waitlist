/**
 * شنطة — نموذج مخصص يحافظ على التصميم المعتمد ويكتب في Google Sheets.
 * شغّل setupShantaWaitlist_ من المحرر أولًا، ثم اختبر وانشر Web app يدويًا.
 * لا يرسل السكربت بريدًا، ولا يعرض التسجيلات للزوار، ولا ينشر نفسه.
 */
const WAITLIST = {
  contact: 'SultanAbdullahAlfaifi@gmail.com',
  limit: 100,
  sheetName: 'التسجيلات',
  noticeVersion: '2026-10-05-v1',
  stages: ['ابتدائي', 'متوسط', 'ثانوي', 'أخرى'],
  tasks: ['التحضير', 'الاختبارات', 'أوراق العمل', 'العروض', 'التصحيح', 'أخرى'],
  headers: ['وقت التسجيل بتوقيت الرياض', 'الاسم', 'البريد الإلكتروني', 'معلم سعودي يدرّس حاليًا', 'المراحل', 'المواد', 'المهام', 'الموافقة', 'نسخة إشعار الخصوصية', 'معرف التسجيل', 'مفتاح منع التكرار'],
};

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('قائمة انتظار شنطة 🎒')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Public Pages submissions omit Google account cookies and send data only in the body.
function doPost(event) {
  let result;
  try {
    const data = event && event.postData;
    if (!data || typeof data.contents !== 'string' || !data.contents.length || data.contents.length > 8192 ||
        !/^text\/plain(?:;|$)/i.test(String(data.type || '')) || Number(event.contentLength) > 8192) {
      throw new Error('invalid request');
    }
    result = submitWaitlist(JSON.parse(data.contents));
  } catch (error) {
    result = {ok: false, code: 'INVALID', message: 'تعذر قراءة طلب التسجيل. راجع الإجابات وأعد المحاولة.'};
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

// The trailing underscore makes this helper unavailable through google.script.run.
function setupShantaWaitlist_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const properties = PropertiesService.getScriptProperties();
    const existingId = properties.getProperty('SHANTA_CUSTOM_SHEET_ID');
    const spreadsheet = existingId ? SpreadsheetApp.openById(existingId) : SpreadsheetApp.create('شنطة — قائمة انتظار الدفعة الأولى');
    if (!existingId) properties.setProperty('SHANTA_CUSTOM_SHEET_ID', spreadsheet.getId());
    let sheet = spreadsheet.getSheetByName(WAITLIST.sheetName);
    if (!sheet) {
      sheet = !existingId ? spreadsheet.getSheets()[0].setName(WAITLIST.sheetName) : spreadsheet.insertSheet(WAITLIST.sheetName);
    }
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, WAITLIST.headers.length).setValues([WAITLIST.headers]);
      sheet.setFrozenRows(1);
      sheet.setRightToLeft(true);
      sheet.getRange(1, 1, 1, WAITLIST.headers.length).setBackground('#2456B5').setFontColor('#ffffff').setFontWeight('bold');
    }
    checkSheet_(sheet);
    SpreadsheetApp.flush();
    console.log('رابط جدول التسجيلات الخاص بك: ' + spreadsheet.getUrl());
    return spreadsheet.getUrl();
  } finally {
    lock.releaseLock();
  }
}

function getPublicConfig_() {
  return {
    contact: WAITLIST.contact,
    ready: !!PropertiesService.getScriptProperties().getProperty('SHANTA_CUSTOM_SHEET_ID'),
  };
}

function submitWaitlist(payload) {
  let clean;
  try {
    clean = validatePayload_(payload);
  } catch (error) {
    return {ok: false, code: 'INVALID', message: 'راجع الإجابات: اكتب اسمًا وبريدًا صحيحين، وحدد المرحلة والمادة ومهمة واحدة أو مهمتين، ووافق على استخدام البيانات.'};
  }
  if (clean.teacher !== 'نعم') {
    return {ok: false, code: 'INELIGIBLE', message: 'هذه الدفعة مخصصة للمعلمين والمعلمات السعوديين الذين يمارسون التدريس حاليًا.'};
  }
  const sheetId = PropertiesService.getScriptProperties().getProperty('SHANTA_CUSTOM_SHEET_ID');
  if (!sheetId) return {ok: false, code: 'UNAVAILABLE', message: 'التسجيل غير متاح حاليًا. حاول لاحقًا.'};
  const lock = LockService.getScriptLock();
  let locked = false;
  try {
    locked = lock.tryLock(10000);
    if (!locked) return {ok: false, code: 'BUSY', message: 'التسجيل مشغول لحظيًا. أعد المحاولة بعد قليل.'};
    const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(WAITLIST.sheetName);
    checkSheet_(sheet);
    const lastRow = sheet.getLastRow();
    // At capacity every address gets the same closure response; do not reveal membership.
    if (lastRow - 1 >= WAITLIST.limit) {
      return {ok: false, code: 'FULL', message: 'اكتملت تسجيلات الدفعة الأولى. إذا سبق أن سجلت، يبقى طلبك محفوظًا. تابع شنطة لمعرفة موعد الدفعة التالية.'};
    }
    const emailKey = emailKey_(clean.email);
    if (lastRow > 1) {
      const keys = sheet.getRange(2, WAITLIST.headers.length, lastRow - 1, 1).getValues();
      // Same acknowledgement for a new request and a retry; no membership disclosure.
      if (keys.some(function(row) { return row[0] === emailKey; })) return acknowledgement_();
    }
    const row = [
      Utilities.formatDate(new Date(), 'Asia/Riyadh', 'yyyy-MM-dd HH:mm:ss'),
      clean.name, clean.email, clean.teacher,
      clean.stages.join('، '), clean.subjects, clean.tasks.join('، '),
      'نعم', WAITLIST.noticeVersion, Utilities.getUuid(), emailKey,
    ].map(sheetText_);
    sheet.getRange(lastRow + 1, 1, 1, row.length).setNumberFormat('@').setValues([row]);
    SpreadsheetApp.flush();
    return acknowledgement_();
  } catch (error) {
    // Never return provider errors, spreadsheet identifiers, submitted data, or stack traces.
    return {ok: false, code: 'UNAVAILABLE', message: 'تعذر تأكيد حفظ التسجيل. أعد المحاولة بنفس البريد بعد قليل، أو تواصل معنا عبر البريد الموضح.'};
  } finally {
    if (locked) lock.releaseLock();
  }
}

function validatePayload_(payload) {
  if (!payload || typeof payload !== 'object' || Array.isArray(payload)) throw new Error('payload');
  const name = text_(payload.name, 120);
  const email = text_(payload.email, 254);
  const subjects = text_(payload.subjects, 250);
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('email');
  if (!['نعم', 'لا'].includes(payload.teacher)) throw new Error('teacher');
  if (payload.consent !== true) throw new Error('consent');
  const stages = selection_(payload.stages, WAITLIST.stages, WAITLIST.stages.length);
  const tasks = selection_(payload.tasks, WAITLIST.tasks, 2);
  return {name: name, email: email, subjects: subjects, teacher: payload.teacher, stages: stages, tasks: tasks};
}

function text_(value, maxLength) {
  if (typeof value !== 'string' || /[\u0000-\u001f\u007f]/.test(value)) throw new Error('text');
  const result = value.trim().normalize('NFC');
  if (!result || result.length > maxLength) throw new Error('length');
  return result;
}

function selection_(value, allowed, maxLength) {
  if (!Array.isArray(value) || value.length < 1 || value.length > maxLength) throw new Error('selection');
  if (new Set(value).size !== value.length || value.some(function(item) { return typeof item !== 'string' || !allowed.includes(item); })) throw new Error('option');
  return value.slice();
}

function emailKey_(email) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, email.toLowerCase(), Utilities.Charset.UTF_8)
    .map(function(byte) { return ('0' + ((byte + 256) % 256).toString(16)).slice(-2); }).join('');
}

function sheetText_(value) {
  const text = String(value);
  // Escape formula-leading text, even when the cell uses the plain-text number format.
  return /^[\s]*[=+\-@]/.test(text) ? "'" + text : text;
}

function checkSheet_(sheet) {
  if (!sheet) throw new Error('missing sheet');
  const headers = sheet.getRange(1, 1, 1, WAITLIST.headers.length).getValues()[0];
  if (headers.some(function(value, index) { return value !== WAITLIST.headers[index]; })) throw new Error('sheet schema');
}

function acknowledgement_() {
  return {ok: true, message: 'تم استلام طلب تسجيلك 🎒❤️\nسنتواصل معك عبر البريد الإلكتروني عند فتح التجربة، ونؤكد لك استحقاق عرض الدفعة الأولى برسالة مستقلة.'};
}
