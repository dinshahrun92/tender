// ================= TENDERFLOW PRO - Code.gs =================

// ================= CONFIGURATION =================
const SPREADSHEET_ID    = '13o93cZansFR0_7_FxxPu0yw-RYKPdEHwBgCyF9LVCPY';
const UPLOAD_FOLDER_ID  = '1bRgxe3TgZdGoEB7RsEoAFEOw5lG7VtSn';
const LIBRARY_FOLDER_ID = '1jMOEeeS2MV3QUm9Cn0Ujosv-wMSfcz40';
const ADMIN_EMAIL       = 'ceoo@axicom.info';

// ===================== TELEGRAM CONFIG =====================
const TELEGRAM_BOT_TOKEN    = '8205662791:AAEjPImQifUFvkTRNYX6Gk3Xj9EhWmFX4PQ';
const TELEGRAM_BOT_USERNAME = 'AxiCom_bot';
// ===========================================================

// ── SHEET NAMES ──────────────────────────────────────────────
const SHEET_TENDERS       = 'Tenders';
const SHEET_COMMENTS      = 'Comments';
const SHEET_TASKS         = 'Tasks';
const SHEET_USERS         = 'Users';
const SHEET_LOGS          = 'Logs';
const SHEET_RESET         = 'PasswordResets';
const SHEET_NOTIFICATIONS = 'Notifications';
const SHEET_ACTIVITY      = 'Activity';
const SHEET_MENTIONS      = 'Mentions';
const SHEET_DM            = 'DirectMessages';

const SUBFOLDERS = ['1. Tender Docs', '2. Technical', '3. Financial', '4. Submission Proof'];

// ── HEADERS ──
const HEADERS = {
  [SHEET_TENDERS]:       ['ID','Title','Agency','DueDate','Price','Briefing','Indicative','Via','FinURL','TechURL','BondURL','AddendumURL','SlipURL','Status','FolderURL','Owner','Team','IsDeleted','CreatedAt','StartDate'],
  [SHEET_COMMENTS]:      ['TenderID','UserEmail','UserName','Timestamp','Message'],
  [SHEET_TASKS]:         ['TenderID','Description','IsDone','CreatedBy','AssignedTo','FileURL','FileName','CreatedAt','TaskID','IsDeleted'],
  [SHEET_USERS]:         ['Email','FullName','Role','PasswordHash','SessionToken','TokenExpiry','CreatedAt','LastLogin','TelegramChatID'],
  [SHEET_LOGS]:          ['Timestamp','User','TenderID','Action','OldValue','NewValue'],
  [SHEET_RESET]:         ['Email','ResetCode','Expiry','Used'],
  [SHEET_NOTIFICATIONS]: ['UserEmail','Type','TenderID','Message','IsRead','Timestamp','ActionURL'],
  [SHEET_ACTIVITY]:      ['Timestamp','UserEmail','UserName','TenderID','TenderTitle','Action','Details'],
  [SHEET_MENTIONS]:      ['UserEmail','TenderID','CommentID','MentionedBy','Message','IsRead','Timestamp'],
  [SHEET_DM]:            ['Timestamp','FromEmail','FromName','ToEmail','Message']
};

const STATUS_VALUES             = ['Calling For Tender','Active','Review','Submitted','Successful','Closed','Unsuccessful','Cancelled'];
const TOKEN_EXPIRY_HOURS        = 24;
const RESET_CODE_EXPIRY_MINUTES = 30;
const CACHE_DURATION            = 600;
const APP_SECRET                = 'TF_2025_SECURE';

// ================= ENTRY POINT =================
function doGet(e) {
  try {
    setupDatabase();
    return HtmlService.createTemplateFromFile('index').evaluate()
      .setTitle('TenderFlow Pro')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (error) {
    Logger.log('doGet error: ' + error.toString());
    return HtmlService.createHtmlOutput('<h1>Error loading application</h1><p>' + error.toString() + '</p>');
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) return ContentService.createTextOutput('No data');
    const update = JSON.parse(e.postData.contents);
    handleTelegramUpdate(update);
    return ContentService.createTextOutput('OK');
  } catch (error) {
    Logger.log('doPost error: ' + error.toString());
    return ContentService.createTextOutput('Error: ' + error.toString());
  }
}

// ================= SETUP & UTILS =================

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function setupDatabase() {
  try {
    const ss = getSpreadsheet();
    Object.keys(HEADERS).forEach(sheetName => {
      let sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        sheet.appendRow(HEADERS[sheetName]);
        if (sheetName === SHEET_TENDERS) applyTenderValidation(sheet);
        sheet.setFrozenRows(1);
        Logger.log('Created sheet: ' + sheetName);
      } else {
        const currentCols  = sheet.getLastColumn();
        const requiredCols = HEADERS[sheetName].length;
        if (currentCols < requiredCols) {
          const newHeaders = HEADERS[sheetName].slice(currentCols);
          sheet.getRange(1, currentCols + 1, 1, newHeaders.length).setValues([newHeaders]);
          Logger.log('Added columns to ' + sheetName + ': ' + newHeaders.join(', '));
        }
      }
    });
  } catch (error) {
    Logger.log('setupDatabase error: ' + error.toString());
    throw new Error('Failed to setup database: ' + error.message);
  }
}

function applyTenderValidation(sheet) {
  try {
    const statusRule = SpreadsheetApp.newDataValidation().requireValueInList(STATUS_VALUES).setAllowInvalid(false).build();
    sheet.getRange(2, 14, 1000).setDataValidation(statusRule);
    const priceRule = SpreadsheetApp.newDataValidation().requireNumberGreaterThanOrEqualTo(0).setAllowInvalid(false).build();
    sheet.getRange(2, 5, 1000).setDataValidation(priceRule);
  } catch (e) { Logger.log('applyTenderValidation error: ' + e.toString()); }
}

function getSheet(name) {
  const ss    = getSpreadsheet();
  let   sheet = ss.getSheetByName(name);
  if (!sheet) {
    Logger.log('Auto-creating missing sheet: ' + name);
    sheet = ss.insertSheet(name);
    if (HEADERS[name]) {
      sheet.appendRow(HEADERS[name]);
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

function hashPassword(password, email) {
  if (!password || !email) return '';
  const salted = password + email.toLowerCase() + APP_SECRET;
  const raw    = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salted);
  return raw.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function generateResetCode() { return Math.floor(100000 + Math.random() * 900000).toString(); }

function logAction(userEmail, tenderId, action, oldVal, newVal) {
  try { getSheet(SHEET_LOGS).appendRow([new Date(), userEmail, tenderId, action, oldVal || '', newVal || '']); }
  catch (e) { Logger.log('logAction failed: ' + e.toString()); }
}

function createCalendarEvent(title, dateObj, description) {
  try { CalendarApp.getDefaultCalendar().createAllDayEvent('[TENDER DUE] ' + title, dateObj, { description }); }
  catch (e) { Logger.log('createCalendarEvent failed: ' + e.toString()); }
}

function invalidateCache() {
  try { const c = CacheService.getScriptCache(); c.remove('tenders_all'); c.remove('stats_all'); }
  catch (e) { Logger.log('invalidateCache failed: ' + e.toString()); }
}

let _ss = null;
function getSpreadsheet() {
  if (!_ss) _ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return _ss;
}

// ================= TELEGRAM FUNCTIONS =================

function sendTelegramMessage(chatId, message) {
  if (!chatId || !TELEGRAM_BOT_TOKEN) return false;
  try {
    const res = UrlFetchApp.fetch('https://api.telegram.org/bot' + TELEGRAM_BOT_TOKEN + '/sendMessage', {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ chat_id: chatId.toString(), text: message, parse_mode: 'Markdown', disable_web_page_preview: true }),
      muteHttpExceptions: true
    });
    const result = JSON.parse(res.getContentText());
    if (!result.ok) Logger.log('Telegram failed: ' + result.description);
    return result.ok;
  } catch (e) { Logger.log('sendTelegramMessage error: ' + e.toString()); return false; }
}

function sendTelegramByEmail(userEmail, message) {
  try {
    const chatId = getTelegramChatId(userEmail);
    if (!chatId) return false;
    return sendTelegramMessage(chatId, message);
  } catch (e) { Logger.log('sendTelegramByEmail error: ' + e.toString()); return false; }
}

function getTelegramChatId(userEmail) {
  try {
    const data = getSheet(SHEET_USERS).getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === userEmail) return data[i][8] ? data[i][8].toString() : null;
    }
    return null;
  } catch (e) { Logger.log('getTelegramChatId error: ' + e.toString()); return null; }
}

function saveTelegramChatId(token, chatId) {
  try {
    const user  = authenticate(token);
    const sheet = getSheet(SHEET_USERS);
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === user.email) {
        sheet.getRange(i + 1, 9).setValue(chatId.toString());
        SpreadsheetApp.flush();
        sendTelegramMessage(chatId, '✅ *TenderFlow Connected!*\n\nHi ' + user.name + '! You will now receive TenderFlow notifications here.\n\n🚀 Bot: @' + TELEGRAM_BOT_USERNAME);
        return { success: true, msg: 'Telegram linked successfully!' };
      }
    }
    return { success: false, msg: 'User not found.' };
  } catch (e) { Logger.log('saveTelegramChatId error: ' + e.toString()); return { success: false, msg: e.message }; }
}

function removeTelegramChatId(token) {
  try {
    const user  = authenticate(token);
    const sheet = getSheet(SHEET_USERS);
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === user.email) { sheet.getRange(i + 1, 9).setValue(''); SpreadsheetApp.flush(); return { success: true }; }
    }
    return { success: false };
  } catch (e) { return { success: false, msg: e.message }; }
}

function getTelegramStatus(token) {
  try {
    const user   = authenticate(token);
    const chatId = getTelegramChatId(user.email);
    return { linked: !!chatId, chatId: chatId };
  } catch (e) { return { linked: false }; }
}

function handleTelegramUpdate(update) {
  try {
    if (!update.message) return;
    const msg       = update.message;
    const chatId    = msg.chat.id.toString();
    const text      = msg.text || '';
    const firstName = msg.from.first_name || 'there';
    if (text.startsWith('/start')) {
      const parts = text.split(' ');
      if (parts.length > 1) {
        linkTelegramByStartToken(parts[1], chatId, firstName);
      } else {
        sendTelegramMessage(chatId,
          '👋 Hi *' + firstName + '*! Welcome to *TenderFlow* bot.\n\nYour Chat ID is: `' + chatId + '`\n\n_Copy this ID and paste it in TenderFlow → Telegram settings._'
        );
      }
    } else if (text === '/status') {
      sendTelegramMessage(chatId, '✅ Bot active! Chat ID: `' + chatId + '`');
    } else if (text === '/chatid') {
      sendTelegramMessage(chatId, '🆔 Your Chat ID: `' + chatId + '`\n\nPaste this in TenderFlow Settings → Telegram.');
    } else if (text === '/help') {
      sendTelegramMessage(chatId, '*TenderFlow Commands*\n\n/start — Link account\n/status — Check connection\n/chatid — Get Chat ID\n/help — This menu');
    }
  } catch (e) { Logger.log('handleTelegramUpdate error: ' + e.toString()); }
}

function linkTelegramByStartToken(token, chatId, firstName) {
  try {
    const email = Utilities.newBlob(Utilities.base64Decode(token)).getDataAsString();
    const sheet = getSheet(SHEET_USERS);
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === email) {
        sheet.getRange(i + 1, 9).setValue(chatId);
        SpreadsheetApp.flush();
        sendTelegramMessage(chatId,
          '🎉 *Account Linked!*\n\nHi *' + data[i][1] + '*! Your TenderFlow account is now connected.\n\n✅ Task assignments\n💬 Mentions & comments\n🔄 Status changes\n⏰ Deadline reminders\n💬 Direct messages'
        );
        return;
      }
    }
    sendTelegramMessage(chatId, '❌ Account not found. Try again from TenderFlow settings.');
  } catch (e) { Logger.log('linkTelegramByStartToken error: ' + e.toString()); }
}

function getTelegramLinkUrl(token) {
  try {
    const user    = authenticate(token);
    const encoded = Utilities.base64Encode(user.email);
    return { success: true, url: 'https://t.me/' + TELEGRAM_BOT_USERNAME + '?start=' + encoded, chatId: getTelegramChatId(user.email) };
  } catch (e) { return { success: false, url: '' }; }
}

function sendTelegramTest(token) {
  try {
    const user   = authenticate(token);
    const chatId = getTelegramChatId(user.email);
    if (!chatId) return { success: false, msg: 'No Telegram account linked.' };
    const sent = sendTelegramMessage(chatId, '🔔 *Test Notification*\n\nHi *' + user.name + '*! TenderFlow notifications are working! ✅');
    return sent ? { success: true, msg: 'Test message sent!' } : { success: false, msg: 'Failed. Check bot token.' };
  } catch (e) { return { success: false, msg: e.message }; }
}

function manualSetWebhook() {
  const PRODUCTION_URL = 'https://script.google.com/macros/s/AKfycbz81UJU8znVLZA9yziLH8HPFAuCBllBMTF2EXWNEoLyCkLvN9AU8y_djcx9tSsx3Ss3tA/exec';
  try {
    const res    = UrlFetchApp.fetch('https://api.telegram.org/bot' + TELEGRAM_BOT_TOKEN + '/setWebhook', {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ url: PRODUCTION_URL, allowed_updates: ['message'] }),
      muteHttpExceptions: true
    });
    const result = JSON.parse(res.getContentText());
    return result.ok ? '✅ Webhook set!' : '❌ Failed: ' + result.description;
  } catch (e) { return 'Error: ' + e.toString(); }
}

function checkTelegramWebhook() {
  try {
    const res    = UrlFetchApp.fetch('https://api.telegram.org/bot' + TELEGRAM_BOT_TOKEN + '/getWebhookInfo', { muteHttpExceptions: true });
    const result = JSON.parse(res.getContentText());
    if (result.ok) {
      const i = result.result;
      return '📊 Webhook:\nURL: ' + (i.url || 'Not set') + '\nPending: ' + (i.pending_update_count || 0) + '\nLast error: ' + (i.last_error_message || 'None');
    }
    return '❌ ' + result.description;
  } catch (e) { return '❌ ' + e.toString(); }
}

function testDoPost() {
  const e = { postData: { contents: JSON.stringify({ update_id: 1, message: { message_id: 1, from: { id: 123, first_name: 'Test' }, chat: { id: 123, type: 'private' }, date: 0, text: '/start' } }) } };
  return doPost(e).getContent();
}

// ================= AUTH FUNCTIONS =================

function handleRegister(email, name, role, password) {
  try {
    if (!email || !email.match(/^[^\s@]+@[^\s@]+\.[^\s@]+$/)) return { success: false, msg: 'Invalid email address' };
    if (!name || name.trim().length < 2)  return { success: false, msg: 'Name must be at least 2 characters' };
    if (!password || password.length < 6) return { success: false, msg: 'Password must be at least 6 characters' };
    const sheet = getSheet(SHEET_USERS);
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) { if (data[i][0] === email) return { success: false, msg: 'Email already registered' }; }
    const token  = Utilities.getUuid();
    const expiry = new Date(Date.now() + TOKEN_EXPIRY_HOURS * 3600000);
    sheet.appendRow([email, name, role || 'User', hashPassword(password, email), token, expiry, new Date(), new Date(), '']);
    SpreadsheetApp.flush();
    const session = { email, name, role: role || 'User', expiry };
    CacheService.getScriptCache().put('session_' + token, JSON.stringify(session), SESSION_CACHE_TTL);
    logAction(email, '', 'USER_REGISTER', '', name);
    return { success: true, token, user: { email, name, role: role || 'User' } };
  } catch (error) { return { success: false, msg: 'Registration failed: ' + error.message }; }
}

function handleLogin(email, password) {
  try {
    if (!email || !password) return { success: false, msg: 'Email and password required' };
    const sheet = getSheet(SHEET_USERS);
    const data  = sheet.getDataRange().getValues();
    const hash  = hashPassword(password, email);
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === email && data[i][3] === hash) {
        const newToken = Utilities.getUuid();
        const expiry   = new Date(Date.now() + TOKEN_EXPIRY_HOURS * 3600000);
        sheet.getRange(i + 1, 5).setValue(newToken);
        sheet.getRange(i + 1, 6).setValue(expiry);
        sheet.getRange(i + 1, 8).setValue(new Date());
        SpreadsheetApp.flush();
        const session = { email: data[i][0], name: data[i][1], role: data[i][2], expiry: expiry };
        CacheService.getScriptCache().put('session_' + newToken, JSON.stringify(session), SESSION_CACHE_TTL);
        logAction(email, '', 'USER_LOGIN', '', '');
        return { success: true, token: newToken, user: { email: data[i][0], name: data[i][1], role: data[i][2] } };
      }
    }
    return { success: false, msg: 'Invalid email or password' };
  } catch (error) { return { success: false, msg: 'Login failed: ' + error.message }; }
}

function requestPasswordReset(email) {
  try {
    if (!email || !email.match(/^[^\s@]+@[^\s@]+\.[^\s@]+$/)) return { success: false, msg: 'Invalid email address' };
    const userSheet = getSheet(SHEET_USERS);
    const userData  = userSheet.getDataRange().getValues();
    let userExists = false, userName = '';
    for (let i = 1; i < userData.length; i++) { if (userData[i][0] === email) { userExists = true; userName = userData[i][1]; break; } }
    if (!userExists) return { success: false, msg: 'Email not found in system' };
    const resetCode = generateResetCode();
    const expiry    = new Date(Date.now() + RESET_CODE_EXPIRY_MINUTES * 60000);
    getSheet(SHEET_RESET).appendRow([email, resetCode, expiry, false]);
    SpreadsheetApp.flush();
    let emailSent = false;
    try {
      if (MailApp.getRemainingDailyQuota() > 0) {
        MailApp.sendEmail({ to: email, subject: '[TenderFlow] Password Reset Code',
          htmlBody: '<h2>Password Reset</h2><p>Hi ' + userName + ',</p><p>Your reset code is: <b style="font-size:24px;letter-spacing:4px">' + resetCode + '</b></p><p>Expires in ' + RESET_CODE_EXPIRY_MINUTES + ' minutes.</p>' });
        emailSent = true;
      }
    } catch (e) { Logger.log('Email failed: ' + e.toString()); }
    sendTelegramByEmail(email, '🔐 *TenderFlow Password Reset*\n\nHi *' + userName + '*!\n\nYour reset code:\n`' + resetCode + '`\n\nExpires in ' + RESET_CODE_EXPIRY_MINUTES + ' minutes.');
    logAction(email, '', 'PASSWORD_RESET_REQUEST', '', emailSent ? 'EMAIL_SENT' : 'EMAIL_FAILED');
    return emailSent
      ? { success: true, msg: 'Reset code sent to ' + email + '. Also check Telegram if linked.' }
      : { success: true, msg: 'Reset code: ' + resetCode + ' (Email failed — check Telegram or copy this code)' };
  } catch (error) { return { success: false, msg: 'Reset request failed: ' + error.message }; }
}

function verifyResetCode(email, code) {
  try {
    const data = getSheet(SHEET_RESET).getDataRange().getValues();
    const now  = new Date();
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === email && data[i][1] === code && !data[i][3]) {
        return new Date(data[i][2]) > now ? { success: true, valid: true } : { success: false, msg: 'Code expired' };
      }
    }
    return { success: false, msg: 'Invalid reset code' };
  } catch (error) { return { success: false, msg: error.message }; }
}

function resetPassword(email, code, newPassword) {
  try {
    if (!newPassword || newPassword.length < 6) return { success: false, msg: 'Password must be at least 6 characters' };
    const resetSheet = getSheet(SHEET_RESET);
    const resetData  = resetSheet.getDataRange().getValues();
    const now        = new Date();
    let rowIndex     = -1;
    for (let i = resetData.length - 1; i >= 1; i--) {
      if (resetData[i][0] === email && resetData[i][1] === code && !resetData[i][3]) {
        if (new Date(resetData[i][2]) < now) return { success: false, msg: 'Code expired' };
        rowIndex = i + 1; break;
      }
    }
    if (rowIndex === -1) return { success: false, msg: 'Invalid or used code' };
    const userSheet = getSheet(SHEET_USERS);
    const userData  = userSheet.getDataRange().getValues();
    for (let i = 1; i < userData.length; i++) {
      if (userData[i][0] === email) {
        userSheet.getRange(i + 1, 4).setValue(hashPassword(newPassword, email));
        userSheet.getRange(i + 1, 5).setValue('');
        userSheet.getRange(i + 1, 6).setValue('');
        break;
      }
    }
    resetSheet.getRange(rowIndex, 4).setValue(true);
    SpreadsheetApp.flush();
    logAction(email, '', 'PASSWORD_RESET_SUCCESS', '', '');
    sendTelegramByEmail(email, '✅ *Password Reset Successful*\n\nYour TenderFlow password has been reset. Please log in.');
    return { success: true, msg: 'Password reset successful. Please login.' };
  } catch (error) { return { success: false, msg: error.message }; }
}

const SESSION_CACHE_TTL = 300;
function authenticate(token) {
  if (!token) throw new Error('Unauthorized: No token provided');
  const cache    = CacheService.getScriptCache();
  const cacheKey = 'session_' + token;
  const cached   = cache.get(cacheKey);
  if (cached) {
    const session = JSON.parse(cached);
    if (new Date(session.expiry) < new Date()) {
      cache.remove(cacheKey);
      throw new Error('Session expired. Please login again');
    }
    return { email: session.email, name: session.name, role: session.role };
  }
  const data = getSheet(SHEET_USERS).getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][4] === token) {
      if (new Date(data[i][5]) < new Date()) throw new Error('Session expired. Please login again');
      const session = { email: data[i][0], name: data[i][1], role: data[i][2], expiry: data[i][5] };
      cache.put(cacheKey, JSON.stringify(session), SESSION_CACHE_TTL);
      return { email: session.email, name: session.name, role: session.role };
    }
  }
  throw new Error('Invalid session. Please login again');
}

function checkSession(token) {
  try { return { valid: true, user: authenticate(token) }; }
  catch (e) { return { valid: false, error: e.message }; }
}

function userHeartbeat(token) {
  try {
    const user = authenticate(token);
    CacheService.getScriptCache().put('online_' + user.email, 'true', 120);
  } catch (e) { Logger.log('userHeartbeat error: ' + e.toString()); }
}

function getAllUsers(token) {
  try {
    authenticate(token);
    const sheet = getSheet(SHEET_USERS);
    if (sheet.getLastRow() <= 1) return [];
    const cache = CacheService.getScriptCache();
    return sheet.getDataRange().getValues().slice(1).map(r => ({
      email: r[0], name: r[1], role: r[2],
      isOnline: cache.get('online_' + r[0]) === 'true',
      hasTelegram: !!r[8]
    }));
  } catch (error) { return []; }
}

function sendNotification(recipient, subject, body) {
  if (!recipient || recipient.indexOf('@') === -1) return;
  try { MailApp.sendEmail({ to: recipient, subject, htmlBody: body }); }
  catch (e) { Logger.log('sendNotification failed: ' + e.toString()); }
}

// ================= TENDERS FUNCTIONS =================

function getTenders(token) {
  try {
    authenticate(token);
    const cache  = CacheService.getScriptCache();
    const cached = cache.get('tenders_all');
    if (cached) return JSON.parse(cached);
    const sheet   = getSheet(SHEET_TENDERS);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];
    const data   = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    const now    = new Date();
    const doneStatuses = ['Submitted', 'Closed', 'Unsuccessful', 'Cancelled', 'Successful'];
    const result = data
      .filter(r => !r[17])
      .map(r => {
        let dueStr = '-', dueObj = null;
        if (r[3]) {
          dueObj = (r[3] instanceof Date) ? r[3] : new Date(r[3]);
          if (!isNaN(dueObj.getTime())) {
            dueStr = Utilities.formatDate(dueObj, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
          }
        }
        let briefingStr = '-';
        if (r[5]) {
          const bd = (r[5] instanceof Date) ? r[5] : new Date(r[5]);
          if (!isNaN(bd.getTime())) {
            briefingStr = Utilities.formatDate(bd, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
          }
        }
        let startStr = '-';
        if (r[19]) {
          const sd = (r[19] instanceof Date) ? r[19] : new Date(r[19]);
          if (!isNaN(sd.getTime())) {
            startStr = Utilities.formatDate(sd, Session.getScriptTimeZone(), 'yyyy-MM-dd');
          }
        }
        const status = r[13] || 'Calling For Tender';
        return {
          id: r[0], title: r[1], agency: r[2],
          dueDate: dueStr,
          briefingDate: briefingStr,
          startDate: startStr,
          submissionPrice: parseFloat(r[4]) || 0,
          status, folderUrl: r[14] || '', owner: r[15] || 'Unassigned', team: r[16] || '',
          daysLeft: doneStatuses.includes(status) ? null : (dueObj && !isNaN(dueObj.getTime())) ? Math.ceil((dueObj - now) / 86400000) : null
        };
      })
      .sort((a, b) => {
        const aSub = ['Submitted','Closed','Unsuccessful','Cancelled','Successful'].includes(a.status);
        const bSub = ['Submitted','Closed','Unsuccessful','Cancelled','Successful'].includes(b.status);
        if (aSub && !bSub) return 1; if (!aSub && bSub) return -1;
        return new Date(a.dueDate) - new Date(b.dueDate);
      });
    cache.put('tenders_all', JSON.stringify(result), CACHE_DURATION);
    return result;
  } catch (error) { Logger.log('getTenders error: ' + error.toString()); throw new Error('Failed to load tenders: ' + error.message); }
}

function generateNextTenderId(sheet) {
  let maxNum = 0;
  if (sheet.getLastRow() >= 2) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().forEach(row => {
      if (typeof row[0] === 'string' && row[0].match(/^T\d+$/)) { const n = parseInt(row[0].substring(1), 10); if (n > maxNum) maxNum = n; }
    });
  }
  return 'T' + String(maxNum + 1).padStart(5, '0');
}

function validateTenderInput(title, agency, dueDate) {
  const errors = [];
  if (!title || title.trim().length < 3) errors.push('Title must be at least 3 characters');
  if (!agency || agency.trim().length < 2) errors.push('Agency name is required');
  if (!dueDate) errors.push('Due date is required');
  if (errors.length > 0) throw new Error(errors.join('; '));
}

// Handles both 'yyyy-MM-dd' and 'yyyy-MM-ddTHH:mm' (datetime-local format)
function parseDateTime_(d) {
  if (!d) return '';
  if (typeof d === 'string' && d.includes('T')) {
    const [datePart, timePart] = d.split('T');
    const [y, mo, day] = datePart.split('-').map(Number);
    const [hr, min]    = timePart.split(':').map(Number);
    return new Date(y, mo - 1, day, hr, min);
  }
  if (typeof d === 'string') {
    const [y, mo, day] = d.split('-').map(Number);
    return new Date(y, mo - 1, day);
  }
  return d;
}

// ── UPDATED: addTender now accepts an optional startDate parameter ──
function addTender(token, title, agency, dueDate, submissionPrice, siteVisitDate, docSaleDate, teamMembers, startDate) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    validateTenderInput(title, agency, dueDate);
    const user  = authenticate(token);
    const sheet = getSheet(SHEET_TENDERS);
    const newId = generateNextTenderId(sheet);
    let folderUrl = '';
    if (UPLOAD_FOLDER_ID && UPLOAD_FOLDER_ID.length > 20) {
      try {
        const parent    = DriveApp.getFolderById(UPLOAD_FOLDER_ID);
        const newFolder = parent.createFolder(newId + ' - ' + title);
        SUBFOLDERS.forEach(sub => newFolder.createFolder(sub));
        if (teamMembers) { teamMembers.split(',').forEach(e => { try { newFolder.addEditor(e.trim()); } catch (err) { Logger.log('Failed to add editor: ' + e.trim()); } }); }
        folderUrl = newFolder.getUrl();
      } catch (e) { Logger.log('Drive folder creation failed: ' + e.toString()); }
    }

    const dueDateObj      = parseDateTime_(dueDate);
    const siteVisitDateObj = parseDateTime_(siteVisitDate);
    const docSaleDateObj  = parseDateTime_(docSaleDate);

    // Use provided startDate if given, otherwise default to today
    const startDateObj = startDate ? parseDateTime_(startDate) : new Date();

    sheet.appendRow([
      newId, title, agency,
      dueDateObj,
      parseFloat(submissionPrice) || 0,
      siteVisitDateObj,
      docSaleDateObj,
      '', '', '', '', '', '',
      'Calling For Tender', folderUrl, user.name, teamMembers || '', false, new Date(),
      startDateObj  // column 20 — StartDate
    ]);
    SpreadsheetApp.flush();
    logAction(user.email, newId, 'CREATE', '', title);
    logActivity(user.email, user.name, newId, title, 'CREATE', 'Project created');
    if (dueDateObj) createCalendarEvent(title, dueDateObj, 'Agency: ' + agency + '\nDrive: ' + folderUrl);
    addComment(token, newId, 'Project created by ' + user.name + (teamMembers ? '<br>Team: ' + teamMembers : ''));
    if (teamMembers) {
      teamMembers.split(',').forEach(email => {
        const e = email.trim();
        createNotification(e, 'TASK_ASSIGNED', newId, 'You were added to project: ' + title, '#/tender/' + newId);
        sendNotification(e, '[TenderFlow] Added to Project: ' + title, 'Added by ' + user.name);
        sendTelegramByEmail(e, '👥 *Added to Project*\n\n*' + title + '*\nAdded by: ' + user.name + '\nID: `' + newId + '`');
      });
    }
    invalidateCache();
    return getTenders(token);
  } catch (error) { Logger.log('addTender error: ' + error.toString()); throw new Error('Failed to create tender: ' + error.message); }
  finally { lock.releaseLock(); }
}

function updateTenderDetails(token, id, title, agency, dueDate, price, status) {
  try {
    const user  = authenticate(token);
    const sheet = getSheet(SHEET_TENDERS);
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] !== id) continue;
      const row = i + 1, oldStatus = data[i][13], oldPrice = parseFloat(data[i][4]) || 0;
      const newPrice = parseFloat(price) || 0, tenderTitle = data[i][1], teamMembers = data[i][16];
      if (status && oldStatus !== status) {
        sheet.getRange(row, 14).setValue(status);
        addComment(token, id, 'Status changed to <b>' + status + '</b> by ' + user.name);
        logAction(user.email, id, 'STATUS_CHANGE', oldStatus, status);
        logActivity(user.email, user.name, id, tenderTitle, 'STATUS_CHANGE', oldStatus + ' → ' + status);
        notifyTeamTelegram(teamMembers, user.name, tenderTitle, id, '🔄 *Status Update*\n\n*' + tenderTitle + '*\n' + oldStatus + ' → *' + status + '*\nBy: ' + user.name);
      }
      if (price !== undefined && Math.abs(oldPrice - newPrice) > 0.01) {
        sheet.getRange(row, 5).setValue(newPrice);
        logAction(user.email, id, 'PRICE_CHANGE', oldPrice.toString(), newPrice.toString());
        addComment(token, id, 'Price updated to <b>RM ' + newPrice.toLocaleString() + '</b>');
      }
      if (title)   sheet.getRange(row, 2).setValue(title);
      if (agency)  sheet.getRange(row, 3).setValue(agency);
      if (dueDate) sheet.getRange(row, 4).setValue(parseDateTime_(dueDate));
      SpreadsheetApp.flush(); invalidateCache(); break;
    }
    return getTenders(token);
  } catch (error) { Logger.log('updateTenderDetails error: ' + error.toString()); throw new Error('Failed to update tender: ' + error.message); }
}

// ── NEW: Update only the Start Date (column 20) ──
function updateTenderStartDate(token, id, newStartDate) {
  try {
    const user  = authenticate(token);
    const sheet = getSheet(SHEET_TENDERS);
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] !== id) continue;
      const oldVal = data[i][19] ? data[i][19].toString() : '';
      sheet.getRange(i + 1, 20).setValue(parseDateTime_(newStartDate));
      logAction(user.email, id, 'START_DATE_CHANGE', oldVal, newStartDate);
      logActivity(user.email, user.name, id, data[i][1], 'START_DATE_CHANGE', 'Start date → ' + newStartDate);
      SpreadsheetApp.flush();
      invalidateCache();
      break;
    }
    return getTenders(token);
  } catch (error) { Logger.log('updateTenderStartDate error: ' + error.toString()); throw new Error('Failed to update start date: ' + error.message); }
}

// ── NEW: Update only the Due Date (column 4) ──
function updateTenderDueDate(token, id, newDueDate) {
  try {
    const user  = authenticate(token);
    const sheet = getSheet(SHEET_TENDERS);
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] !== id) continue;
      const oldVal = data[i][3] ? data[i][3].toString() : '';
      sheet.getRange(i + 1, 4).setValue(parseDateTime_(newDueDate));
      logAction(user.email, id, 'DUE_DATE_CHANGE', oldVal, newDueDate);
      logActivity(user.email, user.name, id, data[i][1], 'DUE_DATE_CHANGE', 'Due date → ' + newDueDate);
      SpreadsheetApp.flush();
      invalidateCache();
      break;
    }
    return getTenders(token);
  } catch (error) { Logger.log('updateTenderDueDate error: ' + error.toString()); throw new Error('Failed to update due date: ' + error.message); }
}

function updateStatusOnly(token, id, newStatus) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const user  = authenticate(token);
    const sheet = getSheet(SHEET_TENDERS);
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] !== id) continue;
      const oldStatus = data[i][13], tenderTitle = data[i][1], teamMembers = data[i][16];
      if (oldStatus !== newStatus) {
        sheet.getRange(i + 1, 14).setValue(newStatus);
        addComment(token, id, 'Status changed to <b>' + newStatus + '</b>');
        logAction(user.email, id, 'STATUS_CHANGE', oldStatus, newStatus);
        logActivity(user.email, user.name, id, tenderTitle, 'STATUS_CHANGE', oldStatus + ' → ' + newStatus);
        SpreadsheetApp.flush(); invalidateCache();
        notifyTeamTelegram(teamMembers, user.name, tenderTitle, id, '🔄 *Status Update*\n\n*' + tenderTitle + '*\n' + oldStatus + ' → *' + newStatus + '*\nBy: ' + user.name);
      }
      break;
    }
    return getTenders(token);
  } catch (error) { Logger.log('updateStatusOnly error: ' + error.toString()); throw new Error('Failed to update status: ' + error.message); }
  finally { lock.releaseLock(); }
}

function deleteTender(token, id) {
  try {
    const user  = authenticate(token);
    const sheet = getSheet(SHEET_TENDERS);
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) { sheet.getRange(i + 1, 18).setValue(true); logAction(user.email, id, 'DELETE', '', ''); SpreadsheetApp.flush(); invalidateCache(); break; }
    }
    return getTenders(token);
  } catch (error) { throw new Error('Failed to delete tender: ' + error.message); }
}

function shareProjectFolder(token, folderUrl) {
  try {
    const user    = authenticate(token);
    const idMatch = folderUrl.match(/[-\w]{25,}/);
    if (!idMatch) throw new Error('Invalid folder URL');
    DriveApp.getFolderById(idMatch[0]).addEditor(user.email);
    return { success: true, msg: 'Folder shared with ' + user.email };
  } catch (e) { return { success: false, msg: 'Failed to share: ' + e.message }; }
}

function getDashboardStats(token) {
  try {
    authenticate(token);
    const cache  = CacheService.getScriptCache();
    const cached = cache.get('stats_all');
    if (cached) return JSON.parse(cached);
    const tenders = getTenders(token);
    const active  = tenders.filter(t => ['Calling For Tender','Active','Review','Submitted'].includes(t.status));
    const stats   = { total: tenders.length, active: active.length, value: active.reduce((s,t)=>s+(t.submissionPrice||0),0), success: tenders.filter(t=>t.status==='Successful').length, recent: tenders.slice(0,5) };
    cache.put('stats_all', JSON.stringify(stats), CACHE_DURATION);
    return stats;
  } catch (error) { throw new Error('Failed to load stats: ' + error.message); }
}

// ================= TASKS FUNCTIONS =================

function findTaskRow_(sheet, taskId) {
  if (!taskId || taskId.toString().trim() === '') return -1;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][8] === taskId && data[i][9] !== true) return i + 1;
  }
  return -1;
}

function getTasks(token, tenderId) {
  try {
    authenticate(token);
    const sheet = getSheet(SHEET_TASKS);
    if (sheet.getLastRow() <= 1) return [];
    return sheet.getDataRange().getValues().slice(1)
      .filter(r => r[0] === tenderId && r[8] && r[9] !== true)
      .map(r => ({
        taskId:   r[8] || '',
        tid:      r[0], desc: r[1], done: r[2],
        assignee: r[4] ? r[4].toString() : '',
        fileUrl:  r[5] || '', fileName: r[6] || ''
      }));
  } catch (error) { throw new Error('Failed to load tasks: ' + error.message); }
}

function getMyTasks(token) {
  try {
    const user       = authenticate(token);
    const sheetTasks = getSheet(SHEET_TASKS);
    if (sheetTasks.getLastRow() <= 1) return [];
    const tasks     = sheetTasks.getDataRange().getValues().slice(1).filter(r => r[8] && r[9] !== true);
    const tenderMap = {};
    getSheet(SHEET_TENDERS).getDataRange().getValues().slice(1).forEach(r => {
      if (r[17]) return;
      let d = r[3];
      if (d instanceof Date) d = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
      tenderMap[r[0]] = { title: r[1], status: r[13]||'Calling For Tender', dueDate: d||'9999-12-31' };
    });
    return tasks.map(r => {
      const info = tenderMap[r[0]] || { title:'Unknown', status:'Unknown', dueDate:'9999-12-31' };
      return { taskId: r[8]||'', tid: r[0], desc: r[1], done: r[2], assignee: r[4]?r[4].toString():'', tenderTitle: info.title, tenderStatus: info.status, tenderDueDate: info.dueDate };
    }).filter(t => {
      const assignees = t.assignee.split(',').map(e=>e.trim());
      return assignees.includes(user.email) && !['Submitted','Closed','Unsuccessful','Cancelled','Successful'].includes(t.tenderStatus);
    });
  } catch (error) { throw new Error('Failed to load tasks: ' + error.message); }
}

function addTask(token, tenderId, desc, assigneeEmails) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const user = authenticate(token);
    if (!desc || !desc.trim()) throw new Error('Task description is required');
    const taskId = 'TK-' + Utilities.getUuid().split('-')[0].toUpperCase();
    getSheet(SHEET_TASKS).appendRow([tenderId, desc, false, user.name, assigneeEmails||'', '', '', new Date(), taskId, false]);
    SpreadsheetApp.flush();
    const tenderData = getSheet(SHEET_TENDERS).getDataRange().getValues();
    let tenderTitle = '';
    for (let i = 1; i < tenderData.length; i++) { if (tenderData[i][0]===tenderId) { tenderTitle=tenderData[i][1]; break; } }
    logActivity(user.email, user.name, tenderId, tenderTitle, 'TASK_CREATED', desc);
    if (assigneeEmails) {
      assigneeEmails.split(',').forEach(email => {
        const e = email.trim();
        createNotification(e, 'TASK_ASSIGNED', tenderId, user.name+' assigned you a task in '+tenderTitle, '#/tender/'+tenderId);
        sendNotification(e, '[TenderFlow] New Task: '+tenderId, 'Assigned by '+user.name+': '+desc);
        sendTelegramByEmail(e, '📋 *New Task Assigned*\n\nProject: *'+tenderTitle+'*\nTask: '+desc+'\nBy: '+user.name);
      });
    }
    return getTasks(token, tenderId);
  } catch (error) { throw new Error('Failed to add task: ' + error.message); }
  finally { lock.releaseLock(); }
}

function toggleTask(token, taskId, isDone) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    authenticate(token);
    const sheet = getSheet(SHEET_TASKS);
    const row   = findTaskRow_(sheet, taskId);
    if (row === -1) throw new Error('Task not found: ' + taskId);
    sheet.getRange(row, 3).setValue(isDone);
    SpreadsheetApp.flush();
    const tid = sheet.getRange(row, 1).getValue();
    return { tid, tasks: getTasks(token, tid) };
  } catch (error) { throw new Error('Failed to toggle task: ' + error.message); }
  finally { lock.releaseLock(); }
}

function updateTaskAssignee(token, taskId, newAssignee) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const user  = authenticate(token);
    const sheet = getSheet(SHEET_TASKS);
    const row   = findTaskRow_(sheet, taskId);
    if (row === -1) throw new Error('Task not found: ' + taskId);
    const oldAssignee = sheet.getRange(row, 5).getValue() || '';
    const tid         = sheet.getRange(row, 1).getValue();
    const taskDesc    = sheet.getRange(row, 2).getValue();
    sheet.getRange(row, 5).setValue(newAssignee||'');
    SpreadsheetApp.flush();
    const tenderData = getSheet(SHEET_TENDERS).getDataRange().getValues();
    let tenderTitle = '';
    for (let i = 1; i < tenderData.length; i++) { if (tenderData[i][0]===tid) { tenderTitle=tenderData[i][1]; break; } }
    if (newAssignee) {
      const oldEmails = oldAssignee ? oldAssignee.split(',').map(e=>e.trim()) : [];
      newAssignee.split(',').forEach(email => {
        const e = email.trim();
        if (!oldEmails.includes(e)) {
          createNotification(e, 'TASK_ASSIGNED', tid, user.name+' assigned you a task in '+tenderTitle, '#/tender/'+tid);
          sendTelegramByEmail(e, '📋 *Task Assigned*\n\nProject: *'+tenderTitle+'*\nTask: '+taskDesc+'\nBy: '+user.name);
        }
      });
    }
    return getTasks(token, tid);
  } catch (error) { throw new Error('Failed to update assignee: ' + error.message); }
  finally { lock.releaseLock(); }
}

function deleteTask(token, taskId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    authenticate(token);
    const sheet = getSheet(SHEET_TASKS);
    const row   = findTaskRow_(sheet, taskId);
    if (row === -1) throw new Error('Task not found: ' + taskId);
    const tid = sheet.getRange(row, 1).getValue();
    sheet.getRange(row, 10).setValue(true);
    SpreadsheetApp.flush();
    return getTasks(token, tid);
  } catch (error) { throw new Error('Failed to delete task: ' + error.message); }
  finally { lock.releaseLock(); }
}

function uploadTaskFile(token, taskId, fileData, fileName, mimeType) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    authenticate(token);
    if (fileData.length * 0.75 > 5*1024*1024) throw new Error('File too large (>5MB)');
    const sheet  = getSheet(SHEET_TASKS);
    const row    = findTaskRow_(sheet, taskId);
    if (row === -1) throw new Error('Task not found: ' + taskId);
    const tenderId = sheet.getRange(row, 1).getValue();
    let parentFolderId = UPLOAD_FOLDER_ID;
    const tData = getSheet(SHEET_TENDERS).getDataRange().getValues();
    for (let i = 1; i < tData.length; i++) { if (tData[i][0]===tenderId && tData[i][14]) { const m=tData[i][14].match(/[-\w]{25,}/); if(m) parentFolderId=m[0]; break; } }
    const folder = DriveApp.getFolderById(parentFolderId);
    const blob   = Utilities.newBlob(Utilities.base64Decode(fileData), mimeType, fileName);
    const file   = folder.createFile(blob);
    sheet.getRange(row, 6).setValue(file.getUrl());
    sheet.getRange(row, 7).setValue(fileName);
    SpreadsheetApp.flush();
    return getTasks(token, tenderId);
  } catch (error) { throw new Error('Upload failed: ' + error.message); }
  finally { lock.releaseLock(); }
}

// ================= COMMENTS FUNCTIONS =================

function getComments(token, tenderId) {
  try {
    authenticate(token);
    const sheet = getSheet(SHEET_COMMENTS);
    if (sheet.getLastRow() <= 1) return [];
    return sheet.getDataRange().getValues().slice(1)
      .filter(r => r[0]===tenderId)
      .map(r => ({ userEmail: r[1], userName: r[2], time: Utilities.formatDate(new Date(r[3]), Session.getScriptTimeZone(), 'dd/MM HH:mm'), msg: r[4] }));
  } catch (error) { throw new Error('Failed to load comments: ' + error.message); }
}

function addComment(token, tenderId, msg) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const user = authenticate(token);
    if (!msg || !msg.trim()) throw new Error('Comment cannot be empty');
    getSheet(SHEET_COMMENTS).appendRow([tenderId, user.email, user.name, new Date(), msg]);
    SpreadsheetApp.flush();
    const tData = getSheet(SHEET_TENDERS).getDataRange().getValues();
    let tenderTitle = '', teamMembers = '';
    for (let i = 1; i < tData.length; i++) { if (tData[i][0]===tenderId) { tenderTitle=tData[i][1]; teamMembers=tData[i][16]; break; } }
    logActivity(user.email, user.name, tenderId, tenderTitle, 'COMMENT_ADDED', msg.substring(0,50));
    const mentions = msg.match(/([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)/gi);
    if (mentions) {
      [...new Set(mentions)].forEach(target => {
        createNotification(target, 'MENTIONED', tenderId, user.name+' mentioned you in '+tenderTitle, '#/tender/'+tenderId);
        sendNotification(target, '[TenderFlow] Mentioned in '+tenderId, '<b>'+user.name+'</b> mentioned you: "'+msg+'"');
        sendTelegramByEmail(target, '💬 *You were mentioned!*\n\nProject: *'+tenderTitle+'*\nBy: '+user.name+'\n'+msg.replace(/<[^>]*>/g,''));
      });
    }
    if (teamMembers) {
      teamMembers.split(',').forEach(email => {
        if (email.trim()!==user.email) {
          createNotification(email.trim(), 'COMMENT', tenderId, user.name+' commented on '+tenderTitle, '#/tender/'+tenderId);
          sendTelegramByEmail(email.trim(), '💬 *New Comment*\n\nProject: *'+tenderTitle+'*\n'+user.name+': '+msg.replace(/<[^>]*>/g,'').substring(0,100));
        }
      });
    }
    return getComments(token, tenderId);
  } catch (error) { throw new Error('Failed to add comment: ' + error.message); }
  finally { lock.releaseLock(); }
}

// ================= DIRECT MESSAGES =================

function sendDirectMessage(token, toEmail, text) {
  try {
    const user = authenticate(token);
    if (!text || !text.trim()) return { success: false, msg: 'Message cannot be empty' };
    if (!toEmail)              return { success: false, msg: 'Recipient required' };

    const sheet = getSheet(SHEET_DM);

    const isFirstMessage = sheet.getLastRow() <= 1 ||
      !sheet.getDataRange().getValues().slice(1).some(r =>
        (r[1] === user.email && r[3] === toEmail) ||
        (r[1] === toEmail    && r[3] === user.email)
      );

    sheet.appendRow([new Date(), user.email, user.name, toEmail, text.trim()]);
    SpreadsheetApp.flush();

    try { sendTelegramByEmail(toEmail, '💬 *DM from ' + user.name + '*\n\n' + text.trim()); } catch (e) {}

    try {
      const cacheKey    = 'dm_email_' + user.email + '_' + toEmail;
      const alreadySent = CacheService.getScriptCache().get(cacheKey);

      if ((isFirstMessage || !alreadySent) && MailApp.getRemainingDailyQuota() > 0) {
        const appUrl = ScriptApp.getService().getUrl();

        MailApp.sendEmail({
          to:      toEmail,
          subject: '💬 New message from @' + user.name + ' — AxiCom Tender System',
          htmlBody:
            '<div style="font-family:\'Helvetica Neue\',Arial,sans-serif;max-width:520px;margin:0 auto;padding:24px 16px;color:#0a0c10">' +
              '<div style="background:#0a0c10;border-radius:10px;padding:16px 20px;margin-bottom:22px;display:flex;align-items:center;gap:10px">' +
                '<span style="font-size:18px">⚡</span>' +
                '<span style="color:#ffffff;font-size:15px;font-weight:700;letter-spacing:-.3px">TenderFlow</span>' +
                '<span style="color:rgba(255,255,255,.35);font-size:12px;margin-left:auto">AxiCom Tender System</span>' +
              '</div>' +
              '<h2 style="font-size:18px;font-weight:700;margin:0 0 6px;color:#0a0c10">You have a new message</h2>' +
              '<p style="font-size:13px;color:#8a91a8;margin:0 0 20px">' +
                'From <b style="color:#0a0c10">@' + user.name + '</b> · via AxiCom Tender System' +
              '</p>' +
              '<div style="background:#f7f8fc;border:1px solid #e8eaf0;border-left:3px solid #2563eb;border-radius:8px;padding:14px 16px;margin-bottom:24px">' +
                '<p style="font-size:14px;line-height:1.65;margin:0;color:#1c2030;word-break:break-word">' +
                  dmEscapeHtml_(text.trim()) +
                '</p>' +
              '</div>' +
              '<a href="' + appUrl + '" ' +
                'style="display:inline-block;background:#0a0c10;color:#ffffff;text-decoration:none;' +
                'padding:11px 24px;border-radius:8px;font-size:13px;font-weight:600;letter-spacing:-.1px">' +
                'Open TenderFlow &rarr;' +
              '</a>' +
              '<p style="font-size:11px;color:#b8bdd0;margin-top:24px;line-height:1.6">' +
                'You received this because you have an account on AxiCom Tender System.<br>' +
                'Please reply directly inside the app — do not reply to this email.' +
              '</p>' +
            '</div>'
        });

        CacheService.getScriptCache().put(cacheKey, '1', 1800);
      }
    } catch (emailErr) {
      Logger.log('DM email notification failed: ' + emailErr.toString());
    }

    return { success: true };
  } catch (e) {
    Logger.log('sendDirectMessage error: ' + e.toString());
    return { success: false, msg: e.message };
  }
}

function dmEscapeHtml_(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/\n/g, '<br>');
}

function getDirectMessages(token, peerEmail) {
  try {
    const user = authenticate(token);
    if (!peerEmail) return [];
    const sheet = getSheet(SHEET_DM);
    if (sheet.getLastRow() <= 1) return [];
    const me   = user.email;
    const rows = sheet.getDataRange().getValues().slice(1);
    return rows
      .filter(r => (r[1]===me && r[3]===peerEmail) || (r[1]===peerEmail && r[3]===me))
      .slice(-80)
      .map(r => ({
        time:      Utilities.formatDate(new Date(r[0]), Session.getScriptTimeZone(), 'dd/MM HH:mm'),
        fromEmail: r[1], fromName: r[2], toEmail: r[3], text: r[4]
      }));
  } catch (e) {
    Logger.log('getDirectMessages error: ' + e.toString());
    return [];
  }
}

// ================= LIBRARY FUNCTIONS =================

function getLibraryFiles(token) {
  try {
    authenticate(token);
    if (!LIBRARY_FOLDER_ID || LIBRARY_FOLDER_ID.length < 20) return [];
    const folder = DriveApp.getFolderById(LIBRARY_FOLDER_ID);
    const files  = folder.getFiles();
    const result = [];
    while (files.hasNext()) { const f = files.next(); result.push({ id: f.getId(), name: f.getName(), url: f.getUrl(), size: (f.getSize()/1024/1024).toFixed(2)+' MB' }); }
    return result;
  } catch (error) { return []; }
}

function uploadTemplate(token, data, name, type) {
  try {
    authenticate(token);
    if (data.length*0.75 > 5*1024*1024) throw new Error('File too large (>5MB)');
    DriveApp.getFolderById(LIBRARY_FOLDER_ID).createFile(Utilities.newBlob(Utilities.base64Decode(data), type, name));
    return getLibraryFiles(token);
  } catch (error) { throw new Error('Upload failed: ' + error.message); }
}

function deleteFile(token, id) {
  try { authenticate(token); DriveApp.getFileById(id).setTrashed(true); return getLibraryFiles(token); }
  catch (error) { throw new Error('Delete failed: ' + error.message); }
}

function getBulkData(token) {
  try {
    const user         = authenticate(token);
    const tenders      = getTenders(token);
    const taskSheet    = getSheet(SHEET_TASKS);
    const allTaskRows  = taskSheet.getLastRow() > 1 ? taskSheet.getDataRange().getValues().slice(1) : [];
    const taskData     = allTaskRows.filter(r => r[8] && r[9] !== true);
    const taskProgress = {};
    taskData.forEach(r => {
      const tid = r[0];
      if (!taskProgress[tid]) taskProgress[tid] = { done:0, total:0, pct:0 };
      taskProgress[tid].total++;
      if (r[2]) taskProgress[tid].done++;
    });
    Object.keys(taskProgress).forEach(tid => { const t=taskProgress[tid]; t.pct=t.total?Math.round((t.done/t.total)*100):0; });

    const tenderMap = {};
    tenders.forEach(t => { tenderMap[t.id] = { title: t.title, status: t.status, dueDate: t.dueDate }; });
    const myTasks = taskData.map((r, i) => {
      const info = tenderMap[r[0]] || { title:'Unknown', status:'Unknown', dueDate:'9999-12-31' };
      return { rowIndex: i+2, tid: r[0], desc: r[1], done: r[2], assignee: r[4]?r[4].toString():'', tenderTitle: info.title, tenderStatus: info.status, tenderDueDate: info.dueDate };
    }).filter(t => {
      const assignees = t.assignee.split(',').map(e=>e.trim());
      return assignees.includes(user.email) && !['Submitted','Closed','Unsuccessful','Cancelled','Successful'].includes(t.tenderStatus);
    });

    const notifSheet   = getSheet(SHEET_NOTIFICATIONS);
    const unreadCount  = notifSheet.getLastRow() > 1
      ? notifSheet.getDataRange().getValues().slice(1).filter(r => r[0]===user.email && !r[4]).length
      : 0;

    const actSheet  = getSheet(SHEET_ACTIVITY);
    const actData   = actSheet.getLastRow() > 1 ? actSheet.getDataRange().getValues().slice(1) : [];
    const actFeed   = actData
      .sort((a,b) => new Date(b[0]) - new Date(a[0]))
      .slice(0, 20)
      .map(r => ({
        timestamp:   Utilities.formatDate(new Date(r[0]), Session.getScriptTimeZone(), 'dd/MM HH:mm'),
        userEmail:   r[1], userName: r[2], tenderId: r[3], tenderTitle: r[4], action: r[5], details: r[6]
      }));

    const usersSheet = getSheet(SHEET_USERS);
    const cache      = CacheService.getScriptCache();
    const users      = usersSheet.getLastRow() > 1
      ? usersSheet.getDataRange().getValues().slice(1).map(r => ({
          email: r[0], name: r[1], role: r[2],
          isOnline: cache.get('online_' + r[0]) === 'true',
          hasTelegram: !!r[8]
        }))
      : [];

    const active  = tenders.filter(t => ['Calling For Tender','Active','Review','Submitted'].includes(t.status));
    const stats   = {
      total:   tenders.length,
      active:  active.length,
      value:   active.reduce((s,t) => s + (t.submissionPrice||0), 0),
      success: tenders.filter(t => t.status==='Successful').length,
      recent:  tenders.slice(0,5)
    };

    return { tenders, stats, tasks: myTasks, users, taskProgress, activityFeed: actFeed, unreadCount };
  } catch (error) { throw new Error('Failed to load data: ' + error.message); }
}

// ================= NOTIFICATIONS =================

function createNotification(userEmail, type, tenderId, message, actionUrl) {
  try { getSheet(SHEET_NOTIFICATIONS).appendRow([userEmail, type, tenderId, message, false, new Date(), actionUrl||'']); SpreadsheetApp.flush(); }
  catch (e) { Logger.log('createNotification failed: ' + e.toString()); }
}

function getNotifications(token, unreadOnly) {
  try {
    const user  = authenticate(token);
    const sheet = getSheet(SHEET_NOTIFICATIONS);
    if (sheet.getLastRow() <= 1) return [];
    return sheet.getDataRange().getValues().slice(1)
      .filter(r => r[0]===user.email && (!unreadOnly||!r[4]))
      .map((r,i) => ({ rowIndex:i+2, type:r[1], tenderId:r[2], message:r[3], isRead:r[4], timestamp:Utilities.formatDate(new Date(r[5]),Session.getScriptTimeZone(),'dd/MM/yyyy HH:mm'), actionUrl:r[6] }))
      .sort((a,b)=>new Date(b.timestamp)-new Date(a.timestamp)).slice(0, 50);
  } catch (error) { return []; }
}

function markNotificationRead(token, rowIndex) {
  try { authenticate(token); getSheet(SHEET_NOTIFICATIONS).getRange(rowIndex,5).setValue(true); SpreadsheetApp.flush(); return { success:true }; }
  catch (error) { return { success:false }; }
}

function markAllNotificationsRead(token) {
  try {
    const user  = authenticate(token);
    const sheet = getSheet(SHEET_NOTIFICATIONS);
    const data  = sheet.getDataRange().getValues();
    for (let i=1; i<data.length; i++) { if(data[i][0]===user.email&&!data[i][4]) sheet.getRange(i+1,5).setValue(true); }
    SpreadsheetApp.flush(); return { success:true };
  } catch (error) { return { success:false }; }
}

function getUnreadNotificationCount(token) {
  try {
    const user  = authenticate(token);
    const sheet = getSheet(SHEET_NOTIFICATIONS);
    if (sheet.getLastRow()<=1) return 0;
    return sheet.getDataRange().getValues().slice(1).filter(r=>r[0]===user.email&&!r[4]).length;
  } catch (error) { return 0; }
}

// ================= ACTIVITY FEED =================

function logActivity(userEmail, userName, tenderId, tenderTitle, action, details) {
  try { getSheet(SHEET_ACTIVITY).appendRow([new Date(), userEmail, userName, tenderId, tenderTitle, action, details||'']); SpreadsheetApp.flush(); }
  catch (e) { Logger.log('logActivity failed: ' + e.toString()); }
}

function getActivityFeed(token, tenderId, limit) {
  try {
    authenticate(token);
    const sheet = getSheet(SHEET_ACTIVITY);
    if (sheet.getLastRow()<=1) return [];
    let filtered = sheet.getDataRange().getValues().slice(1);
    if (tenderId) filtered = filtered.filter(r=>r[3]===tenderId);
    filtered.sort((a,b)=>new Date(b[0])-new Date(a[0]));
    if (limit) filtered = filtered.slice(0, limit);
    return filtered.map(r => ({ timestamp:Utilities.formatDate(new Date(r[0]),Session.getScriptTimeZone(),'dd/MM HH:mm'), userEmail:r[1], userName:r[2], tenderId:r[3], tenderTitle:r[4], action:r[5], details:r[6] }));
  } catch (error) { return []; }
}

// ================= PRESENCE =================

function updateUserPresence(token, tenderId, status) {
  try {
    const user = authenticate(token);
    CacheService.getScriptCache().put('presence_'+tenderId+'_'+user.email, JSON.stringify({ userName:user.name, email:user.email, status, timestamp:new Date().getTime() }), 300);
    return { success:true };
  } catch (error) { return { success:false }; }
}

function getActiveUsers(token, tenderId) {
  try {
    authenticate(token);
    const cache    = CacheService.getScriptCache();
    const allUsers = getAllUsers(token);
    const active   = [];
    allUsers.forEach(u => {
      const presence = cache.get('presence_'+tenderId+'_'+u.email);
      if (presence) { const d=JSON.parse(presence); if(new Date().getTime()-d.timestamp<120000) active.push({ name:d.userName, email:d.email, status:d.status }); }
    });
    return active;
  } catch (error) { return []; }
}

// ================= FILE VERSIONS =================

function getFileVersions(token, tenderId) {
  try {
    authenticate(token);
    const tData = getSheet(SHEET_TENDERS).getDataRange().getValues();
    let folderUrl = '';
    for (let i=1; i<tData.length; i++) { if(tData[i][0]===tenderId) { folderUrl=tData[i][14]; break; } }
    if (!folderUrl) return [];
    const folderId = folderUrl.match(/[-\w]{25,}/);
    if (!folderId) return [];
    const folder = DriveApp.getFolderById(folderId[0]);
    const files  = folder.getFiles();
    const list   = [];
    while (files.hasNext()) {
      const f = files.next();
      list.push({ id:f.getId(), name:f.getName(), url:f.getUrl(), size:(f.getSize()/1024/1024).toFixed(2)+' MB', lastModified:Utilities.formatDate(f.getLastUpdated(),Session.getScriptTimeZone(),'dd/MM/yyyy HH:mm'), modifiedBy:f.getOwner().getEmail() });
    }
    return list.sort((a,b)=>new Date(b.lastModified)-new Date(a.lastModified));
  } catch (error) { return []; }
}

// ================= HELPERS =================

function notifyTeamTelegram(teamMembers, actorName, tenderTitle, tenderId, message) {
  if (!teamMembers) return;
  teamMembers.split(',').forEach(email => { sendTelegramByEmail(email.trim(), message); });
}

// ================= DUE DATE REMINDERS =================

function sendDueDateReminders() {
  try {
    const data  = getSheet(SHEET_TENDERS).getDataRange().getValues();
    const now   = new Date();
    const uData = getSheet(SHEET_USERS).getDataRange().getValues();
    for (let i=1; i<data.length; i++) {
      if (data[i][17]) continue;
      const dueDate = new Date(data[i][3]);
      const status  = data[i][13];
      if (['Submitted','Successful','Closed','Unsuccessful','Cancelled'].includes(status)) continue;
      const daysLeft = Math.ceil((dueDate-now)/86400000);
      if (![7,3,1].includes(daysLeft)) continue;
      const tenderId=data[i][0], title=data[i][1], owner=data[i][15], team=data[i][16];
      const msg   = '⏰ *Deadline Reminder*\n\n*'+title+'*\nDue in *'+daysLeft+' day'+(daysLeft>1?'s':'')+'*\nStatus: '+status;
      const inApp = 'Reminder: "'+title+'" is due in '+daysLeft+' day'+(daysLeft>1?'s':'');
      if (owner) { for (let j=1;j<uData.length;j++) { if(uData[j][1]===owner) { createNotification(uData[j][0],'DUE_SOON',tenderId,inApp,'#/tender/'+tenderId); sendTelegramByEmail(uData[j][0],msg); break; } } }
      if (team) { team.split(',').forEach(email=>{ const e=email.trim(); createNotification(e,'DUE_SOON',tenderId,inApp,'#/tender/'+tenderId); sendTelegramByEmail(e,msg); }); }
    }
    Logger.log('Due date reminders sent.');
  } catch (error) { Logger.log('sendDueDateReminders error: ' + error.toString()); }
}

function setupDailyReminders() {
  ScriptApp.getProjectTriggers().forEach(t=>{ if(t.getHandlerFunction()==='sendDueDateReminders') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('sendDueDateReminders').timeBased().everyDays(1).atHour(9).create();
  Logger.log('Daily reminder trigger created.');
}

// ================= TEAM METRICS =================

function getTeamMetrics(token) {
  try {
    authenticate(token);
    const actData    = getSheet(SHEET_ACTIVITY).getDataRange().getValues();
    const taskData   = getSheet(SHEET_TASKS).getDataRange().getValues();
    const tenderData = getSheet(SHEET_TENDERS).getDataRange().getValues();
    const metrics    = { totalProjects:tenderData.length-1, activeProjects:tenderData.slice(1).filter(r=>!r[17]&&['Calling For Tender','Active','Review'].includes(r[13])).length,
                         completedTasks:taskData.slice(1).filter(r=>r[2]&&r[8]&&r[9]!==true).length,
                         pendingTasks:taskData.slice(1).filter(r=>!r[2]&&r[8]&&r[9]!==true).length,
                         teamActivity:{}, projectsByStatus:{} };
    actData.slice(1).forEach(r=>{ const n=r[2]; metrics.teamActivity[n]=(metrics.teamActivity[n]||0)+1; });
    tenderData.slice(1).forEach(r=>{ const s=r[13]; metrics.projectsByStatus[s]=(metrics.projectsByStatus[s]||0)+1; });
    return metrics;
  } catch (error) { Logger.log('getTeamMetrics error: '+error.toString()); throw new Error('Failed to get metrics: '+error.message); }
}

// ================= ADMIN & MAINTENANCE =================

function cleanupExpiredSessions() {
  try {
    const sheet=getSheet(SHEET_USERS), data=sheet.getDataRange().getValues(), now=new Date(); let cleaned=0;
    for (let i=1; i<data.length; i++) { if(new Date(data[i][5])<now) { sheet.getRange(i+1,5).setValue(''); sheet.getRange(i+1,6).setValue(''); cleaned++; } }
    SpreadsheetApp.flush(); return { success:true, cleaned };
  } catch (error) { return { success:false, error:error.message }; }
}

function cleanupExpiredResetCodes() {
  try {
    const sheet=getSheet(SHEET_RESET), data=sheet.getDataRange().getValues(), now=new Date(), toDelete=[];
    for (let i=data.length-1; i>=1; i--) { if(new Date(data[i][2])<now||data[i][3]) toDelete.push(i+1); }
    toDelete.forEach(row=>sheet.deleteRow(row)); SpreadsheetApp.flush(); return { success:true, cleaned:toDelete.length };
  } catch (error) { return { success:false, error:error.message }; }
}

function fixOrphanTaskRows() {
  const sheet = getSheet(SHEET_TASKS);
  const data  = sheet.getDataRange().getValues();
  let fixed   = 0;
  for (let i = 1; i < data.length; i++) {
    const hasTaskId  = data[i][8] && data[i][8].toString().trim() !== '';
    const isDeleted  = data[i][9];
    if (!hasTaskId && isDeleted !== true) {
      sheet.getRange(i + 1, 10).setValue(true);
      fixed++;
    }
  }
  SpreadsheetApp.flush();
  Logger.log('fixOrphanTaskRows: marked ' + fixed + ' orphan rows as deleted.');
  return 'Fixed ' + fixed + ' orphan task rows.';
}

function getAuditLog(token, tenderId, limit) {
  try {
    authenticate(token);
    const sheet = getSheet(SHEET_LOGS);
    if (sheet.getLastRow()<=1) return [];
    let data = sheet.getDataRange().getValues().slice(1);
    if (tenderId) data = data.filter(r=>r[2]===tenderId);
    data.sort((a,b)=>new Date(b[0])-new Date(a[0]));
    if (limit) data = data.slice(0, limit);
    return data.map(r=>({ timestamp:Utilities.formatDate(new Date(r[0]),Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm:ss'), user:r[1], tenderId:r[2], action:r[3], oldValue:r[4], newValue:r[5] }));
  } catch (error) { throw new Error('Failed to get audit log: '+error.message); }
}

function runDailyMaintenance() {
  Logger.log('=== DAILY MAINTENANCE ===');
  try {
    Logger.log('Sessions cleaned: '+cleanupExpiredSessions().cleaned);
    Logger.log('Reset codes cleaned: '+cleanupExpiredResetCodes().cleaned);
    sendDueDateReminders(); Logger.log('Reminders sent');
    CacheService.getScriptCache().removeAll(['tenders_all','stats_all']); Logger.log('Cache cleared');
    Logger.log('=== COMPLETE ==='); return 'Maintenance completed successfully';
  } catch (error) { return 'Maintenance failed: '+error.message; }
}

function setupMaintenanceTrigger() {
  ScriptApp.getProjectTriggers().forEach(t=>{ if(t.getHandlerFunction()==='runDailyMaintenance') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('runDailyMaintenance').timeBased().everyDays(1).atHour(2).create();
  return 'Daily maintenance scheduled for 2:00 AM';
}

function testEmailConfiguration() {
  try {
    if (MailApp.getRemainingDailyQuota()<=0) return 'ERROR: Daily email quota exceeded';
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) return 'ERROR: Cannot determine user email.';
    MailApp.sendEmail({ to:userEmail, subject:'TenderFlow Email Test', htmlBody:'<h1>✅ Email Working!</h1>' });
    return 'SUCCESS: Test email sent to '+userEmail;
  } catch (error) { return 'ERROR: '+error.toString(); }
}

function getSystemStats() {
  try {
    const s = { users:getSheet(SHEET_USERS).getLastRow()-1, tenders:getSheet(SHEET_TENDERS).getLastRow()-1,
                tasks:getSheet(SHEET_TASKS).getLastRow()-1, comments:getSheet(SHEET_COMMENTS).getLastRow()-1,
                notifications:getSheet(SHEET_NOTIFICATIONS).getLastRow()-1,
                directMessages:getSheet(SHEET_DM).getLastRow()-1,
                emailQuota:MailApp.getRemainingDailyQuota() };
    Logger.log(JSON.stringify(s)); return s;
  } catch (error) { return 'Error: '+error.message; }
}

function forcePasswordReset(email, newPassword) {
  try {
    if (!email||!newPassword||newPassword.length<6) return 'Error: Invalid input';
    const sheet=getSheet(SHEET_USERS), data=sheet.getDataRange().getValues();
    for (let i=1; i<data.length; i++) {
      if (data[i][0]===email) { sheet.getRange(i+1,4).setValue(hashPassword(newPassword,email)); sheet.getRange(i+1,5).setValue(''); sheet.getRange(i+1,6).setValue(''); SpreadsheetApp.flush(); return 'Password reset for '+email; }
    }
    return 'User not found: '+email;
  } catch (error) { return 'Error: '+error.message; }
}

function clearAllNotifications() {
  try {
    const sheet = getSheet(SHEET_NOTIFICATIONS);
    if (sheet.getLastRow()>1) { sheet.deleteRows(2, sheet.getLastRow()-1); SpreadsheetApp.flush(); }
    return 'Notifications cleared';
  } catch (error) { return 'Error: '+error.message; }
}

function runFullSystemTest() {
  Logger.log('=== TENDERFLOW SYSTEM TEST ===');
  try {
    Logger.log('1. Database setup...');
    setupDatabase();
    Logger.log('✓ All ' + Object.keys(HEADERS).length + ' sheets OK (including DirectMessages)');
    Logger.log('2. Email...'); Logger.log(testEmailConfiguration());
    Logger.log('3. Telegram...');
    const tgRes  = UrlFetchApp.fetch('https://api.telegram.org/bot'+TELEGRAM_BOT_TOKEN+'/getMe', { muteHttpExceptions:true });
    const tgData = JSON.parse(tgRes.getContentText());
    Logger.log(tgData.ok ? '✓ Telegram OK: @'+tgData.result.username : '✗ Telegram FAIL: '+tgData.description);
    Logger.log('4. Drive...');
    if (UPLOAD_FOLDER_ID&&UPLOAD_FOLDER_ID.length>20) Logger.log('✓ Drive: '+DriveApp.getFolderById(UPLOAD_FOLDER_ID).getName());
    else Logger.log('⚠ Drive folder not configured');
    Logger.log('5. DirectMessages rows: '+(getSheet(SHEET_DM).getLastRow()-1));
    Logger.log('=== TEST COMPLETE ===');
    return 'All tests done. Check Execution Log.';
  } catch (error) { return 'Test failed: '+error.message; }
}

// ═══════════════════════════════════════════════════════
//  TENDER FILE UPLOAD & LIST
//  Add these functions to your existing Code.gs
// ═══════════════════════════════════════════════════════

/**
 * Upload a file directly into the tender's Drive folder.
 * Called from client: uploadTenderFile(TOKEN, tenderId, base64, fileName, mimeType)
 */
function uploadTenderFile(token, tenderId, base64Data, fileName, mimeType) {
  const user = verifyToken_(token);          // ← your existing auth helper
  if (!user) throw new Error('Session expired');

  const tender = getTenderById_(tenderId);   // ← see helper below
  if (!tender) throw new Error('Project not found');
  if (!tender.folderId) throw new Error('No Drive folder linked to this project');

  // Decode and write to Drive
  const bytes  = Utilities.base64Decode(base64Data);
  const blob   = Utilities.newBlob(bytes, mimeType || 'application/octet-stream', fileName);
  const folder = DriveApp.getFolderById(tender.folderId);
  folder.createFile(blob);

  // Log activity (safe — ignores if your logActivity_ signature differs)
  try {
    logActivity_(user, tenderId, 'FILE_UPLOAD', 'Uploaded ' + fileName);
  } catch(e) { /* non-fatal */ }

  // Return refreshed file list so the UI can update immediately
  return listTenderFiles_(tenderId);
}

/**
 * Get list of files in a tender's Drive folder.
 * Called from client: getTenderFiles(TOKEN, tenderId)
 */
function getTenderFiles(token, tenderId) {
  const user = verifyToken_(token);
  if (!user) throw new Error('Session expired');
  return listTenderFiles_(tenderId);
}

/**
 * Internal helper — builds the file list object for a given tender.
 * Also used by getFileVersions if you want to unify them (see below).
 */
function listTenderFiles_(tenderId) {
  const tender = getTenderById_(tenderId);
  if (!tender || !tender.folderId) return [];

  const folder = DriveApp.getFolderById(tender.folderId);
  const iter   = folder.getFiles();
  const tz     = Session.getScriptTimeZone() || 'UTC';
  const out    = [];

  while (iter.hasNext()) {
    const f = iter.next();
    out.push({
      id           : f.getId(),
      name         : f.getName(),
      mimeType     : f.getMimeType(),
      size         : formatFileSize_(f.getSize()),
      lastModified : Utilities.formatDate(f.getLastUpdated(), tz, 'yyyy-MM-dd HH:mm'),
      // previewUrl used by iframe; viewUrl for "Open in Drive" link
      previewUrl   : 'https://drive.google.com/file/d/' + f.getId() + '/preview',
      viewUrl      : 'https://drive.google.com/file/d/' + f.getId() + '/view'
    });
  }

  // Newest first
  out.sort((a, b) => b.lastModified.localeCompare(a.lastModified));
  return out;
}

/**
 * Internal helper — formats bytes into human-readable size.
 */
function formatFileSize_(bytes) {
  if (bytes < 1024)        return bytes + ' B';
  if (bytes < 1024 * 1024) return Math.round(bytes / 1024) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

/**
 * Internal helper — retrieve a single tender row by ID.
 *
 * ── IMPORTANT ──
 * Replace the sheet/column names below with whatever you actually use.
 * The key fields you MUST return are:
 *   { id, folderId }
 * folderId is the Google Drive folder ID stored with each tender.
 * If you store the full folderUrl instead, extract the ID like:
 *   folderId = folderUrl.match(/[-\w]{25,}/)?.[0]
 */
function getTenderById_(tenderId) {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const sh     = ss.getSheetByName('Tenders');  // ← change to your sheet name
  const values = sh.getDataRange().getValues();
  const header = values[0];

  // ── adjust these column names to match your sheet headers ──
  const idCol       = header.indexOf('ID');
  const folderCol   = header.indexOf('FolderUrl');  // or 'FolderId'
  const titleCol    = header.indexOf('Title');

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idCol]) === String(tenderId)) {
      let folderId = String(values[i][folderCol] || '');
      // If you stored the full URL, extract just the ID
      if (folderId.startsWith('http')) {
        const m = folderId.match(/[-\w]{25,}/);
        folderId = m ? m[0] : '';
      }
      return {
        id       : values[i][idCol],
        title    : values[i][titleCol],
        folderId : folderId
      };
    }
  }
  return null;
}

/**
 * Optional: keep getFileVersions pointing to the same source
 * so any other callers still work.
 */
function getFileVersions(token, tenderId) {
  return getTenderFiles(token, tenderId);
}
// ================= END OF Code.gs =================
