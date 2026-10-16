/**
 * =========================================================================
 * WORK TRACKER (โปรแกรมติดตามงาน)
 * Backend Google Apps Script (Code.js)
 * รองรับการทำงานร่วมกับ Google Sheets, Web Application, และ Telegram Notify
 * =========================================================================
 */

// ชื่อแผ่นงานหลัก
const SHEET_TASKS = 'Tasks';
const SHEET_SETTINGS = 'Settings';
const SHEET_USERS = 'Users';
const SHEET_ARCHIVE = 'Tasks_Archive';

// สถานะงานมาตรฐาน
const TASK_STATUSES = ['รอดำเนินการ', 'กำลังดำเนินการ', 'รอตรวจสอบ', 'เสร็จสิ้น'];

// ระดับความสำคัญ
const TASK_PRIORITIES = ['ด่วนมาก', 'สูง', 'ปานกลาง', 'ทั่วไป'];

// ไซต์งานเริ่มต้น (หากยังไม่มีการกำหนด)
const DEFAULT_SITES = [
  'ไซต์ A - อาคารสำนักงานใหญ่',
  'ไซต์ B - โครงการคอนโดสุขุมวิท',
  'ไซต์ C - โรงงานนิคมชลบุรี',
  'งานซ่อมบำรุงและบริการทั่วไป'
];

/**
 * =========================================================================
 * ฟังก์ชันสำหรับให้ผู้ดูแลระบบกด "เรียกใช้ (Run)" ใน Google Apps Script Editor ครั้งแรก
 * เพื่อยินยอมสิทธิ์การเข้าถึง Google Drive (One-time OAuth Authorization)
 * ลิงก์ Editor: https://script.google.com/d/1cd4_bnt_nnuKTQzu3F6gDmMpwWoeBBcuNyqPCn3lDBsAfmguacW1aUBE/edit
 * =========================================================================
 */
function authorizeDriveApp() {
  const root = getAttachmentsRootFolder_();
  console.log('✅ ยืนยันสิทธิ์ Google Drive สำเร็จ! โฟลเดอร์ราก:', root.getName(), 'ID:', root.getId());
  return 'ยืนยันสิทธิ์ Google Drive สำเร็จเรียบร้อยแล้ว!';
}

/**
 * ให้บริการหน้าเว็บหลัก (Web Application)
 */
function doGet(e) {
  try {
    // บริการ JSON API สำหรับ Health Check / Diagnostic
    if (e && e.parameter && e.parameter.action) {
      const action = e.parameter.action;
      let result = null;
      if (action === 'ping') {
        result = { success: true, message: 'Work Tracker API is online', timestamp: Date.now() };
      } else if (action === 'getInitialData') {
        result = getInitialData(e.parameter.username || '');
      } else if (action === 'checkTasksUpdate') {
        result = checkTasksUpdate(e.parameter.version || '', e.parameter.username || '');
      }
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ตรวจสอบและเตรียมโครงสร้างฐานข้อมูลอัตโนมัติ
    setupDatabase();

    const template = HtmlService.createTemplateFromFile('Index');
    return template.evaluate()
      .setTitle('ระบบติดตามงาน (Work Tracker)')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    return HtmlService.createHtmlOutput('<h3>เกิดข้อผิดพลาดในการโหลดระบบ: ' + error.message + '</h3>');
  }
}

/**
 * รวมไฟล์ HTML ย่อย (Include Helper)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * ให้บริการรับข้อมูล HTTP POST (Universal API Router & Upload File)
 * รองรับทั้งการอัปโหลดไฟล์ขนาดใหญ่ และการเรียกใช้งาน API จาก Frontend ภายนอก (GitHub Pages)
 * รับข้อมูลแบบ Content-Type: text/plain เพื่อข้าม CORS Preflight (OPTIONS)
 */
function doPost(e) {
  try {
    let payload = null;
    if (e && e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      payload = e.parameter;
    }

    if (!payload) {
      throw new Error('ไม่พบข้อมูลส่งมาใน Request');
    }

    // 1. Upload Attachment (เดิม)
    if (payload.action === 'uploadAttachment') {
      const uploadRes = uploadTaskAttachment(payload.taskId, payload.site, {
        fileName: payload.fileName,
        mimeType: payload.mimeType,
        size: payload.size,
        base64: payload.base64
      });

      return ContentService.createTextOutput(JSON.stringify(uploadRes))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 2. Universal API Router สำหรับเชื่อมต่อกับ Frontend (เช่น GitHub Pages)
    const PUBLIC_API_ACTIONS = {
      saveSettings: saveSettings,
      saveSitesList: saveSitesList,
      getInitialData: getInitialData,
      checkTasksUpdate: checkTasksUpdate,
      saveTask: saveTask,
      updateTaskStatus: updateTaskStatus,
      deleteTask: deleteTask,
      archiveCompletedTasks: archiveCompletedTasks,
      archiveSingleTask: archiveSingleTask,
      getArchivedTasks: getArchivedTasks,
      restoreArchivedTask: restoreArchivedTask,
      testTelegramConnection: testTelegramConnection,
      requestLoginPin: requestLoginPin,
      verifyLoginPin: verifyLoginPin,
      getUsersList: getUsersList,
      saveUser: saveUser,
      deleteUser: deleteUser,
      pingOnlineUser: pingOnlineUser,
      getOnlineUsers: getOnlineUsers,
      logoutOnlineUser: logoutOnlineUser,
      uploadTaskAttachment: uploadTaskAttachment,
      deleteTaskAttachment: deleteTaskAttachment
    };

    const targetFn = PUBLIC_API_ACTIONS[payload.action];
    if (typeof targetFn !== 'function') {
      throw new Error('Invalid or unsupported action: ' + payload.action);
    }

    const args = Array.isArray(payload.args) ? payload.args : [];
    const result = targetFn.apply(null, args);

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    console.error('doPost error: ' + err.message);
    return ContentService.createTextOutput(JSON.stringify({ success: false, isServerException: true, message: err.message, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * เชื่อมต่อไปยัง Google Spreadsheet ที่ใช้งานอยู่
 * รองรับทั้ง Container-bound (เชื่อมกับ Sheet ปัจจุบัน) หรือ Standalone (สร้างชีตใหม่/อ่านจาก ID)
 */
function getDb_() {
  // 1. ลองดึงจาก Active Spreadsheet ก่อน (กรณีเปิด Apps Script จากตัวชีต)
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;

  // 2. กรณี Standalone script: ค้นหาจาก Script Properties
  const props = PropertiesService.getScriptProperties();
  let sheetId = props.getProperty('SPREADSHEET_ID');

  if (sheetId) {
    try {
      return SpreadsheetApp.openById(sheetId);
    } catch (e) {
      console.warn('Cannot open sheet by ID, creating a new one: ' + e.message);
    }
  }

  // 3. หากยังไม่มี ให้สร้าง Google Sheet ใหม่ทันที
  ss = SpreadsheetApp.create('Work Tracker Database (โปรแกรมติดตามงาน)');
  props.setProperty('SPREADSHEET_ID', ss.getId());
  console.log('Created new database sheet with ID: ' + ss.getId());
  return ss;
}

/**
 * สร้างและจัดฟอร์แมตโครงสร้างฐานข้อมูล Tasks และ Settings
 */
function setupDatabase() {
  const ss = getDb_();

  // --- 1. ตาราง Tasks ---
  let taskSheet = ss.getSheetByName(SHEET_TASKS);
  if (!taskSheet) {
    taskSheet = ss.insertSheet(SHEET_TASKS);
    
    // ลบชีตเปล่าเริ่มต้น (เช่น Sheet1) ถ้ามี
    const sheets = ss.getSheets();
    for (let i = 0; i < sheets.length; i++) {
      if (sheets[i].getName() === 'Sheet1' && sheets.length > 1) {
        try { ss.deleteSheet(sheets[i]); } catch (e) {}
      }
    }

    const headers = [
      'Task ID',
      'Site',
      'Title',
      'Description',
      'Assignee',
      'Priority',
      'Status',
      'DueDate',
      'CreatedAt',
      'UpdatedAt',
      'Attachments',
      'FolderId',
      'Reporter',
      'ReportDate',
      'Checklist'
    ];

    // กำหนดหัวตาราง
    taskSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // ตกแต่งหัวตาราง (Dark Slate / White Text / Bold)
    const headerRange = taskSheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#1e293b')
               .setFontColor('#ffffff')
               .setFontWeight('bold')
               .setHorizontalAlignment('center')
               .setVerticalAlignment('middle');
    taskSheet.setRowHeight(1, 40);
    taskSheet.setFrozenRows(1);

    // เพิ่มตัวอย่างงานเบื้องต้นเพื่อความสะดวกในการเริ่มต้น
    const sampleDate = Utilities.formatDate(new Date(Date.now() + 86400000 * 3), 'Asia/Bangkok', 'yyyy-MM-dd');
    const today = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm');
    const todayDate = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');

    const sampleTasks = [
      [
        'TSK-1001',
        'ไซต์ A - อาคารสำนักงานใหญ่',
        'ตรวจเช็กระบบไฟฟ้าและตู้ MDB ชั้น 3',
        'ตรวจสอบสายไฟและระบบเบรกเกอร์ป้องกันไฟกระชาก',
        'สมชาย ช่างไฟ',
        'สูง',
        'กำลังดำเนินการ',
        sampleDate,
        today,
        today,
        '[]',
        '',
        'สมชาย ช่างไฟ',
        todayDate,
        JSON.stringify([
          { id: 'chk_sample_1', text: 'ตรวจเช็กความร้อนตู้ MDB', done: true },
          { id: 'chk_sample_2', text: 'วัดแรงดันไฟ 3 เฟส', done: false },
          { id: 'chk_sample_3', text: 'ส่งใบรายงานผลการตรวจ', done: false }
        ])
      ],
      [
        'TSK-1002',
        'ไซต์ B - โครงการคอนโดสุขุมวิท',
        'สั่งปูนผสมเสร็จสำหรับเทพื้นชั้น 5',
        'ติดต่อแพล้นท์ปูน CPAC จำนวน 40 คิว',
        'วิชัย โฟร์แมน',
        'ด่วนมาก',
        'รอดำเนินการ',
        sampleDate,
        today,
        today,
        '[]',
        '',
        'วิชัย โฟร์แมน',
        todayDate,
        '[]'
      ],
      [
        'TSK-1003',
        'ไซต์ C - โรงงานนิคมชลบุรี',
        'ส่งมอบงานติดตั้งโครงเหล็กหลังคา',
        'ทำเอกสารตรวจรับงานพร้อมแนบภาพถ่ายให้วิศวกรผู้ควบคุมงาน',
        'กิตติ วิศวกร',
        'ปานกลาง',
        'รอตรวจสอบ',
        sampleDate,
        today,
        today,
        '[]',
        '',
        'กิตติ วิศวกร',
        todayDate,
        '[]'
      ]
    ];

    taskSheet.getRange(2, 1, sampleTasks.length, headers.length).setValues(sampleTasks);
    
    // ปรับความกว้างคอลัมน์อัตโนมัติ
    for (let c = 1; c <= headers.length; c++) {
      taskSheet.autoResizeColumn(c);
    }
  } else {
    // หากตารางมีอยู่แล้วแต่ยังไม่มีคอลัมน์ Attachments, FolderId, Reporter, ReportDate หรือ Checklist ให้เพิ่มหัวตาราง
    const lastCol = taskSheet.getLastColumn();
    if (lastCol < 11) {
      taskSheet.getRange(1, 11).setValue('Attachments');
    }
    if (lastCol < 12) {
      taskSheet.getRange(1, 12).setValue('FolderId');
    }
    if (lastCol < 13) {
      taskSheet.getRange(1, 13).setValue('Reporter')
               .setBackground('#1e293b')
               .setFontColor('#ffffff')
               .setFontWeight('bold')
               .setHorizontalAlignment('center')
               .setVerticalAlignment('middle');
      taskSheet.autoResizeColumn(13);
    }
    if (lastCol < 14) {
      taskSheet.getRange(1, 14).setValue('ReportDate')
               .setBackground('#1e293b')
               .setFontColor('#ffffff')
               .setFontWeight('bold')
               .setHorizontalAlignment('center')
               .setVerticalAlignment('middle');
      taskSheet.autoResizeColumn(14);
    }
    if (lastCol < 15) {
      taskSheet.getRange(1, 15).setValue('Checklist')
               .setBackground('#1e293b')
               .setFontColor('#ffffff')
               .setFontWeight('bold')
               .setHorizontalAlignment('center')
               .setVerticalAlignment('middle');
      taskSheet.autoResizeColumn(15);
    }
  }

  // --- 2. ตาราง Settings ---
  let settingsSheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet(SHEET_SETTINGS);
    const setHeaders = ['Key', 'Value', 'Description'];
    settingsSheet.getRange(1, 1, 1, 3).setValues([setHeaders]);
    settingsSheet.getRange(1, 1, 1, 3).setBackground('#334155').setFontColor('#ffffff').setFontWeight('bold');
    settingsSheet.setRowHeight(1, 35);
    settingsSheet.setFrozenRows(1);

    const defaultSettings = [
      ['TELEGRAM_BOT_TOKEN', '', 'Token ของ Telegram Bot ที่ได้จาก @BotFather'],
      ['TELEGRAM_CHAT_ID', '', 'Chat ID หรือ Group ID หลัก สำหรับรับข้อความแจ้งเตือน (Default)'],
      ['NOTIFY_ON_CREATE', 'true', 'แจ้งเตือนเมื่อสร้างงานใหม่ (true/false)'],
      ['NOTIFY_ON_STATUS_CHANGE', 'true', 'แจ้งเตือนเมื่อเปลี่ยนสถานะงาน (true/false)'],
      ['NOTIFY_ASSIGNEE_ON_CREATE', 'true', 'แจ้งเตือนไปยังผู้รับผิดชอบโดยตรงเมื่อสร้างงานใหม่ (true/false)'],
      ['NOTIFY_MAIN_GROUP_ALWAYS', 'false', 'ส่งเข้ากลุ่มหลักส่วนกลางควบคู่ไปด้วยเสมอ (true/false)'],
      ['SITE_TELEGRAM_CHAT_IDS', '{}', 'Chat ID ของ Telegram แต่ละไซต์งาน (JSON Map)'],
      ['SITES', JSON.stringify(DEFAULT_SITES), 'รายชื่อไซต์งานทั้งหมดในระบบ (JSON Array)']
    ];

    settingsSheet.getRange(2, 1, defaultSettings.length, 3).setValues(defaultSettings);
    settingsSheet.autoResizeColumn(1);
    settingsSheet.autoResizeColumn(2);
    settingsSheet.autoResizeColumn(3);
  }

  // --- 3. ตาราง Users ---
  let userSheet = ss.getSheetByName(SHEET_USERS);
  if (!userSheet) {
    userSheet = ss.insertSheet(SHEET_USERS);
    const userHeaders = ['Username', 'FullName', 'TelegramChatId', 'Role', 'Status', 'LastLogin', 'AllowedSites'];
    userSheet.getRange(1, 1, 1, userHeaders.length).setValues([userHeaders]);
    userSheet.getRange(1, 1, 1, userHeaders.length)
             .setBackground('#0f172a')
             .setFontColor('#ffffff')
             .setFontWeight('bold')
             .setHorizontalAlignment('center')
             .setVerticalAlignment('middle');
    userSheet.setRowHeight(1, 35);
    userSheet.setFrozenRows(1);

    const defaultUsers = [
      ['admin', 'ผู้ดูแลระบบ', '', 'Admin', 'Active', '', 'ALL'],
      ['somchai', 'สมชาย ช่างไฟ', '', 'Member', 'Active', '', 'ALL'],
      ['wichai', 'วิชัย โฟร์แมน', '', 'Member', 'Active', '', 'ALL']
    ];
    userSheet.getRange(2, 1, defaultUsers.length, userHeaders.length).setValues(defaultUsers);
    for (let c = 1; c <= userHeaders.length; c++) {
      userSheet.autoResizeColumn(c);
    }
  } else {
    // หากตารางมีอยู่แล้วแต่ยังไม่มีคอลัมน์ AllowedSites ให้เพิ่มหัวตารางคอลัมน์ที่ 7
    const lastCol = userSheet.getLastColumn();
    if (lastCol < 7) {
      userSheet.getRange(1, 7).setValue('AllowedSites')
               .setBackground('#0f172a')
               .setFontColor('#ffffff')
               .setFontWeight('bold')
               .setHorizontalAlignment('center')
               .setVerticalAlignment('middle');
      userSheet.autoResizeColumn(7);
    }
  }

  return { success: true, sheetUrl: ss.getUrl() };
}

/**
 * ดึงการตั้งค่าทั้งหมดจากชีต Settings
 */
function getSettings_() {
  const ss = getDb_();
  let sheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!sheet) {
    setupDatabase();
    sheet = ss.getSheetByName(SHEET_SETTINGS);
  }

  const values = sheet.getDataRange().getValues();
  const settings = {
    TELEGRAM_BOT_TOKEN: '',
    TELEGRAM_CHAT_ID: '',
    NOTIFY_ON_CREATE: true,
    NOTIFY_ON_STATUS_CHANGE: true,
    NOTIFY_ASSIGNEE_ON_CREATE: true,
    NOTIFY_MAIN_GROUP_ALWAYS: false,
    SITE_TELEGRAM_CHAT_IDS: {},
    SITES: DEFAULT_SITES
  };

  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][0]).trim();
    const val = String(values[i][1]).trim();
    if (key === 'TELEGRAM_BOT_TOKEN') settings.TELEGRAM_BOT_TOKEN = val;
    if (key === 'TELEGRAM_CHAT_ID') settings.TELEGRAM_CHAT_ID = val;
    if (key === 'NOTIFY_ON_CREATE') settings.NOTIFY_ON_CREATE = (val.toLowerCase() === 'true');
    if (key === 'NOTIFY_ON_STATUS_CHANGE') settings.NOTIFY_ON_STATUS_CHANGE = (val.toLowerCase() === 'true');
    if (key === 'NOTIFY_ASSIGNEE_ON_CREATE') settings.NOTIFY_ASSIGNEE_ON_CREATE = (val.toLowerCase() === 'true');
    if (key === 'NOTIFY_MAIN_GROUP_ALWAYS') settings.NOTIFY_MAIN_GROUP_ALWAYS = (val.toLowerCase() === 'true');
    if (key === 'SITE_TELEGRAM_CHAT_IDS') {
      try {
        const parsed = JSON.parse(val);
        if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) {
          settings.SITE_TELEGRAM_CHAT_IDS = parsed;
        }
      } catch (e) {
        settings.SITE_TELEGRAM_CHAT_IDS = {};
      }
    }
    if (key === 'SITES') {
      try {
        const parsed = JSON.parse(val);
        if (Array.isArray(parsed) && parsed.length > 0) {
          settings.SITES = parsed;
        }
      } catch (e) {
        if (val) settings.SITES = val.split(',').map(s => s.trim()).filter(Boolean);
      }
    }
  }

  return settings;
}

/**
 * บันทึกการตั้งค่ากลับลงในชีต Settings
 */
function saveSettings(newSettings) {
  try {
    const ss = getDb_();
    let sheet = ss.getSheetByName(SHEET_SETTINGS);
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName(SHEET_SETTINGS);
    }

    const values = sheet.getDataRange().getValues();
    const keysMap = {};
    for (let i = 1; i < values.length; i++) {
      keysMap[values[i][0]] = i + 1; // row index in Sheet (1-based)
    }

    const updateOrAdd = (key, value, desc) => {
      if (keysMap[key]) {
        sheet.getRange(keysMap[key], 2).setValue(value);
      } else {
        sheet.appendRow([key, value, desc || '']);
      }
    };

    if (newSettings.TELEGRAM_BOT_TOKEN !== undefined) {
      updateOrAdd('TELEGRAM_BOT_TOKEN', newSettings.TELEGRAM_BOT_TOKEN, 'Token Telegram Bot');
    }
    if (newSettings.TELEGRAM_CHAT_ID !== undefined) {
      updateOrAdd('TELEGRAM_CHAT_ID', newSettings.TELEGRAM_CHAT_ID, 'Chat ID หลัก รับการแจ้งเตือน');
    }
    if (newSettings.NOTIFY_ON_CREATE !== undefined) {
      updateOrAdd('NOTIFY_ON_CREATE', String(newSettings.NOTIFY_ON_CREATE), 'แจ้งเตือนเมื่องานถูกสร้าง');
    }
    if (newSettings.NOTIFY_ON_STATUS_CHANGE !== undefined) {
      updateOrAdd('NOTIFY_ON_STATUS_CHANGE', String(newSettings.NOTIFY_ON_STATUS_CHANGE), 'แจ้งเตือนเมื่อเปลี่ยนสถานะ');
    }
    if (newSettings.NOTIFY_ASSIGNEE_ON_CREATE !== undefined) {
      updateOrAdd('NOTIFY_ASSIGNEE_ON_CREATE', String(newSettings.NOTIFY_ASSIGNEE_ON_CREATE), 'แจ้งเตือนผู้รับผิดชอบเมื่อมีงานใหม่');
    }
    if (newSettings.NOTIFY_MAIN_GROUP_ALWAYS !== undefined) {
      updateOrAdd('NOTIFY_MAIN_GROUP_ALWAYS', String(newSettings.NOTIFY_MAIN_GROUP_ALWAYS), 'ส่งเข้ากลุ่มหลักส่วนกลางควบคู่ไปด้วยเสมอ');
    }
    if (newSettings.SITE_TELEGRAM_CHAT_IDS !== undefined && typeof newSettings.SITE_TELEGRAM_CHAT_IDS === 'object') {
      updateOrAdd('SITE_TELEGRAM_CHAT_IDS', JSON.stringify(newSettings.SITE_TELEGRAM_CHAT_IDS), 'Chat ID กลุ่ม Telegram แต่ละไซต์งาน (JSON Object)');
    }
    if (newSettings.SITES !== undefined && Array.isArray(newSettings.SITES)) {
      updateOrAdd('SITES', JSON.stringify(newSettings.SITES), 'รายชื่อไซต์งานทั้งหมด');
    }

    return { success: true, message: 'บันทึกการตั้งค่าสำเร็จ' };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * บันทึกรายชื่อไซต์งานพร้อมการจับคู่ Chat ID ของ Telegram แต่ละไซต์ลงในชีต Settings
 * @param {Array<string>} sitesArray รายชื่อไซต์งาน
 * @param {Object} [siteChatIds] การจับคู่ { [siteName]: chatId }
 */
function saveSitesList(sitesArray, siteChatIds) {
  try {
    if (!Array.isArray(sitesArray)) {
      return { success: false, message: 'ข้อมูลไซต์งานต้องเป็นรายการ (Array)' };
    }

    const cleanSites = sitesArray.map(s => String(s).trim()).filter(Boolean);
    const ss = getDb_();
    let sheet = ss.getSheetByName(SHEET_SETTINGS);
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName(SHEET_SETTINGS);
    }

    const values = sheet.getDataRange().getValues();
    const keysMap = {};
    for (let i = 1; i < values.length; i++) {
      keysMap[values[i][0]] = i + 1;
    }

    const updateOrAdd = (key, value, desc) => {
      if (keysMap[key]) {
        sheet.getRange(keysMap[key], 2).setValue(value);
      } else {
        sheet.appendRow([key, value, desc || '']);
      }
    };

    updateOrAdd('SITES', JSON.stringify(cleanSites), 'รายชื่อไซต์งานทั้งหมด');

    let cleanChatIds = {};
    if (siteChatIds && typeof siteChatIds === 'object') {
      cleanSites.forEach(s => {
        if (siteChatIds[s] && String(siteChatIds[s]).trim()) {
          cleanChatIds[s] = String(siteChatIds[s]).trim();
        }
      });
      updateOrAdd('SITE_TELEGRAM_CHAT_IDS', JSON.stringify(cleanChatIds), 'Chat ID กลุ่ม Telegram แต่ละไซต์งาน (JSON Object)');
    }

    return { 
      success: true, 
      message: 'บันทึกรายชื่อไซต์งานสำเร็จ', 
      sites: cleanSites,
      siteTelegramChatIds: cleanChatIds 
    };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * ดึงข้อมูลเริ่มต้นทั้งหมด (งาน, สถิติ, ไซต์งาน, การตั้งค่า)
 * รองรับการกรองข้อมูลตามสิทธิ์การเข้าถึงไซต์งานของผู้ใช้ (requestUsername)
 */
function getInitialData(requestUsername) {
  try {
    const ss = getDb_();
    let sheet = ss.getSheetByName(SHEET_TASKS);
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName(SHEET_TASKS);
    }

    // ดึงรายชื่อผู้ใช้จากชีต Users และตรวจสอบสิทธิ์ของผู้ใช้ปัจจุบัน
    let userSheet = ss.getSheetByName(SHEET_USERS);
    if (!userSheet) {
      setupDatabase();
      userSheet = ss.getSheetByName(SHEET_USERS);
    }

    let currentUser = null;
    let allowedSites = 'ALL';
    let userRole = 'Member';

    const uData = userSheet.getDataRange().getValues();
    const users = [];

    for (let i = 1; i < uData.length; i++) {
      if (!uData[i][0]) continue;
      const uName = String(uData[i][0]).trim();
      const status = String(uData[i][4] || 'Active').trim();
      const role = String(uData[i][3] || 'Member').trim();

      // อ่าน AllowedSites (คอลัมน์ 7)
      let sitesPerm = 'ALL';
      if (uData[i][6] !== undefined && String(uData[i][6]).trim() !== '') {
        const raw = String(uData[i][6]).trim();
        if (raw !== 'ALL') {
          try {
            const parsed = JSON.parse(raw);
            if (Array.isArray(parsed)) sitesPerm = parsed;
          } catch (e) {
            sitesPerm = 'ALL';
          }
        }
      }

      if (status.toLowerCase() !== 'inactive') {
        users.push({
          username: uName,
          fullName: String(uData[i][1] || '').trim() || uName,
          role: role,
          status: status
        });
      }

      // ตรวจสอบผู้ใช้ที่ร้องขอข้อมูล
      if (requestUsername && uName.toLowerCase() === String(requestUsername).trim().toLowerCase()) {
        currentUser = {
          username: uName,
          fullName: String(uData[i][1] || '').trim() || uName,
          role: role,
          allowedSites: sitesPerm
        };
        userRole = role;
        allowedSites = (role === 'Admin') ? 'ALL' : sitesPerm;
      }
    }

    const settings = getSettings_();
    const siteSet = new Set(settings.SITES);

    const data = sheet.getDataRange().getValues();
    const tasks = [];

    // แถวที่ 1 เป็น Headers
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue; // ข้ามแถวว่าง

      let dueDateFormatted = '';
      if (row[7]) {
        dueDateFormatted = normalizeDateToIso_(row[7]);
      }

      const taskSite = String(row[1] || 'ไม่ระบุไซต์งาน').trim();
      if (taskSite) siteSet.add(taskSite);

      let attachments = [];
      if (row[10]) {
        try {
          const parsed = JSON.parse(row[10]);
          if (Array.isArray(parsed)) attachments = parsed;
        } catch (e) {}
      }

      let reportDateFormatted = '';
      if (row[13]) {
        reportDateFormatted = normalizeDateToIso_(row[13]);
      } else if (row[8]) {
        reportDateFormatted = normalizeDateToIso_(row[8]);
      }

      let checklist = [];
      if (row[14]) {
        try {
          const parsed = JSON.parse(row[14]);
          if (Array.isArray(parsed)) checklist = parsed;
        } catch (e) {}
      }

      tasks.push({
        id: String(row[0]),
        site: taskSite,
        title: String(row[2] || ''),
        description: String(row[3] || ''),
        assignee: String(row[4] || 'ไม่ระบุ'),
        priority: String(row[5] || 'ทั่วไป'),
        status: String(row[6] || 'รอดำเนินการ'),
        dueDate: dueDateFormatted,
        createdAt: row[8] ? String(row[8]) : '',
        updatedAt: row[9] ? String(row[9]) : '',
        attachments: attachments,
        folderId: row[11] ? String(row[11]) : '',
        reporter: String(row[12] || ''),
        reportDate: reportDateFormatted,
        checklist: checklist
      });
    }

    const allSites = Array.from(siteSet);

    // กรองงานและไซต์งานตามสิทธิ์ของผู้ใช้
    let filteredTasks = tasks;
    let filteredSites = allSites;

    if (userRole !== 'Admin' && Array.isArray(allowedSites)) {
      const allowedSet = new Set(allowedSites);
      filteredTasks = tasks.filter(t => allowedSet.has(t.site));
      filteredSites = allSites.filter(s => allowedSet.has(s));
      // ให้แน่ใจว่าไซต์ที่ได้รับอนุญาตทั้งหมดมีอยู่ในรายการตัวเลือก
      allowedSites.forEach(s => {
        if (!filteredSites.includes(s)) filteredSites.push(s);
      });
    }

    // คำนวณสรุปสถิติ (Stats) จากรายการงานที่ผู้ใช้มีสิทธิ์เข้าถึง
    const todayStr = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
    const stats = {
      total: filteredTasks.length,
      todo: filteredTasks.filter(t => t.status === 'รอดำเนินการ').length,
      inProgress: filteredTasks.filter(t => t.status === 'กำลังดำเนินการ').length,
      review: filteredTasks.filter(t => t.status === 'รอตรวจสอบ').length,
      done: filteredTasks.filter(t => t.status === 'เสร็จสิ้น').length,
      overdue: filteredTasks.filter(t => t.dueDate && t.dueDate < todayStr && t.status !== 'เสร็จสิ้น').length
    };

    // ดึงสถานะผู้ใช้ออนไลน์เบื้องต้น
    const onlineInfo = getOnlineUsers();

    return {
      success: true,
      tasks: filteredTasks,
      sites: filteredSites,
      allSites: allSites, // รายชื่อไซต์ทั้งหมดในระบบ (สำหรับหน้าจอตั้งค่า)
      users: users,
      currentUserProfile: currentUser,
      statuses: TASK_STATUSES,
      priorities: TASK_PRIORITIES,
      settings: {
        hasTelegramBot: !!settings.TELEGRAM_BOT_TOKEN,
        telegramChatId: settings.TELEGRAM_CHAT_ID,
        notifyOnCreate: settings.NOTIFY_ON_CREATE,
        notifyOnStatusChange: settings.NOTIFY_ON_STATUS_CHANGE,
        notifyAssigneeOnCreate: settings.NOTIFY_ASSIGNEE_ON_CREATE !== false,
        notifyMainGroupAlways: settings.NOTIFY_MAIN_GROUP_ALWAYS === true,
        siteTelegramChatIds: settings.SITE_TELEGRAM_CHAT_IDS || {},
        sites: settings.SITES
      },
      stats: stats,
      version: getTasksDataVersion_(),
      onlineUsers: onlineInfo.users || [],
      onlineCount: onlineInfo.count || 0,
      sheetUrl: ss.getUrl(),
      webAppUrl: (function() {
        try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; }
      })()
    };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * =========================================================================
 * REAL-TIME TASK SYNC SYSTEM (ระบบซิงค์สถานะงานแบบ Real-time)
 * =========================================================================
 */

/**
 * ดึงเวอร์ชันข้อมูลปัจจุบันจาก CacheService / PropertiesService
 * @returns {string} Timestamp string
 */
function getTasksDataVersion_() {
  const cache = CacheService.getScriptCache();
  let v = cache.get('TASKS_DATA_VERSION');
  if (!v) {
    v = PropertiesService.getScriptProperties().getProperty('TASKS_DATA_VERSION');
    if (!v) {
      v = String(Date.now());
      PropertiesService.getScriptProperties().setProperty('TASKS_DATA_VERSION', v);
    }
    cache.put('TASKS_DATA_VERSION', v, 21600); // 6 ชั่วโมง
  }
  return v;
}

/**
 * บันทึกการเปลี่ยนแปลงข้อมูลเพื่อแจ้งเตือนทุกไคลเอนต์แบบ Real-time
 * @param {Object} [meta] ข้อมูลเหตุการณ์การเปลี่ยนแปลง เช่น { type: 'STATUS_CHANGE', taskId, oldStatus, newStatus }
 */
function notifyDataChanged_(meta) {
  const timestamp = String(Date.now());
  const cache = CacheService.getScriptCache();
  cache.put('TASKS_DATA_VERSION', timestamp, 21600);
  PropertiesService.getScriptProperties().setProperty('TASKS_DATA_VERSION', timestamp);

  if (meta) {
    try {
      cache.put('TASKS_LAST_CHANGE_META', JSON.stringify(meta), 300); // 5 นาที
    } catch (e) {}
  }
}

/**
 * ตรวจสอบอัปเดตแบบ Real-time โดยไม่แตะต้อง Google Sheets หากไม่มีการเปลี่ยนแปลง (Zero-Quota)
 * @param {string} clientVersion
 * @param {string} username
 * @returns {Object}
 */
function checkTasksUpdate(clientVersion, username) {
  try {
    const serverVersion = getTasksDataVersion_();
    if (clientVersion && String(clientVersion) === String(serverVersion)) {
      return { hasUpdate: false, version: serverVersion };
    }

    const cache = CacheService.getScriptCache();
    const metaRaw = cache.get('TASKS_LAST_CHANGE_META');
    let meta = null;
    if (metaRaw) {
      try { meta = JSON.parse(metaRaw); } catch (e) {}
    }

    const initialData = getInitialData(username);
    return {
      hasUpdate: true,
      version: serverVersion,
      meta: meta,
      data: initialData
    };
  } catch (error) {
    return { hasUpdate: false, error: error.message };
  }
}

/**
 * Simple Trigger เมื่อมีการแก้ไขข้อมูลใน Google Sheets โดยตรง
 */
function onEdit(e) {
  try {
    notifyDataChanged_({ type: 'SHEET_EDIT', desc: 'แก้ไขข้อมูลใน Google Sheets โดยตรง' });
  } catch (err) {}
}

/**
 * ปรับรูปแบบวันที่ให้เป็นมาตรฐาน ISO YYYY-MM-DD (ค.ศ.) เสมอ
 * รองรับทั้ง DD/MM/YYYY, DD/MM/BBBB (พ.ศ.), YYYY-MM-DD
 * @param {string} dateStr 
 * @returns {string} YYYY-MM-DD หรือ ''
 */
function normalizeDateToIso_(dateStr) {
  if (!dateStr) return '';
  if (dateStr instanceof Date) {
    if (isNaN(dateStr.getTime())) return '';
    return Utilities.formatDate(dateStr, 'Asia/Bangkok', 'yyyy-MM-dd');
  }
  const s = String(dateStr).trim();
  if (!s) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  if (/^\d{4}-\d{2}-\d{2}[T\s]/.test(s)) return s.substring(0, 10);
  const match = s.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})$/);
  if (match) {
    let day = parseInt(match[1], 10);
    let month = parseInt(match[2], 10);
    let year = parseInt(match[3], 10);
    if (year < 100) {
      year = year >= 50 ? (year + 2500) - 543 : 2000 + year;
    } else if (year > 2400) {
      year = year - 543;
    }
    const mm = String(month).padStart(2, '0');
    const dd = String(day).padStart(2, '0');
    return `${year}-${mm}-${dd}`;
  }
  return s;
}

/**
 * บันทึกงาน (เพิ่มงานใหม่ หรือแก้ไขงานเดิม)
 */
function saveTask(taskData) {
  try {
    const ss = getDb_();
    const sheet = ss.getSheetByName(SHEET_TASKS);
    if (!sheet) throw new Error('ไม่พบแผ่นงาน Tasks');

    const nowStr = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm');
    const values = sheet.getDataRange().getValues();
    
    let rowIndex = -1;
    let isNew = false;
    let oldTask = null;

    // ตรวจสอบความถูกต้องของวันที่
    const normalizedDueDate = normalizeDateToIso_(taskData.dueDate);
    const normalizedReportDate = normalizeDateToIso_(taskData.reportDate) || Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');

    // ตรวจสอบว่าเป็นการแก้ไขงานเดิมหรือไม่
    if (taskData.id && !taskData.id.startsWith('NEW-')) {
      for (let i = 1; i < values.length; i++) {
        if (String(values[i][0]) === String(taskData.id)) {
          rowIndex = i + 1;
          oldTask = {
            id: values[i][0],
            site: values[i][1],
            title: values[i][2],
            description: values[i][3],
            assignee: values[i][4],
            priority: values[i][5],
            status: values[i][6],
            dueDate: values[i][7]
          };
          break;
        }
      }
    }

    let finalTaskId = taskData.id;

    if (rowIndex === -1) {
      // สร้างงานใหม่: สร้างรหัส TSK-XXXX
      isNew = true;
      let maxNum = 1000;
      for (let i = 1; i < values.length; i++) {
        const match = String(values[i][0]).match(/^TSK-(\d+)$/);
        if (match) {
          const num = parseInt(match[1], 10);
          if (num > maxNum) maxNum = num;
        }
      }
      finalTaskId = 'TSK-' + (maxNum + 1);

      const checklistJson = JSON.stringify(Array.isArray(taskData.checklist) ? taskData.checklist : []);

      const newRow = [
        finalTaskId,
        taskData.site || 'ไม่ระบุไซต์งาน',
        taskData.title || 'ไม่มีชื่อเรื่อง',
        taskData.description || '',
        taskData.assignee || 'ไม่ระบุ',
        taskData.priority || 'ทั่วไป',
        taskData.status || 'รอดำเนินการ',
        normalizedDueDate,
        nowStr,
        nowStr,
        '[]',
        '',
        taskData.reporter || '',
        normalizedReportDate,
        checklistJson
      ];

      sheet.appendRow(newRow);
      SpreadsheetApp.flush();
    } else {
      // อัปเดตงานเดิม
      sheet.getRange(rowIndex, 2, 1, 7).setValues([[
        taskData.site || 'ไม่ระบุไซต์งาน',
        taskData.title || 'ไม่มีชื่อเรื่อง',
        taskData.description || '',
        taskData.assignee || 'ไม่ระบุ',
        taskData.priority || 'ทั่วไป',
        taskData.status || 'รอดำเนินการ',
        normalizedDueDate
      ]]);
      sheet.getRange(rowIndex, 10).setValue(nowStr); // UpdatedAt
      if (taskData.reporter !== undefined) {
        sheet.getRange(rowIndex, 13).setValue(taskData.reporter || '');
      }
      if (taskData.reportDate !== undefined) {
        sheet.getRange(rowIndex, 14).setValue(normalizedReportDate);
      }
      if (taskData.checklist !== undefined) {
        sheet.getRange(rowIndex, 15).setValue(JSON.stringify(Array.isArray(taskData.checklist) ? taskData.checklist : []));
      }
      SpreadsheetApp.flush();

      // หากไซต์งานเปลี่ยน ให้ย้ายโฟลเดอร์ใน Drive ไปยังโฟลเดอร์ไซต์ใหม่ด้วย (ถ้ามี)
      if (oldTask && oldTask.site !== taskData.site) {
        const folderId = values[rowIndex - 1][11];
        if (folderId) {
          try {
            relocateTaskFolder_(folderId, taskData.site);
          } catch (e) {
            console.warn('relocateTaskFolder_ warning: ' + e.message);
          }
        }
      }
    }

    const currentAttachmentsRaw = (rowIndex !== -1 && values[rowIndex - 1][10]) ? values[rowIndex - 1][10] : '[]';
    let currentAttachments = [];
    try {
      currentAttachments = JSON.parse(currentAttachmentsRaw) || [];
    } catch (e) {}

    const currentFolderId = (rowIndex !== -1 && values[rowIndex - 1][11]) ? String(values[rowIndex - 1][11]) : '';

    const savedTask = {
      id: finalTaskId,
      site: taskData.site,
      title: taskData.title,
      description: taskData.description,
      assignee: taskData.assignee,
      reporter: taskData.reporter || '',
      reportDate: normalizedReportDate,
      priority: taskData.priority,
      status: taskData.status,
      dueDate: normalizedDueDate,
      updatedAt: nowStr,
      attachments: currentAttachments,
      folderId: currentFolderId,
      checklist: Array.isArray(taskData.checklist) ? taskData.checklist : []
    };

    // ส่งการแจ้งเตือน Telegram
    try {
      const settings = getSettings_();
      if (settings.TELEGRAM_BOT_TOKEN) {
        // 1. ส่งแจ้งเตือนเข้าห้องแชตกลุ่ม (แยกตามไซต์งาน หรือกลุ่มหลัก)
        if (isNew) {
          notifyTelegramGroupForTask_('CREATE', savedTask);
        } else if (oldTask && oldTask.status !== savedTask.status) {
          notifyTelegramGroupForTask_('STATUS_CHANGE', savedTask, oldTask.status);
        }

        // 2. ส่งแจ้งเตือนตรงไปยัง Telegram Chat ID ส่วนตัวของผู้รับผิดชอบเมื่องานถูกสร้างใหม่
        if (isNew && settings.NOTIFY_ASSIGNEE_ON_CREATE) {
          try {
            const personalChatId = getAssigneeTelegramChatId_(savedTask.assignee);
            if (personalChatId) {
              const personalMsg = buildPersonalAssigneeMessage_('CREATE', savedTask);
              sendTelegramRaw_(settings.TELEGRAM_BOT_TOKEN, personalChatId, personalMsg);
            }
          } catch (pErr) {
            console.error('Telegram personal assignee notification error: ' + pErr.message);
          }
        }
      }
    } catch (notifyErr) {
      console.error('Telegram notification error: ' + notifyErr.message);
    }

    notifyDataChanged_({
      type: 'TASK_SAVED',
      taskId: finalTaskId,
      isNew: isNew,
      title: savedTask.title,
      status: savedTask.status
    });

    return { success: true, isNew: isNew, task: savedTask };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * อัปเดตเฉพาะสถานะงาน (ใช้สำหรับลากเปลี่ยนสถานะใน Kanban หรือคลิกเปลี่ยนสถานะด่วน)
 */
function updateTaskStatus(taskId, newStatus) {
  try {
    const ss = getDb_();
    const sheet = ss.getSheetByName(SHEET_TASKS);
    if (!sheet) throw new Error('ไม่พบแผ่นงาน Tasks');

    const values = sheet.getDataRange().getValues();
    let rowIndex = -1;
    let currentTask = null;

    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === String(taskId)) {
        rowIndex = i + 1;
        currentTask = {
          id: values[i][0],
          site: values[i][1],
          title: values[i][2],
          description: values[i][3],
          assignee: values[i][4],
          priority: values[i][5],
          oldStatus: values[i][6],
          status: newStatus,
          dueDate: values[i][7],
          reporter: String(values[i][12] || ''),
          reportDate: String(values[i][13] || '')
        };
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: 'ไม่พบรหัสงาน ' + taskId };
    }

    const nowStr = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm');
    sheet.getRange(rowIndex, 7).setValue(newStatus); // Status (Col 7)
    sheet.getRange(rowIndex, 10).setValue(nowStr);   // UpdatedAt (Col 10)

    // ส่งแจ้งเตือน Telegram เข้ากลุ่ม (แยกตามไซต์งาน หรือกลุ่มหลัก)
    try {
      if (currentTask.oldStatus !== newStatus) {
        notifyTelegramGroupForTask_('STATUS_CHANGE', currentTask, currentTask.oldStatus);
      }
    } catch (e) {
      console.error('Telegram error: ' + e.message);
    }

    notifyDataChanged_({
      type: 'STATUS_CHANGE',
      taskId: taskId,
      oldStatus: currentTask.oldStatus,
      newStatus: newStatus,
      title: currentTask.title
    });

    return { success: true, taskId: taskId, newStatus: newStatus };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * ลบรายการงาน พร้อมลบโฟลเดอร์และไฟล์แนบที่เกี่ยวข้องใน Google Drive (Cascade Delete)
 */
function deleteTask(taskId) {
  try {
    const ss = getDb_();
    const sheet = ss.getSheetByName(SHEET_TASKS);
    if (!sheet) throw new Error('ไม่พบแผ่นงาน Tasks');

    const values = sheet.getDataRange().getValues();
    let rowIndex = -1;
    let targetRow = null;

    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === String(taskId)) {
        rowIndex = i + 1;
        targetRow = values[i];
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: 'ไม่พบรหัสงานที่จะลบ' };
    }

    // ลบโฟลเดอร์ของงานและไฟล์แนบทั้งหมดใน Google Drive (Cascade Delete)
    try {
      const folderId = targetRow[11] ? String(targetRow[11]).trim() : '';
      const site = targetRow[1] ? String(targetRow[1]).trim() : '';
      deleteTaskFolderAndFiles_(folderId, site, taskId);
    } catch (driveErr) {
      console.warn('Error deleting task folder in Drive: ' + driveErr.message);
    }

    sheet.deleteRow(rowIndex);

    notifyDataChanged_({
      type: 'TASK_DELETED',
      taskId: taskId
    });

    return { success: true, taskId: taskId };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * =========================================================================
 * TASKS ARCHIVE SYSTEM (ระบบคลังจัดเก็บงานถาวร)
 * =========================================================================
 */

/**
 * ดึงหรือสร้างแผ่นงาน Tasks_Archive (คลังถาวร)
 */
function getArchiveSheet_(ss) {
  if (!ss) ss = getDb_();
  let sheet = ss.getSheetByName(SHEET_ARCHIVE);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_ARCHIVE);
    const headers = [
      'Task ID',
      'Site',
      'Title',
      'Description',
      'Assignee',
      'Priority',
      'Status',
      'DueDate',
      'CreatedAt',
      'UpdatedAt',
      'Attachments',
      'FolderId',
      'Reporter',
      'ReportDate',
      'Checklist',
      'ArchivedAt'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
         .setBackground('#334155')
         .setFontColor('#ffffff')
         .setFontWeight('bold')
         .setHorizontalAlignment('center')
         .setVerticalAlignment('middle');
    sheet.setRowHeight(1, 40);
    sheet.setFrozenRows(1);
    for (let c = 1; c <= headers.length; c++) {
      sheet.autoResizeColumn(c);
    }
  } else {
    // หากตาราง Archive เดิมยังไม่มีคอลัมน์ Reporter, ReportDate หรือ Checklist ให้เพิ่ม
    let lastCol = sheet.getLastColumn();
    let currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    if (!currentHeaders.includes('Reporter')) {
      const archivedAtCol = currentHeaders.indexOf('ArchivedAt') + 1;
      const insertAt = archivedAtCol > 0 ? archivedAtCol : lastCol;
      sheet.insertColumnBefore(insertAt);
      sheet.getRange(1, insertAt).setValue('Reporter')
           .setBackground('#334155')
           .setFontColor('#ffffff')
           .setFontWeight('bold')
           .setHorizontalAlignment('center')
           .setVerticalAlignment('middle');
      sheet.autoResizeColumn(insertAt);
    }
    lastCol = sheet.getLastColumn();
    currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    if (!currentHeaders.includes('ReportDate')) {
      const archivedAtCol = currentHeaders.indexOf('ArchivedAt') + 1;
      const insertAt = archivedAtCol > 0 ? archivedAtCol : lastCol;
      sheet.insertColumnBefore(insertAt);
      sheet.getRange(1, insertAt).setValue('ReportDate')
           .setBackground('#334155')
           .setFontColor('#ffffff')
           .setFontWeight('bold')
           .setHorizontalAlignment('center')
           .setVerticalAlignment('middle');
      sheet.autoResizeColumn(insertAt);
    }
    lastCol = sheet.getLastColumn();
    currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    if (!currentHeaders.includes('Checklist')) {
      const archivedAtCol = currentHeaders.indexOf('ArchivedAt') + 1;
      const insertAt = archivedAtCol > 0 ? archivedAtCol : lastCol;
      sheet.insertColumnBefore(insertAt);
      sheet.getRange(1, insertAt).setValue('Checklist')
           .setBackground('#334155')
           .setFontColor('#ffffff')
           .setFontWeight('bold')
           .setHorizontalAlignment('center')
           .setVerticalAlignment('middle');
      sheet.autoResizeColumn(insertAt);
    }
  }
  return sheet;
}

/**
 * จัดการย้ายงานที่สถานะ "เสร็จสิ้น" ไปยังแผ่นงาน Tasks_Archive (คลังถาวร)
 * @param {number|string} daysThreshold - 0 หรือ 'ALL' หมายถึงทั้งหมด, หรือตัวเลขจำนวนวัน เช่น 7 หรือ 30 วัน
 * @returns {Object} { success: boolean, archivedCount: number, message: string }
 */
function archiveCompletedTasks(daysThreshold) {
  try {
    const ss = getDb_();
    const taskSheet = ss.getSheetByName(SHEET_TASKS);
    if (!taskSheet) throw new Error('ไม่พบแผ่นงาน Tasks');

    const archiveSheet = getArchiveSheet_(ss);
    const data = taskSheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { success: true, archivedCount: 0, message: 'ไม่มีงานในระบบ' };
    }

    const now = new Date();
    const thresholdDays = (daysThreshold === 'ALL' || daysThreshold === undefined || daysThreshold === null) ? 0 : parseInt(daysThreshold, 10);
    const cutoffTime = thresholdDays > 0 ? (now.getTime() - (thresholdDays * 86400000)) : now.getTime() + 86400000;

    const rowsToArchive = [];
    const rowIndicesToDelete = [];
    const archivedAtStr = Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');

    // วนลูปจากแถวล่างขึ้นบน เพื่อให้ลบแถวได้ถูกต้องโดยไม่ต้องคำนวณ offset ใหม่
    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      const status = String(row[6] || '').trim();

      if (status === 'เสร็จสิ้น') {
        let shouldArchive = false;
        if (thresholdDays <= 0) {
          shouldArchive = true;
        } else {
          // ตรวจสอบจาก UpdatedAt (row[9]) หรือ DueDate (row[7]) หรือ CreatedAt (row[8])
          const dateVal = row[9] || row[7] || row[8];
          if (dateVal) {
            const parsedTime = new Date(dateVal).getTime();
            if (!isNaN(parsedTime) && parsedTime <= cutoffTime) {
              shouldArchive = true;
            }
          } else {
            shouldArchive = true;
          }
        }

        if (shouldArchive) {
          // แถวที่จะย้ายไป Archive: 15 คอลัมน์เดิม + ArchivedAt = 16 คอลัมน์
          const archiveRow = [
            row[0], row[1], row[2], row[3], row[4],
            row[5], row[6], row[7], row[8], row[9],
            row[10], row[11], row[12] || '', row[13] || '',
            row[14] || '[]', archivedAtStr
          ];
          rowsToArchive.push(archiveRow);
          rowIndicesToDelete.push(i + 1); // 1-based index
        }
      }
    }

    if (rowsToArchive.length === 0) {
      return { success: true, archivedCount: 0, message: 'ไม่พบงานที่เข้าเงื่อนไขการจัดเก็บ' };
    }

    // เขียนลงในแผ่นงาน Tasks_Archive (กลับลำดับเพื่อให้เรียงจากเก่าไปใหม่ตามปกติ)
    rowsToArchive.reverse();
    archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, rowsToArchive.length, 16).setValues(rowsToArchive);

    // ลบแถวออกจาก Tasks
    for (let r = 0; r < rowIndicesToDelete.length; r++) {
      taskSheet.deleteRow(rowIndicesToDelete[r]);
    }

    SpreadsheetApp.flush();
    notifyDataChanged_({
      type: 'TASKS_ARCHIVED',
      count: rowsToArchive.length
    });
    return { success: true, archivedCount: rowsToArchive.length };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * จัดการย้ายงานเดี่ยวเข้าคลังถาวร
 * @param {string} taskId
 */
function archiveSingleTask(taskId) {
  try {
    const ss = getDb_();
    const taskSheet = ss.getSheetByName(SHEET_TASKS);
    if (!taskSheet) throw new Error('ไม่พบแผ่นงาน Tasks');

    const archiveSheet = getArchiveSheet_(ss);
    const data = taskSheet.getDataRange().getValues();
    let rowIndex = -1;
    let targetRow = null;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(taskId)) {
        rowIndex = i + 1;
        targetRow = data[i];
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: 'ไม่พบรหัสงาน ' + taskId };
    }

    const archivedAtStr = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
    const archiveRow = [
      targetRow[0], targetRow[1], targetRow[2], targetRow[3], targetRow[4],
      targetRow[5], targetRow[6], targetRow[7], targetRow[8], targetRow[9],
      targetRow[10], targetRow[11], targetRow[12] || '', targetRow[13] || '',
      targetRow[14] || '[]', archivedAtStr
    ];

    archiveSheet.appendRow(archiveRow);
    taskSheet.deleteRow(rowIndex);
    SpreadsheetApp.flush();

    notifyDataChanged_({
      type: 'TASK_ARCHIVED',
      taskId: taskId
    });

    return { success: true, taskId: taskId };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * ดึงรายการงานที่จัดเก็บในคลังถาวร พร้อมระบบค้นหาและกรอง
 * @param {Object} params - { query: string, site: string, limit: number }
 */
function getArchivedTasks(params) {
  try {
    params = params || {};
    const ss = getDb_();
    const archiveSheet = getArchiveSheet_(ss);
    const data = archiveSheet.getDataRange().getValues();

    if (data.length <= 1) {
      return { success: true, tasks: [], totalCount: 0 };
    }

    const query = (params.query || '').trim().toLowerCase();
    const siteFilter = params.site || 'ALL';
    const limit = params.limit || 300;

    const tasks = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;

      const taskId = String(row[0]);
      const taskSite = String(row[1] || 'ไม่ระบุไซต์งาน').trim();
      const title = String(row[2] || '');
      const desc = String(row[3] || '');
      const assignee = String(row[4] || 'ไม่ระบุ');
      const priority = String(row[5] || 'ทั่วไป');
      const status = String(row[6] || 'เสร็จสิ้น');

      // ตรวจสอบเงื่อนไขไซต์งาน
      if (siteFilter !== 'ALL' && taskSite !== siteFilter) {
        continue;
      }

      let taskReporter = '';
      let taskReportDate = '';
      let taskChecklist = [];
      let archivedAtStr = '';
      if (row.length >= 16) {
        taskReporter = String(row[12] || '');
        if (row[13]) {
          taskReportDate = normalizeDateToIso_(row[13]);
        }
        if (row[14]) {
          try {
            const p = JSON.parse(row[14]);
            if (Array.isArray(p)) taskChecklist = p;
          } catch (e) {}
        }
        archivedAtStr = row[15] ? String(row[15]) : '';
      } else if (row.length === 15) {
        taskReporter = String(row[12] || '');
        if (row[13]) {
          taskReportDate = normalizeDateToIso_(row[13]);
        }
        archivedAtStr = row[14] ? String(row[14]) : '';
      } else if (row.length === 14) {
        taskReporter = String(row[12] || '');
        archivedAtStr = row[13] ? String(row[13]) : '';
      } else {
        archivedAtStr = row[12] ? String(row[12]) : '';
      }

      // ตรวจสอบเงื่อนไขคำค้นหา
      if (query) {
        const matchId = taskId.toLowerCase().includes(query);
        const matchTitle = title.toLowerCase().includes(query);
        const matchDesc = desc.toLowerCase().includes(query);
        const matchAssignee = assignee.toLowerCase().includes(query);
        const matchReporter = taskReporter.toLowerCase().includes(query);
        const matchSite = taskSite.toLowerCase().includes(query);
        if (!matchId && !matchTitle && !matchDesc && !matchAssignee && !matchReporter && !matchSite) {
          continue;
        }
      }

      let dueDateFormatted = '';
      if (row[7]) {
        dueDateFormatted = normalizeDateToIso_(row[7]);
      }

      let attachments = [];
      if (row[10]) {
        try {
          const parsed = JSON.parse(row[10]);
          if (Array.isArray(parsed)) attachments = parsed;
        } catch (e) {}
      }

      tasks.push({
        id: taskId,
        site: taskSite,
        title: title,
        description: desc,
        assignee: assignee,
        reporter: taskReporter,
        reportDate: taskReportDate,
        priority: priority,
        status: status,
        dueDate: dueDateFormatted,
        createdAt: row[8] ? String(row[8]) : '',
        updatedAt: row[9] ? String(row[9]) : '',
        attachments: attachments,
        folderId: row[11] ? String(row[11]) : '',
        checklist: taskChecklist,
        archivedAt: archivedAtStr
      });
    }

    // เรียงลำดับจากงานที่เพิ่งจัดเก็บบันทึกล่าสุดขึ้นก่อน (Latest archived first)
    tasks.reverse();

    const resultTasks = tasks.slice(0, limit);
    return { success: true, tasks: resultTasks, totalCount: tasks.length };
  } catch (err) {
    return { success: false, message: err.message, tasks: [] };
  }
}

/**
 * กู้คืนงานจากคลังถาวรกลับมายังระบบงานหลัก (Tasks)
 * @param {string} taskId
 */
function restoreArchivedTask(taskId) {
  try {
    const ss = getDb_();
    const taskSheet = ss.getSheetByName(SHEET_TASKS);
    const archiveSheet = getArchiveSheet_(ss);
    if (!taskSheet || !archiveSheet) throw new Error('ไม่พบชีตที่จำเป็น');

    const archiveData = archiveSheet.getDataRange().getValues();
    let rowIndex = -1;
    let targetRow = null;

    for (let i = 1; i < archiveData.length; i++) {
      if (String(archiveData[i][0]) === String(taskId)) {
        rowIndex = i + 1;
        targetRow = archiveData[i];
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: 'ไม่พบงานในคลัง ' + taskId };
    }

    // แถวสำหรับชีต Tasks (15 คอลัมน์ ตัด ArchivedAt ออก)
    const taskReporter = (targetRow.length >= 14) ? String(targetRow[12] || '') : '';
    let taskReportDate = '';
    if (targetRow.length >= 15 && targetRow[13]) {
      taskReportDate = normalizeDateToIso_(targetRow[13]);
    }
    const taskChecklistJson = (targetRow.length >= 16 && targetRow[14]) ? String(targetRow[14]) : '[]';

    const taskRow = [
      targetRow[0], targetRow[1], targetRow[2], targetRow[3], targetRow[4],
      targetRow[5], targetRow[6], targetRow[7], targetRow[8], targetRow[9],
      targetRow[10], targetRow[11], taskReporter, taskReportDate, taskChecklistJson
    ];

    taskSheet.appendRow(taskRow);
    archiveSheet.deleteRow(rowIndex);
    SpreadsheetApp.flush();

    let attachments = [];
    if (targetRow[10]) {
      try {
        const parsed = JSON.parse(targetRow[10]);
        if (Array.isArray(parsed)) attachments = parsed;
      } catch (e) {}
    }

    let dueDateFormatted = '';
    if (targetRow[7]) {
      dueDateFormatted = normalizeDateToIso_(targetRow[7]);
    }

    let restoredChecklist = [];
    try {
      restoredChecklist = JSON.parse(taskChecklistJson) || [];
    } catch (e) {}

    const restoredTask = {
      id: String(targetRow[0]),
      site: String(targetRow[1] || 'ไม่ระบุไซต์งาน').trim(),
      title: String(targetRow[2] || ''),
      description: String(targetRow[3] || ''),
      assignee: String(targetRow[4] || 'ไม่ระบุ'),
      reporter: taskReporter,
      reportDate: taskReportDate,
      priority: String(targetRow[5] || 'ทั่วไป'),
      status: String(targetRow[6] || 'เสร็จสิ้น'),
      dueDate: dueDateFormatted,
      createdAt: targetRow[8] ? String(targetRow[8]) : '',
      updatedAt: targetRow[9] ? String(targetRow[9]) : '',
      attachments: attachments,
      folderId: targetRow[11] ? String(targetRow[11]) : '',
      checklist: restoredChecklist
    };

    notifyDataChanged_({
      type: 'TASK_RESTORED',
      taskId: taskId,
      title: restoredTask.title
    });

    return { success: true, task: restoredTask };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * สร้างข้อความแจ้งเตือน Telegram ตามเหตุการณ์
 */
function buildTelegramMessage_(type, task, oldStatus) {
  const priorityIcons = {
    'ด่วนมาก': '🔴 [ด่วนมาก]',
    'สูง': '🟠 [สูง]',
    'ปานกลาง': '🟡 [ปานกลาง]',
    'ทั่วไป': '🟢 [ทั่วไป]'
  };

  const priorityLabel = priorityIcons[task.priority] || task.priority;

  let checklistInfo = '';
  if (Array.isArray(task.checklist) && task.checklist.length > 0) {
    const total = task.checklist.length;
    const done = task.checklist.filter(c => c.done).length;
    const totalFormatted = String(total).padStart(2, '0');
    checklistInfo = `☑️ <b>รายการย่อย:</b> เสร็จแล้ว ${done}/${totalFormatted} ข้อ\n`;
  }

  if (type === 'CREATE') {
    return (
      `📌 <b>มีงานใหม่เพิ่มเข้ามาในระบบ</b>\n` +
      `━━━━━━━━━━━━━━━━━━\n` +
      `🏢 <b>ไซต์งาน:</b> ${escapeHtml_(task.site)}\n` +
      `📋 <b>ชื่องาน:</b> ${escapeHtml_(task.title)}\n` +
      (task.reporter ? `📢 <b>ผู้แจ้ง:</b> ${escapeHtml_(task.reporter)}\n` : '') +
      (task.reportDate ? `🗓️ <b>วันที่แจ้ง:</b> ${task.reportDate}\n` : '') +
      `👤 <b>ผู้รับผิดชอบ:</b> ${escapeHtml_(task.assignee)}\n` +
      `🎯 <b>ความสำคัญ:</b> ${priorityLabel}\n` +
      `📅 <b>กำหนดส่ง:</b> ${task.dueDate || 'ไม่ระบุ'}\n` +
      checklistInfo +
      (task.description ? `📝 <b>รายละเอียด:</b> ${escapeHtml_(task.description)}\n` : '') +
      `🆔 <code>${task.id}</code>`
    );
  } else if (type === 'STATUS_CHANGE') {
    return (
      `🔄 <b>อัปเดตสถานะงาน</b>\n` +
      `━━━━━━━━━━━━━━━━━━\n` +
      `🏢 <b>ไซต์งาน:</b> ${escapeHtml_(task.site)}\n` +
      `📋 <b>ชื่องาน:</b> ${escapeHtml_(task.title)}\n` +
      (task.reporter ? `📢 <b>ผู้แจ้ง:</b> ${escapeHtml_(task.reporter)}\n` : '') +
      (task.reportDate ? `🗓️ <b>วันที่แจ้ง:</b> ${task.reportDate}\n` : '') +
      `👤 <b>ผู้รับผิดชอบ:</b> ${escapeHtml_(task.assignee)}\n` +
      `📊 <b>สถานะ:</b> ${oldStatus || 'เดิม'} ➡️ <b>${task.status}</b>\n` +
      checklistInfo +
      `🆔 <code>${task.id}</code>`
    );
  }
  return `📢 <b>แจ้งเตือนงาน:</b> ${task.title}`;
}

function escapeHtml_(text) {
  if (!text) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

/**
 * ค้นหา Telegram Chat ID ส่วนตัวของผู้รับผิดชอบจากชีต Users
 * โดยเทียบจากชื่อ-นามสกุล (FullName) หรือ Username
 */
function getAssigneeTelegramChatId_(assigneeName) {
  if (!assigneeName || assigneeName === 'ไม่ระบุ') return null;
  try {
    const target = String(assigneeName).trim().toLowerCase();
    const ss = getDb_();
    const userSheet = ss.getSheetByName(SHEET_USERS);
    if (!userSheet) return null;

    const data = userSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const uName = String(data[i][0] || '').trim().toLowerCase();
      const fName = String(data[i][1] || '').trim().toLowerCase();
      const chatId = String(data[i][2] || '').trim();
      const status = String(data[i][4] || '').trim().toLowerCase();

      // เทียบกับ Username หรือ FullName และต้องไม่ใช่ Inactive
      if ((uName === target || fName === target) && status !== 'inactive') {
        return chatId || null;
      }
    }
    return null;
  } catch (err) {
    console.error('getAssigneeTelegramChatId_ error: ' + err.message);
    return null;
  }
}

/**
 * สร้างข้อความแจ้งเตือนส่วนตัวส่งตรงถึงผู้รับผิดชอบงาน
 */
function buildPersonalAssigneeMessage_(type, task) {
  const priorityIcons = {
    'ด่วนมาก': '🔴 [ด่วนมาก]',
    'สูง': '🟠 [สูง]',
    'ปานกลาง': '🟡 [ปานกลาง]',
    'ทั่วไป': '🟢 [ทั่วไป]'
  };

  const priorityLabel = priorityIcons[task.priority] || task.priority;

  if (type === 'CREATE') {
    return (
      `🔔 <b>คุณได้รับมอบหมายงานใหม่!</b>\n` +
      `━━━━━━━━━━━━━━━━━━\n` +
      `🏢 <b>ไซต์งาน:</b> ${escapeHtml_(task.site)}\n` +
      `📋 <b>ชื่องาน:</b> ${escapeHtml_(task.title)}\n` +
      (task.reporter ? `📢 <b>ผู้แจ้ง:</b> ${escapeHtml_(task.reporter)}\n` : '') +
      (task.reportDate ? `🗓️ <b>วันที่แจ้ง:</b> ${task.reportDate}\n` : '') +
      `🎯 <b>ความสำคัญ:</b> ${priorityLabel}\n` +
      `📅 <b>กำหนดส่ง:</b> ${task.dueDate || 'ไม่ระบุ'}\n` +
      (task.description ? `📝 <b>รายละเอียด:</b> ${escapeHtml_(task.description)}\n` : '') +
      `━━━━━━━━━━━━━━━━━━\n` +
      `🆔 <code>${task.id}</code>\n` +
      `👉 <i>กรุณาตรวจสอบรายละเอียดและดำเนินการในระบบ Work Tracker</i>`
    );
  }
  return `📢 <b>แจ้งเตือนงาน:</b> ${task.title}`;
}

/**
 * ส่งข้อความไปยัง Telegram Bot API (ฟังก์ชันภายใน)
 */
function sendTelegramRaw_(token, chatId, htmlText) {
  const url = `https://api.telegram.org/bot${token}/sendMessage`;
  const payload = {
    chat_id: chatId,
    text: htmlText,
    parse_mode: 'HTML',
    disable_web_page_preview: true
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const resCode = response.getResponseCode();
  const resText = response.getContentText();

  if (resCode !== 200) {
    throw new Error(`Telegram API Error (${resCode}): ${resText}`);
  }

  return JSON.parse(resText);
}

/**
 * ส่งแจ้งเตือนงานไปยัง Telegram กลุ่ม (รองรับการแยกกลุ่มตามไซต์งาน)
 * @param {string} type 'CREATE' | 'STATUS_CHANGE'
 * @param {Object} task ข้อมูลงาน
 * @param {string} [oldStatus] สถานะเดิม (กรณี STATUS_CHANGE)
 */
function notifyTelegramGroupForTask_(type, task, oldStatus) {
  try {
    const settings = getSettings_();
    if (!settings.TELEGRAM_BOT_TOKEN) return;

    if (type === 'CREATE' && !settings.NOTIFY_ON_CREATE) return;
    if (type === 'STATUS_CHANGE' && !settings.NOTIFY_ON_STATUS_CHANGE) return;

    // หา Chat ID กลุ่มเป้าหมาย
    const targetChatIds = new Set();
    const siteName = String(task.site || '').trim();
    const siteChatIds = settings.SITE_TELEGRAM_CHAT_IDS || {};
    const siteSpecificChatId = siteChatIds[siteName] ? String(siteChatIds[siteName]).trim() : '';

    if (siteSpecificChatId) {
      targetChatIds.add(siteSpecificChatId);
      // หากเปิดตัวเลือกส่งเข้ากลุ่มหลักส่วนกลางด้วยเสมอ
      if (settings.NOTIFY_MAIN_GROUP_ALWAYS && settings.TELEGRAM_CHAT_ID) {
        const mainId = String(settings.TELEGRAM_CHAT_ID).trim();
        if (mainId && mainId !== siteSpecificChatId) targetChatIds.add(mainId);
      }
    } else if (settings.TELEGRAM_CHAT_ID) {
      // ไซต์นี้ไม่มี Chat ID เฉพาะ -> ส่งเข้ากลุ่มหลัก (Fallback)
      const mainId = String(settings.TELEGRAM_CHAT_ID).trim();
      if (mainId) targetChatIds.add(mainId);
    }

    if (targetChatIds.size === 0) return;

    const msg = buildTelegramMessage_(type, task, oldStatus);
    targetChatIds.forEach(chatId => {
      try {
        sendTelegramRaw_(settings.TELEGRAM_BOT_TOKEN, chatId, msg);
      } catch (err) {
        console.error(`Telegram send to group (${chatId}) for site (${siteName}) failed: ${err.message}`);
      }
    });
  } catch (e) {
    console.error('notifyTelegramGroupForTask_ error: ' + e.message);
  }
}

/**
 * ทดสอบการเชื่อมต่อ Telegram Bot (รองรับทั้งกลุ่มหลักและกลุ่มประจำไซต์งาน)
 */
function testTelegramConnection(token, chatId, siteName) {
  try {
    let activeToken = String(token || '').trim();
    if (!activeToken || activeToken.includes('****')) {
      const settings = getSettings_();
      activeToken = settings.TELEGRAM_BOT_TOKEN;
    }

    const activeChatId = String(chatId || '').trim();

    if (!activeToken || !activeChatId) {
      return { success: false, message: 'กรุณากรอกทั้ง Bot Token และ Chat ID สำหรับทดสอบ' };
    }

    const nowStr = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss');
    let testMsg = '';

    if (siteName && String(siteName).trim()) {
      testMsg = (
        `✅ <b>ทดสอบการเชื่อมต่อกลุ่มไซต์งานสำเร็จ!</b>\n` +
        `━━━━━━━━━━━━━━━━━━\n` +
        `🏢 <b>ไซต์งาน:</b> ${escapeHtml_(siteName)}\n` +
        `🤖 <b>ระบบ:</b> ระบบติดตามงาน (Work Tracker)\n` +
        `💬 <b>กลุ่ม Chat ID:</b> <code>${activeChatId}</code>\n` +
        `🕒 <b>วันที่ทดสอบ:</b> ${nowStr}`
      );
    } else {
      testMsg = (
        `✅ <b>ทดสอบการเชื่อมต่อสำเร็จ!</b>\n` +
        `━━━━━━━━━━━━━━━━━━\n` +
        `ระบบติดตามงาน เชื่อมต่อกับ Telegram สำเร็จแล้ว\n` +
        `💬 <b>Chat ID:</b> <code>${activeChatId}</code>\n` +
        `🕒 <b>วันที่ทดสอบ:</b> ${nowStr}`
      );
    }

    sendTelegramRaw_(activeToken, activeChatId, testMsg);
    return { 
      success: true, 
      message: siteName 
        ? `ส่งข้อความทดสอบไปยังกลุ่มของไซต์ "${siteName}" สำเร็จแล้ว!` 
        : 'ส่งข้อความทดสอบไปยัง Telegram สำเร็จแล้ว!' 
    };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * =========================================================================
 * AUTHENTICATION & USER MANAGEMENT (TELEGRAM OTP / PIN)
 * =========================================================================
 */

/**
 * ผู้ใช้ขอรหัส PIN 6 หลักเพื่อเข้าสู่ระบบ (ส่งเข้า Telegram ส่วนตัว)
 */
function requestLoginPin(username) {
  try {
    if (!username) {
      return { success: false, message: 'กรุณากรอกชื่อผู้ใช้ (Username)' };
    }

    username = String(username).trim().toLowerCase();
    const ss = getDb_();
    let userSheet = ss.getSheetByName(SHEET_USERS);
    if (!userSheet) {
      setupDatabase();
      userSheet = ss.getSheetByName(SHEET_USERS);
    }

    const userData = userSheet.getDataRange().getValues();
    let foundUser = null;

    for (let i = 1; i < userData.length; i++) {
      const uName = String(userData[i][0]).trim().toLowerCase();
      if (uName === username) {
        foundUser = {
          rowIndex: i + 1,
          username: String(userData[i][0]).trim(),
          fullName: String(userData[i][1]).trim(),
          telegramChatId: String(userData[i][2]).trim(),
          role: String(userData[i][3]).trim() || 'Member',
          status: String(userData[i][4]).trim() || 'Active'
        };
        break;
      }
    }

    if (!foundUser) {
      return { success: false, message: `ไม่พบชื่อผู้ใช้ "${username}" ในระบบ กรุณาติดต่อผู้ดูแลระบบ` };
    }

    if (foundUser.status.toLowerCase() !== 'active') {
      return { success: false, message: 'บัญชีผู้ใช้นี้ถูกระงับการใช้งาน กรุณาติดต่อผู้ดูแลระบบ' };
    }

    if (!foundUser.telegramChatId) {
      return { 
        success: false, 
        message: `ผู้ใช้ "${foundUser.username}" ยังไม่ได้ผูก Telegram Chat ID กรุณาติดต่อผู้ดูแลระบบเพื่อเพิ่ม Chat ID ของคุณ` 
      };
    }

    const settings = getSettings_();
    if (!settings.TELEGRAM_BOT_TOKEN) {
      return { success: false, message: 'ระบบยังไม่ได้ตั้งค่า Telegram Bot Token กรุณาให้ผู้ดูแลระบบตั้งค่าในเมนู Settings' };
    }

    // สุ่มรหัส PIN 6 หลัก (100000 - 999999)
    const pin = String(Math.floor(100000 + Math.random() * 900000));

    // จัดเก็บใน CacheService กำหนดอายุ 300 วินาที (5 นาที)
    const cache = CacheService.getScriptCache();
    cache.put('PIN_' + username, pin, 300);
    cache.put('ATTEMPTS_' + username, '0', 300);

    // ส่งข้อความแจ้งรหัสเข้า Telegram ส่วนตัว
    const msg = (
      `🔢 <code>${pin}</code> คือรหัส PIN สำหรับเข้าสู่ระบบของคุณ\n` +
      `━━━━━━━━━━━━━━━━━━\n` +
      `🏢 <b>ระบบติดตามงาน</b>\n` +
      `👤 <b>ผู้ใช้:</b> ${escapeHtml_(foundUser.fullName)} (<code>${foundUser.username}</code>)\n` +
      `⏳ รหัสนี้มีอายุ <b>5 นาที</b> และใช้ได้เพียงครั้งเดียว\n` +
      `⚠️ <i>กรุณาอย่านำรหัสนี้เปิดเผยแก่ผู้อื่น</i>`
    );

    sendTelegramRaw_(settings.TELEGRAM_BOT_TOKEN, foundUser.telegramChatId, msg);

    // ปิดบังหมายเลข Chat ID เพื่อความเป็นส่วนตัว
    const maskedChatId = foundUser.telegramChatId.length > 4 
      ? '***' + foundUser.telegramChatId.slice(-4) 
      : '****';

    return {
      success: true,
      message: `ส่งรหัส PIN ไปยัง Telegram (${maskedChatId}) เรียบร้อยแล้ว`,
      username: foundUser.username,
      fullName: foundUser.fullName,
      expiresIn: 300
    };
  } catch (error) {
    return { success: false, message: 'ส่งรหัส PIN ไม่สำเร็จ: ' + error.message };
  }
}

/**
 * ตรวจสอบรหัส PIN 6 หลักเพื่อเข้าสู่ระบบ
 */
function verifyLoginPin(username, pin) {
  try {
    if (!username || !pin) {
      return { success: false, message: 'กรุณากรอกทั้งชื่อผู้ใช้และรหัส PIN' };
    }

    username = String(username).trim().toLowerCase();
    pin = String(pin).trim();

    const cache = CacheService.getScriptCache();
    const cachedPin = cache.get('PIN_' + username);

    if (!cachedPin) {
      return { success: false, message: 'รหัส PIN หมดอายุหรือยังไม่ได้ขอรหัส กรุณากดขอรหัสใหม่อีกครั้ง' };
    }

    let attempts = parseInt(cache.get('ATTEMPTS_' + username) || '0', 10);
    if (attempts >= 3) {
      cache.remove('PIN_' + username);
      cache.remove('ATTEMPTS_' + username);
      return { success: false, message: 'คุณกรอกรหัสผิดเกิน 3 ครั้ง รหัสถูกยกเลิกแล้ว กรุณากดขอรหัส PIN ใหม่' };
    }

    if (cachedPin !== pin) {
      attempts++;
      const remaining = 3 - attempts;
      if (remaining > 0) {
        cache.put('ATTEMPTS_' + username, String(attempts), 300);
        return { success: false, message: `รหัส PIN ไม่ถูกต้อง (เหลือโอกาสกรอกอีก ${remaining} ครั้ง)` };
      } else {
        cache.remove('PIN_' + username);
        cache.remove('ATTEMPTS_' + username);
        return { success: false, message: 'คุณกรอกรหัสผิดเกินกำหนด รหัสถูกยกเลิกแล้ว กรุณากดขอรหัสใหม่' };
      }
    }

    // รหัสถูกต้อง เคลียร์รหัสออกจาก Cache (One-Time PIN)
    cache.remove('PIN_' + username);
    cache.remove('ATTEMPTS_' + username);

    // ดึงข้อมูลโปรไฟล์ผู้ใช้ และบันทึก LastLogin
    const ss = getDb_();
    let userSheet = ss.getSheetByName(SHEET_USERS);
    if (!userSheet) {
      setupDatabase();
      userSheet = ss.getSheetByName(SHEET_USERS);
    }

    const userData = userSheet.getDataRange().getValues();
    let profile = null;
    const nowStr = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');

    for (let i = 1; i < userData.length; i++) {
      if (String(userData[i][0]).trim().toLowerCase() === username) {
        let allowedSites = 'ALL';
        if (userData[i][6] !== undefined && String(userData[i][6]).trim() !== '') {
          const raw = String(userData[i][6]).trim();
          if (raw !== 'ALL') {
            try {
              const parsed = JSON.parse(raw);
              if (Array.isArray(parsed)) allowedSites = parsed;
            } catch (e) {
              allowedSites = 'ALL';
            }
          }
        }

        profile = {
          username: String(userData[i][0]).trim(),
          fullName: String(userData[i][1]).trim(),
          role: String(userData[i][3]).trim() || 'Member',
          allowedSites: allowedSites
        };
        userSheet.getRange(i + 1, 6).setValue(nowStr);
        break;
      }
    }

    const sessionToken = Utilities.getUuid();

    return {
      success: true,
      message: 'เข้าสู่ระบบสำเร็จ',
      token: sessionToken,
      user: profile || { username: username, fullName: username, role: 'Member', allowedSites: 'ALL' }
    };
  } catch (error) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + error.message };
  }
}

/**
 * ดึงรายชื่อผู้ใช้ทั้งหมดในระบบ (สำหรับหน้าต่างตั้งค่า)
 */
function getUsersList() {
  try {
    const ss = getDb_();
    let sheet = ss.getSheetByName(SHEET_USERS);
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName(SHEET_USERS);
    }

    const rows = sheet.getDataRange().getValues();
    const users = [];

    for (let i = 1; i < rows.length; i++) {
      if (!rows[i][0]) continue;
      let allowedSites = 'ALL';
      if (rows[i][6] !== undefined && String(rows[i][6]).trim() !== '') {
        const raw = String(rows[i][6]).trim();
        if (raw !== 'ALL') {
          try {
            const parsed = JSON.parse(raw);
            if (Array.isArray(parsed)) allowedSites = parsed;
          } catch (e) {
            allowedSites = 'ALL';
          }
        }
      }

      users.push({
        username: String(rows[i][0]).trim(),
        fullName: String(rows[i][1]).trim(),
        telegramChatId: String(rows[i][2]).trim(),
        role: String(rows[i][3]).trim() || 'Member',
        status: String(rows[i][4]).trim() || 'Active',
        lastLogin: rows[i][5] ? String(rows[i][5]) : '-',
        allowedSites: allowedSites
      });
    }

    return { success: true, users: users };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * บันทึกผู้ใช้ (เพิ่มผู้ใช้ใหม่ หรือแก้ไขผู้ใช้เดิม)
 */
function saveUser(userData) {
  try {
    if (!userData.username || !userData.fullName) {
      return { success: false, message: 'กรุณากรอก Username และชื่อ-นามสกุล' };
    }

    const username = String(userData.username).trim().toLowerCase();
    const ss = getDb_();
    let sheet = ss.getSheetByName(SHEET_USERS);
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName(SHEET_USERS);
    }

    const data = sheet.getDataRange().getValues();
    let targetRow = -1;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === username) {
        targetRow = i + 1;
        break;
      }
    }

    let allowedSitesVal = 'ALL';
    if (userData.allowedSites && userData.allowedSites !== 'ALL') {
      if (Array.isArray(userData.allowedSites)) {
        allowedSitesVal = JSON.stringify(userData.allowedSites);
      } else {
        allowedSitesVal = String(userData.allowedSites);
      }
    }

    const rowValues = [
      username,
      String(userData.fullName).trim(),
      String(userData.telegramChatId || '').trim(),
      String(userData.role || 'Member').trim(),
      String(userData.status || 'Active').trim()
    ];

    if (targetRow > 0) {
      // แก้ไขข้อมูลเดิม: อัปเดตคอลัมน์ 1-5 และคอลัมน์ 7 (AllowedSites) โดยคงคอลัมน์ 6 (LastLogin) ไว้
      sheet.getRange(targetRow, 1, 1, 5).setValues([rowValues]);
      sheet.getRange(targetRow, 7).setValue(allowedSitesVal);
    } else {
      // เพิ่มผู้ใช้ใหม่: [username, fullName, chatId, role, status, lastLogin='', allowedSites]
      sheet.appendRow([...rowValues, '', allowedSitesVal]);
    }

    return { success: true, message: 'บันทึกข้อมูลผู้ใช้เรียบร้อย' };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * ลบผู้ใช้
 */
function deleteUser(username) {
  try {
    username = String(username).trim().toLowerCase();
    if (username === 'admin') {
      return { success: false, message: 'ไม่สามารถลบบัญชี admin ได้' };
    }

    const ss = getDb_();
    const sheet = ss.getSheetByName(SHEET_USERS);
    if (!sheet) return { success: false, message: 'ไม่พบตาราง Users' };

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === username) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'ลบผู้ใช้เรียบร้อยแล้ว' };
      }
    }

    return { success: false, message: 'ไม่พบผู้ใช้นี้ในระบบ' };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * =========================================================================
 * REAL-TIME PRESENCE & ONLINE USERS (ระบบตรวจจับและแสดงผู้ใช้งานออนไลน์)
 * ใช้งานผ่าน CacheService (Script Cache) เพื่อความรวดเร็วและไม่กระทบ Quota ของ Sheet
 * =========================================================================
 */

const CACHE_KEY_ONLINE_USERS = 'APP_ONLINE_USERS';
const ONLINE_TIMEOUT_MS = 90 * 1000; // 90 วินาที (หากไม่ส่ง heartbeat จะถือว่าออฟไลน์)

/**
 * ผู้ใช้ส่งสัญญาณ Heartbeat ว่ากำลังใช้งานอยู่
 */
function pingOnlineUser(userData) {
  try {
    if (!userData || !userData.username) {
      return getOnlineUsers();
    }

    const username = String(userData.username).trim().toLowerCase();
    const fullName = String(userData.fullName || userData.username).trim();
    const role = String(userData.role || 'Member').trim();

    const lock = LockService.getScriptLock();
    try {
      lock.tryLock(2000);
    } catch (lockErr) {}

    const cache = CacheService.getScriptCache();
    const raw = cache.get(CACHE_KEY_ONLINE_USERS);
    let usersMap = {};

    if (raw) {
      try {
        usersMap = JSON.parse(raw) || {};
      } catch (e) {
        usersMap = {};
      }
    }

    const now = Date.now();
    // ตัดผู้ใช้ที่ไม่มี heartbeat เกิน 90 วินาทีออก
    for (const u in usersMap) {
      if (!usersMap[u] || !usersMap[u].lastActive || (now - usersMap[u].lastActive > ONLINE_TIMEOUT_MS)) {
        delete usersMap[u];
      }
    }

    // เพิ่มหรืออัปเดตผู้ใช้ปัจจุบัน
    usersMap[username] = {
      username: username,
      fullName: fullName,
      role: role,
      lastActive: now
    };

    // บันทึกลง Script Cache มีอายุ 600 วินาที (10 นาที)
    cache.put(CACHE_KEY_ONLINE_USERS, JSON.stringify(usersMap), 600);

    try { lock.releaseLock(); } catch (e) {}

    const list = Object.values(usersMap).sort((a, b) => b.lastActive - a.lastActive);
    return {
      success: true,
      count: list.length,
      users: list
    };
  } catch (error) {
    return {
      success: false,
      count: 1,
      users: userData && userData.username ? [{
        username: userData.username,
        fullName: userData.fullName || userData.username,
        role: userData.role || 'Member',
        lastActive: Date.now()
      }] : [],
      error: error.message
    };
  }
}

/**
 * ดึงรายชื่อผู้ใช้ที่กำลังออนไลน์อยู่
 */
function getOnlineUsers() {
  try {
    const cache = CacheService.getScriptCache();
    const raw = cache.get(CACHE_KEY_ONLINE_USERS);
    if (!raw) {
      return { success: true, count: 0, users: [] };
    }

    let usersMap = {};
    try {
      usersMap = JSON.parse(raw) || {};
    } catch (e) {
      return { success: true, count: 0, users: [] };
    }

    const now = Date.now();
    let hasExpired = false;

    for (const u in usersMap) {
      if (!usersMap[u] || !usersMap[u].lastActive || (now - usersMap[u].lastActive > ONLINE_TIMEOUT_MS)) {
        delete usersMap[u];
        hasExpired = true;
      }
    }

    if (hasExpired) {
      cache.put(CACHE_KEY_ONLINE_USERS, JSON.stringify(usersMap), 600);
    }

    const list = Object.values(usersMap).sort((a, b) => b.lastActive - a.lastActive);
    return {
      success: true,
      count: list.length,
      users: list
    };
  } catch (error) {
    return { success: false, count: 0, users: [], error: error.message };
  }
}

/**
 * แจ้งออกจากระบบ ลบผู้ใช้ออกจากรายการออนไลน์ทันที
 */
function logoutOnlineUser(username) {
  try {
    if (!username) return { success: true };
    username = String(username).trim().toLowerCase();

    const lock = LockService.getScriptLock();
    try {
      lock.tryLock(2000);
    } catch (lockErr) {}

    const cache = CacheService.getScriptCache();
    const raw = cache.get(CACHE_KEY_ONLINE_USERS);
    if (raw) {
      let usersMap = {};
      try {
        usersMap = JSON.parse(raw) || {};
      } catch (e) {
        usersMap = {};
      }
      if (usersMap[username]) {
        delete usersMap[username];
        cache.put(CACHE_KEY_ONLINE_USERS, JSON.stringify(usersMap), 600);
      }
    }

    try { lock.releaseLock(); } catch (e) {}
    return { success: true };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * =========================================================================
 * GOOGLE DRIVE FILE ATTACHMENTS & FOLDER MANAGEMENT
 * จัดการโฟลเดอร์ตามไซต์งานและรหัสงาน พร้อมอัปโหลดและลบไฟล์แนบ
 * =========================================================================
 */

const ROOT_ATTACHMENTS_FOLDER_NAME = 'Work Tracker Attachments';

/**
 * ดึงหรือสร้างโฟลเดอร์หลักสำหรับจัดเก็บไฟล์แนบทั้งหมด
 */
function getAttachmentsRootFolder_() {
  const props = PropertiesService.getScriptProperties();
  const cachedRootId = props.getProperty('ATTACHMENTS_ROOT_FOLDER_ID');
  if (cachedRootId) {
    try {
      const cachedFolder = DriveApp.getFolderById(cachedRootId);
      if (cachedFolder && !cachedFolder.isTrashed()) {
        return cachedFolder;
      }
    } catch (e) {
      console.warn('Cached root folder invalid: ' + e.message);
    }
  }

  const ss = getDb_();
  let parentFolder = null;
  try {
    const ssFile = DriveApp.getFileById(ss.getId());
    const parents = ssFile.getParents();
    if (parents.hasNext()) {
      parentFolder = parents.next();
    }
  } catch (e) {
    console.warn('Cannot get parent folder of spreadsheet: ' + e.message);
  }

  const rootTarget = parentFolder || DriveApp.getRootFolder();
  const folders = rootTarget.getFoldersByName(ROOT_ATTACHMENTS_FOLDER_NAME);
  let folder;
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = rootTarget.createFolder(ROOT_ATTACHMENTS_FOLDER_NAME);
  }

  if (folder) {
    try {
      props.setProperty('ATTACHMENTS_ROOT_FOLDER_ID', folder.getId());
    } catch (e) {}
  }

  return folder;
}

/**
 * ค้นหาหรือสร้างโฟลเดอร์ประจำไซต์งาน
 */
function getOrCreateSiteFolder_(siteName) {
  const root = getAttachmentsRootFolder_();
  const cleanSite = String(siteName || 'ไม่ระบุไซต์งาน').replace(/[\\/:*?"<>|]/g, '_').trim();
  const folders = root.getFoldersByName(cleanSite);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return root.createFolder(cleanSite);
  }
}

/**
 * ค้นหาหรือสร้างโฟลเดอร์ประจำรหัสงาน [Root] / [Site] / [TaskID]
 */
function getOrCreateTaskFolder_(siteName, taskId) {
  const siteFolder = getOrCreateSiteFolder_(siteName);
  const cleanTask = String(taskId).trim();
  const folders = siteFolder.getFoldersByName(cleanTask);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return siteFolder.createFolder(cleanTask);
  }
}

/**
 * ย้ายโฟลเดอร์งานไปยังไซต์งานใหม่ (กรณีย้ายไซต์งาน)
 */
function relocateTaskFolder_(folderId, newSiteName) {
  if (!folderId || !newSiteName) return;
  try {
    const taskFolder = DriveApp.getFolderById(folderId);
    const targetSiteFolder = getOrCreateSiteFolder_(newSiteName);
    
    const parents = taskFolder.getParents();
    if (parents.hasNext()) {
      const currentParent = parents.next();
      if (currentParent.getId() === targetSiteFolder.getId()) {
        return; // อยู่ในโฟลเดอร์ไซต์ที่ถูกต้องแล้ว
      }
      taskFolder.moveTo(targetSiteFolder);
    }
  } catch (e) {
    console.warn('relocateTaskFolder_ error: ' + e.message);
  }
}

/**
 * ลบโฟลเดอร์งานและไฟล์ทั้งหมดใน Google Drive ลงถังขยะ
 */
function deleteTaskFolderAndFiles_(folderId, siteName, taskId) {
  // 1. ถ้ามี Folder ID ตรงๆ ให้ย้ายลงถังขยะทันที
  if (folderId) {
    try {
      const folder = DriveApp.getFolderById(folderId);
      if (folder) {
        folder.setTrashed(true);
        return true;
      }
    } catch (e) {
      console.warn('Folder not found by ID: ' + e.message);
    }
  }

  // 2. ถ้าไม่มี ID หรือหาไม่พบ ให้ลองหาตามโครงสร้าง [Root] / [Site] / [TaskID]
  try {
    const root = getAttachmentsRootFolder_();
    const cleanSite = String(siteName || 'ไม่ระบุไซต์งาน').replace(/[\\/:*?"<>|]/g, '_').trim();
    const siteFolders = root.getFoldersByName(cleanSite);
    if (siteFolders.hasNext()) {
      const siteFolder = siteFolders.next();
      const cleanTask = String(taskId).trim();
      const taskFolders = siteFolder.getFoldersByName(cleanTask);
      if (taskFolders.hasNext()) {
        const taskFolder = taskFolders.next();
        taskFolder.setTrashed(true);
        return true;
      }
    }
  } catch (e) {
    console.warn('deleteTaskFolderAndFiles_ search error: ' + e.message);
  }
  return false;
}

/**
 * อัปโหลดไฟล์แนบเข้าสู่โฟลเดอร์งานใน Google Drive
 * รับข้อมูล: taskId, site, fileData { fileName, mimeType, base64, size }
 */
function uploadTaskAttachment(taskId, site, fileData) {
  try {
    if (!taskId || !fileData || !fileData.base64) {
      return { success: false, message: 'ข้อมูลไฟล์ไม่ครบถ้วน' };
    }

    const ss = getDb_();
    const sheet = ss.getSheetByName(SHEET_TASKS);
    if (!sheet) throw new Error('ไม่พบแผ่นงาน Tasks');

    // บังคับ Flush ให้มั่นใจว่าได้ข้อมูลแถวใหม่ล่าสุดที่เพิ่งบันทึก
    SpreadsheetApp.flush();

    let values = sheet.getDataRange().getValues();
    let rowIndex = -1;
    let targetRow = null;

    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]).trim() === String(taskId).trim()) {
        rowIndex = i + 1;
        targetRow = values[i];
        break;
      }
    }

    // หากยังไม่พบ (อาจเกิดจาก Sheets replication latency) ให้รอ 300ms แล้ว flush ซ้ำ
    if (rowIndex === -1) {
      Utilities.sleep(300);
      SpreadsheetApp.flush();
      values = sheet.getDataRange().getValues();
      for (let i = 1; i < values.length; i++) {
        if (String(values[i][0]).trim() === String(taskId).trim()) {
          rowIndex = i + 1;
          targetRow = values[i];
          break;
        }
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: `ไม่พบงานรหัส ${taskId} ในชีต (กรุณาลองใหม่อีกครั้ง)` };
    }

    // 1. ดึงหรือสร้างโฟลเดอร์ของงานใน Drive
    let folderId = targetRow[11] ? String(targetRow[11]).trim() : '';
    let taskFolder = null;

    if (folderId) {
      try {
        taskFolder = DriveApp.getFolderById(folderId);
      } catch (e) {
        taskFolder = null;
      }
    }

    if (!taskFolder) {
      const taskSite = site || String(targetRow[1]) || 'ไม่ระบุไซต์งาน';
      taskFolder = getOrCreateTaskFolder_(taskSite, taskId);
      folderId = taskFolder.getId();
      // บันทึก Folder ID ลงใน Column 12
      sheet.getRange(rowIndex, 12).setValue(folderId);
      SpreadsheetApp.flush();
    }

    // 2. แปลง Base64 และสร้างไฟล์ในโฟลเดอร์
    let cleanBase64 = String(fileData.base64);
    const commaIndex = cleanBase64.indexOf(',');
    if (commaIndex !== -1) {
      cleanBase64 = cleanBase64.substring(commaIndex + 1);
    }
    cleanBase64 = cleanBase64.replace(/\s/g, '');

    const decoded = Utilities.base64Decode(cleanBase64);
    const mimeType = fileData.mimeType || 'application/octet-stream';
    const fileName = fileData.fileName || 'unnamed_file';
    const blob = Utilities.newBlob(decoded, mimeType, fileName);
    const driveFile = taskFolder.createFile(blob);

    // เปิดสิทธิ์เปิดดูได้ด้วยลิงก์ (ถ้าองค์กรอนุญาต)
    try {
      driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) {}

    const fileMeta = {
      id: driveFile.getId(),
      name: driveFile.getName(),
      mimeType: driveFile.getMimeType(),
      size: driveFile.getSize(),
      url: driveFile.getUrl(),
      downloadUrl: `https://drive.google.com/uc?export=download&id=${driveFile.getId()}`,
      uploadedAt: Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm')
    };

    // 3. อ่านรายการไฟล์เดิมใน Column 11 แล้วเพิ่มไฟล์ใหม่เข้าไป
    let currentAttachments = [];
    if (targetRow[10]) {
      try {
        const parsed = JSON.parse(targetRow[10]);
        if (Array.isArray(parsed)) currentAttachments = parsed;
      } catch (e) {}
    }

    currentAttachments.push(fileMeta);

    // บันทึกกลับลงชีต
    sheet.getRange(rowIndex, 11).setValue(JSON.stringify(currentAttachments));
    sheet.getRange(rowIndex, 10).setValue(Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm')); // UpdatedAt
    SpreadsheetApp.flush();

    return {
      success: true,
      file: fileMeta,
      attachments: currentAttachments,
      folderId: folderId
    };
  } catch (error) {
    console.error('uploadTaskAttachment error: ' + error.message);
    return { success: false, message: error.message };
  }
}

/**
 * ลบไฟล์แนบเดี่ยวออกจากงานและ Google Drive
 */
function deleteTaskAttachment(taskId, fileId) {
  try {
    const ss = getDb_();
    const sheet = ss.getSheetByName(SHEET_TASKS);
    if (!sheet) throw new Error('ไม่พบแผ่นงาน Tasks');

    const values = sheet.getDataRange().getValues();
    let rowIndex = -1;
    let targetRow = null;

    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === String(taskId)) {
        rowIndex = i + 1;
        targetRow = values[i];
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: 'ไม่พบงาน' };
    }

    // ย้ายไฟล์ลงถังขยะใน Drive
    try {
      const file = DriveApp.getFileById(fileId);
      if (file) file.setTrashed(true);
    } catch (e) {
      console.warn('File not found in Drive: ' + e.message);
    }

    // อัปเดตรายการใน Column 11
    let currentAttachments = [];
    if (targetRow[10]) {
      try {
        const parsed = JSON.parse(targetRow[10]);
        if (Array.isArray(parsed)) currentAttachments = parsed;
      } catch (e) {}
    }

    const updatedAttachments = currentAttachments.filter(f => f.id !== fileId);
    sheet.getRange(rowIndex, 11).setValue(JSON.stringify(updatedAttachments));
    sheet.getRange(rowIndex, 10).setValue(Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm'));

    return {
      success: true,
      attachments: updatedAttachments
    };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

