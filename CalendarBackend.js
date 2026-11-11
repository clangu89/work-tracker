/**
 * =========================================================================
 * SITE CALENDAR BACKEND MODULE (CalendarBackend.js)
 * Google Apps Script Backend for Site Duty & Visit Schedule Management
 * ระบบปฏิทินเข้าไซต์งาน: กำหนดวัน/ไซต์/ผู้ปฏิบัติงาน, แสดงผลปฏิทินรายเดือน,
 * แจ้งเตือน Telegram ทันที และแจ้งเตือนล่วงหน้า 1 วัน (Time Trigger 07:00 น.)
 * =========================================================================
 */

const SHEET_SITE_SCHEDULES = 'SiteSchedules';

/**
 * สร้างและจัดฟอร์แมตโครงสร้างชีตสำหรับระบบปฏิทินเข้าไซต์งาน (SiteSchedules)
 */
function setupSiteSchedulesDatabase_(ss) {
  if (!ss) ss = getDb_();

  let sheet = ss.getSheetByName(SHEET_SITE_SCHEDULES);
  const headers = [
    'ScheduleId',      // Col 1: A - รหัสรายการ (e.g. SCH-20261110-1001)
    'Date',            // Col 2: B - วันที่เข้าไซต์งาน (YYYY-MM-DD)
    'Site',            // Col 3: C - ชื่อไซต์งาน
    'Username',        // Col 4: D - Username ของผู้เข้าไซต์
    'FullName',        // Col 5: E - ชื่อ-นามสกุลแสดงผล
    'Purpose',         // Col 6: F - วัตถุประสงค์สั้นๆ หรือรายละเอียดงาน
    'TaskId',          // Col 7: G - รหัส Task ใน Work Tracker (Optional)
    'CreatedBy',       // Col 8: H - Username ผู้บันทึกมอบหมาย
    'CreatedAt',       // Col 9: I - วันที่ทำรายการ
    'ReminderSent'     // Col 10: J - สถานะส่งแจ้งเตือนล่วงหน้า 1 วันแล้วหรือยัง (TRUE/FALSE)
  ];

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_SITE_SCHEDULES);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (typeof formatHeaderRow_ === 'function') {
      formatHeaderRow_(sheet, '#059669'); // สีเขียว Emerald สดใส
    }
    for (let c = 1; c <= headers.length; c++) {
      sheet.autoResizeColumn(c);
    }
  } else {
    // ตรวจสอบและเติมหัวตารางหากคอลัมน์ไม่ครบ
    const lastCol = sheet.getLastColumn();
    if (lastCol < headers.length) {
      for (let i = lastCol; i < headers.length; i++) {
        sheet.getRange(1, i + 1).setValue(headers[i]);
      }
      if (typeof formatHeaderRow_ === 'function') {
        formatHeaderRow_(sheet, '#059669');
      }
    }
  }

  return sheet;
}

/**
 * Helper จัดฟอร์แมตวันที่เป็น YYYY-MM-DD อย่างปลอดภัย
 */
function formatDateIso_(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, 'Asia/Bangkok', 'yyyy-MM-dd');
  }
  const str = String(val).trim();
  // ตรวจจับกรณีเป็น DD/MM/YYYY หรือ YYYY-MM-DD
  const parts = str.split(/[\/\-]/);
  if (parts.length === 3) {
    if (parts[0].length === 4) {
      return `${parts[0]}-${String(parts[1]).padStart(2, '0')}-${String(parts[2]).padStart(2, '0')}`;
    } else if (parts[2].length === 4) {
      return `${parts[2]}-${String(parts[1]).padStart(2, '0')}-${String(parts[0]).padStart(2, '0')}`;
    }
  }
  return str.substring(0, 10);
}

/**
 * Helper แปลงวัน YYYY-MM-DD เป็นข้อความไทยสวยงาม (เช่น 10 พฤศจิกายน 2569)
 */
function formatThaiDateLong_(isoDateStr) {
  if (!isoDateStr) return '';
  const parts = String(isoDateStr).trim().split('-');
  if (parts.length !== 3) return isoDateStr;
  const y = parseInt(parts[0], 10);
  const m = parseInt(parts[1], 10) - 1;
  const d = parseInt(parts[2], 10);
  const thaiMonths = [
    'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
    'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
  ];
  const thaiYear = y > 2400 ? y : y + 543;
  return `${d} ${thaiMonths[m] || ''} ${thaiYear}`;
}

/**
 * ตรวจสอบสิทธิ์ว่าผู้ใช้งานเป็น Admin หรือ Supervisor หรือไม่
 */
function checkIsAdminOrSupervisor_(username) {
  if (!username) return false;
  try {
    const ss = getDb_();
    const sheet = ss.getSheetByName(SHEET_USERS);
    if (!sheet) return false;
    const data = sheet.getDataRange().getValues();
    const target = String(username).trim().toLowerCase();
    for (let i = 1; i < data.length; i++) {
      const u = String(data[i][0] || '').trim().toLowerCase();
      if (u === target) {
        const role = String(data[i][3] || '').trim();
        return role === 'Admin' || role === 'Supervisor';
      }
    }
    return false;
  } catch (e) {
    console.error('checkIsAdminOrSupervisor_ error: ' + e.message);
    return false;
  }
}

/**
 * ดึงรายการกำหนดการเข้าไซต์งาน (รองรับฟิลเตอร์ ปี, เดือน, ไซต์งาน, ผู้ปฏิบัติงาน)
 */
function getSiteSchedules(year, month, siteFilter, userFilter) {
  try {
    const ss = getDb_();
    let sheet = ss.getSheetByName(SHEET_SITE_SCHEDULES);
    if (!sheet) {
      sheet = setupSiteSchedulesDatabase_(ss);
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { success: true, schedules: [] };
    }

    const schedules = [];
    const targetPrefix = (year && month) ? `${year}-${String(month).padStart(2, '0')}` : '';

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0] && !row[1]) continue;

      const dateIso = formatDateIso_(row[1]);
      const site = String(row[2] || '').trim();
      const username = String(row[3] || '').trim();

      // ฟิลเตอร์เดือน-ปี (ถ้ามีการระบุ)
      if (targetPrefix && !dateIso.startsWith(targetPrefix)) {
        continue;
      }

      // ฟิลเตอร์ไซต์งาน
      if (siteFilter && siteFilter !== 'ALL' && site !== siteFilter) {
        continue;
      }

      // ฟิลเตอร์ผู้ปฏิบัติงาน
      if (userFilter && userFilter !== 'ALL' && username.toLowerCase() !== userFilter.toLowerCase()) {
        continue;
      }

      schedules.push({
        id: String(row[0] || '').trim(),
        date: dateIso,
        site: site,
        username: username,
        fullName: String(row[4] || '').trim(),
        purpose: String(row[5] || '').trim(),
        taskId: String(row[6] || '').trim(),
        createdBy: String(row[7] || '').trim(),
        createdAt: row[8] ? String(row[8]) : '',
        reminderSent: Boolean(row[9] === true || String(row[9]).toUpperCase() === 'TRUE')
      });
    }

    // เรียงตามวันที่ (Ascending)
    schedules.sort((a, b) => a.date.localeCompare(b.date));

    return { success: true, schedules: schedules };
  } catch (error) {
    console.error('getSiteSchedules error: ' + error.message);
    return { success: false, message: error.message, schedules: [] };
  }
}

/**
 * ดึงข้อมูลตั้งต้นแบบครบวงจรสำหรับระบบปฏิทินเข้าไซต์งาน (Bootstrap Data)
 * คืนค่าทั้ง: schedules (ตารางงาน), sites (ไซต์งานทั้งหมด), tasks (งานที่ยังไม่เสร็จสำหรับเชื่อมโยง),
 * users (ผู้ใช้ที่ Active), และ settings (การตั้งค่าปฏิทิน) ในการเรียกครั้งเดียว
 */
function getCalendarBootstrapData(year, month, siteFilter, userFilter, operatorUsername) {
  try {
    const ss = getDb_();

    // 1. ดึงตารางงานเข้าไซต์ (Schedules)
    const schedulesRes = getSiteSchedules(year, month, siteFilter, userFilter);
    const schedules = (schedulesRes && schedulesRes.success) ? schedulesRes.schedules : [];

    // 2. ดึงการตั้งค่าระบบ (Settings) และรายชื่อไซต์งาน (Sites)
    let settings = {};
    if (typeof getSettings_ === 'function') {
      try { settings = getSettings_(); } catch (e) {}
    }
    const siteSet = new Set();
    if (settings && Array.isArray(settings.SITES)) {
      settings.SITES.forEach(s => { if (s && String(s).trim()) siteSet.add(String(s).trim()); });
    }
    if (typeof DEFAULT_SITES !== 'undefined' && Array.isArray(DEFAULT_SITES)) {
      DEFAULT_SITES.forEach(s => { if (s && String(s).trim()) siteSet.add(String(s).trim()); });
    }

    // 3. ดึงงานที่ยังไม่เสร็จสิ้น (Active Tasks) จากชีต Tasks
    const activeTasks = [];
    const taskSheet = ss.getSheetByName(typeof SHEET_TASKS !== 'undefined' ? SHEET_TASKS : 'Tasks');
    if (taskSheet) {
      const taskData = taskSheet.getDataRange().getValues();
      for (let i = 1; i < taskData.length; i++) {
        const row = taskData[i];
        if (!row[0]) continue;
        const status = String(row[6] || '').trim();
        const taskSite = String(row[1] || '').trim();
        if (taskSite) siteSet.add(taskSite);

        // ดึงเฉพาะงานที่ยังไม่เสร็จสิ้น (รอดำเนินการ, กำลังดำเนินการ, รอตรวจสอบ)
        if (status !== 'เสร็จสิ้น') {
          activeTasks.push({
            id: String(row[0]).trim(),
            site: taskSite,
            title: String(row[2] || '').trim(),
            description: String(row[3] || '').trim(),
            assignee: String(row[4] || '').trim(),
            priority: String(row[5] || 'ทั่วไป').trim(),
            status: status || 'รอดำเนินการ',
            dueDate: row[7] ? (typeof normalizeDateToIso_ === 'function' ? normalizeDateToIso_(row[7]) : String(row[7])) : ''
          });
        }
      }
    }

    // 4. ดึงรายชื่อผู้ใช้งานที่เปิดใช้งาน (Active Users)
    const activeUsers = [];
    const userSheet = ss.getSheetByName(typeof SHEET_USERS !== 'undefined' ? SHEET_USERS : 'Users');
    if (userSheet) {
      const uData = userSheet.getDataRange().getValues();
      for (let i = 1; i < uData.length; i++) {
        const uRow = uData[i];
        if (!uRow[0]) continue;
        const uStatus = String(uRow[4] || 'Active').trim();
        if (uStatus === 'Inactive') continue;
        activeUsers.push({
          username: String(uRow[0]).trim(),
          fullName: String(uRow[1] || uRow[0]).trim(),
          role: String(uRow[3] || 'Member').trim(),
          status: uStatus
        });
      }
    }

    const sitesList = Array.from(siteSet);

    return {
      success: true,
      schedules: schedules,
      sites: sitesList,
      tasks: activeTasks,
      users: activeUsers,
      settings: {
        CALENDAR_ENABLE_ADVANCE_REMINDER: settings.CALENDAR_ENABLE_ADVANCE_REMINDER !== false,
        CALENDAR_REMINDER_DAYS_BEFORE: settings.CALENDAR_REMINDER_DAYS_BEFORE || '1',
        CALENDAR_REMINDER_TIME: settings.CALENDAR_REMINDER_TIME || '07:00',
        CALENDAR_NOTIFY_ON_SCHEDULE: settings.CALENDAR_NOTIFY_ON_SCHEDULE !== false,
        CALENDAR_NOTIFY_SITE_GROUP: settings.CALENDAR_NOTIFY_SITE_GROUP !== false,
        CALENDAR_NOTIFY_MAIN_GROUP: settings.CALENDAR_NOTIFY_MAIN_GROUP === true,
        CALENDAR_FIRST_DAY_OF_WEEK: settings.CALENDAR_FIRST_DAY_OF_WEEK || '0',
        CALENDAR_MEMBER_CAN_SCHEDULE: settings.CALENDAR_MEMBER_CAN_SCHEDULE === true,
        CALENDAR_MAX_CHIPS_PER_CELL: settings.CALENDAR_MAX_CHIPS_PER_CELL || '3'
      }
    };
  } catch (error) {
    console.error('getCalendarBootstrapData error: ' + error.message);
    return {
      success: false,
      message: error.message,
      schedules: [],
      sites: typeof DEFAULT_SITES !== 'undefined' ? DEFAULT_SITES : [],
      tasks: [],
      users: []
    };
  }
}

/**
 * บันทึกกำหนดการเข้าไซต์งาน (เฉพาะ Admin / Supervisor)
 * @param {Object} scheduleData ข้อมูลกำหนดการ
 * @param {string} operatorUsername ผู้ทำรายการ
 */
function saveSiteSchedule(scheduleData, operatorUsername) {
  try {
    if (!scheduleData) {
      return { success: false, message: 'ไม่มีข้อมูลกำหนดการที่ส่งมา' };
    }

    // 1. ตรวจสอบสิทธิ์ (เฉพาะ Admin / Supervisor หรือ Member ที่ได้รับอนุญาตให้ลงคิวตนเอง)
    const settings = (typeof getSettings_ === 'function') ? getSettings_() : null;
    const isAdminOrSup = checkIsAdminOrSupervisor_(operatorUsername);

    // ดึงรายชื่อผู้ปฏิบัติงานที่ส่งมา (รองรับทั้ง usernames Array และ username String)
    let usernames = [];
    if (Array.isArray(scheduleData.usernames) && scheduleData.usernames.length > 0) {
      usernames = scheduleData.usernames.map(u => String(u || '').trim()).filter(Boolean);
    } else if (scheduleData.username) {
      usernames = [String(scheduleData.username).trim()];
    }

    if (usernames.length === 0) {
      return { success: false, message: 'กรุณาระบุผู้ปฏิบัติงานอย่างน้อย 1 คน' };
    }

    // หากเป็น Member ตรวจสอบว่าเลือกเฉพาะตัวเองหรือไม่
    const isMemberSelfSchedule = (settings && settings.CALENDAR_MEMBER_CAN_SCHEDULE === true && 
      usernames.length === 1 && String(operatorUsername || '').trim().toLowerCase() === usernames[0].toLowerCase());

    if (!isAdminOrSup && !isMemberSelfSchedule) {
      return { 
        success: false, 
        message: 'ท่านไม่มีสิทธิ์กำหนดหรือแก้ไขตารางเข้าไซต์งาน (จำกัดเฉพาะสิทธิ์ Admin หรือ Supervisor เท่านั้น)' 
      };
    }

    // 2. ตรวจสอบข้อมูลบังคับ
    const dateIso = formatDateIso_(scheduleData.date);
    const site = String(scheduleData.site || '').trim();

    if (!dateIso || !site) {
      return { success: false, message: 'กรุณาระบุ วันที่ และ ไซต์งาน ให้ครบถ้วน' };
    }

    const purpose = String(scheduleData.purpose || '').trim();
    const taskId = String(scheduleData.taskId || '').trim();

    const ss = getDb_();
    let sheet = ss.getSheetByName(SHEET_SITE_SCHEDULES);
    if (!sheet) {
      sheet = setupSiteSchedulesDatabase_(ss);
    }

    // ดึง FullName ของผู้ใช้งานจากชีต Users
    const usersMap = {};
    const uSheet = ss.getSheetByName(SHEET_USERS);
    if (uSheet) {
      const uData = uSheet.getDataRange().getValues();
      for (let i = 1; i < uData.length; i++) {
        const u = String(uData[i][0] || '').trim().toLowerCase();
        usersMap[u] = String(uData[i][1] || '').trim();
      }
    }

    const nowStr = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
    const savedRecords = [];
    const scheduleId = scheduleData.id ? String(scheduleData.id).trim() : '';

    if (scheduleId) {
      // โหมดแก้ไขรายการเดิม (Update single schedule)
      const data = sheet.getDataRange().getValues();
      let targetRow = -1;
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0] || '').trim() === scheduleId) {
          targetRow = i + 1;
          break;
        }
      }

      if (targetRow > 0) {
        const uName = usernames[0];
        const fName = usersMap[uName.toLowerCase()] || uName;

        sheet.getRange(targetRow, 2).setValue(dateIso);
        sheet.getRange(targetRow, 3).setValue(site);
        sheet.getRange(targetRow, 4).setValue(uName);
        sheet.getRange(targetRow, 5).setValue(fName);
        sheet.getRange(targetRow, 6).setValue(purpose);
        sheet.getRange(targetRow, 7).setValue(taskId);
        sheet.getRange(targetRow, 8).setValue(operatorUsername || 'admin');
        sheet.getRange(targetRow, 9).setValue(nowStr);
        sheet.getRange(targetRow, 10).setValue(false);

        savedRecords.push({
          id: scheduleId,
          date: dateIso,
          site: site,
          username: uName,
          fullName: fName,
          purpose: purpose,
          taskId: taskId,
          createdBy: operatorUsername || 'admin'
        });

        // หากมีการเลือกช่างเพิ่มในขณะแก้ไข ให้บันทึกเพิ่มเป็นแถวใหม่
        for (let j = 1; j < usernames.length; j++) {
          const addU = usernames[j];
          const addF = usersMap[addU.toLowerCase()] || addU;
          const addId = `SCH-${dateIso.replace(/-/g, '')}-${Math.floor(1000 + Math.random() * 9000)}`;

          sheet.appendRow([
            addId,
            dateIso,
            site,
            addU,
            addF,
            purpose,
            taskId,
            operatorUsername || 'admin',
            nowStr,
            false
          ]);

          savedRecords.push({
            id: addId,
            date: dateIso,
            site: site,
            username: addU,
            fullName: addF,
            purpose: purpose,
            taskId: taskId,
            createdBy: operatorUsername || 'admin'
          });
        }
      }
    } else {
      // โหมดสร้างรายการใหม่ (รองรับหลายคนพร้อมกันในคลิกเดียว)
      for (let i = 0; i < usernames.length; i++) {
        const uName = usernames[i];
        const fName = usersMap[uName.toLowerCase()] || uName;
        const randSuffix = Math.floor(1000 + Math.random() * 9000);
        const newScheduleId = `SCH-${dateIso.replace(/-/g, '')}-${randSuffix}`;

        sheet.appendRow([
          newScheduleId,
          dateIso,
          site,
          uName,
          fName,
          purpose,
          taskId,
          operatorUsername || 'admin',
          nowStr,
          false
        ]);

        savedRecords.push({
          id: newScheduleId,
          date: dateIso,
          site: site,
          username: uName,
          fullName: fName,
          purpose: purpose,
          taskId: taskId,
          createdBy: operatorUsername || 'admin'
        });
      }
    }

    // 3. ส่งการแจ้งเตือนทันทีทาง Telegram (Immediate Alert)
    if (scheduleData.notifyTelegram !== false && savedRecords.length > 0) {
      try {
        if (savedRecords.length === 1) {
          // หากมีคนเดียว ส่งตามปกติ (ทั้งส่วนตัวและกลุ่ม)
          notifyTelegramSiteSchedule_(savedRecords[0], operatorUsername, false);
        } else {
          // หากมีหลายคน ส่งตรงส่วนตัวทีละคน (skipGroup = true)
          savedRecords.forEach(rec => {
            notifyTelegramSiteSchedule_(rec, operatorUsername, true);
          });
          // ส่งข้อความรวมสรุปรายชื่อทีมงานเข้ากลุ่มไซต์งานและกลุ่มหลักเพียง 1 ข้อความ (ไม่สแปมกลุ่ม)
          notifyTelegramCombinedSiteSchedule_(savedRecords, operatorUsername);
        }
      } catch (tgErr) {
        console.error('Telegram notification failed: ' + tgErr.message);
      }
    }

    const count = savedRecords.length;
    const msg = count > 1
      ? `ลงตารางเข้าไซต์งานเรียบร้อยแล้ว (${count} คน)`
      : 'บันทึกกำหนดการเข้าไซต์งานเรียบร้อยแล้ว';

    return { 
      success: true, 
      count: count,
      schedule: savedRecords[0], 
      schedules: savedRecords,
      message: msg 
    };
  } catch (error) {
    console.error('saveSiteSchedule error: ' + error.message);
    return { success: false, message: error.message };
  }
}

/**
 * ลบกำหนดการเข้าไซต์งาน (เฉพาะ Admin / Supervisor)
 */
function deleteSiteSchedule(scheduleId, operatorUsername) {
  try {
    if (!scheduleId) {
      return { success: false, message: 'ไม่พบรหัสรายการที่ต้องการลบ' };
    }

    const settings = (typeof getSettings_ === 'function') ? getSettings_() : null;
    const isAdminOrSup = checkIsAdminOrSupervisor_(operatorUsername);

    const ss = getDb_();
    const sheet = ss.getSheetByName(SHEET_SITE_SCHEDULES);
    if (!sheet) {
      return { success: false, message: 'ไม่พบตารางข้อมูล' };
    }

    const data = sheet.getDataRange().getValues();
    let targetRow = -1;
    let targetUsername = '';
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === String(scheduleId).trim()) {
        targetRow = i + 1;
        targetUsername = String(data[i][3] || '').trim();
        break;
      }
    }

    if (targetRow <= 0) {
      return { success: false, message: 'ไม่พบรายการที่ต้องการลบในระบบ' };
    }

    const isMemberSelf = (settings && settings.CALENDAR_MEMBER_CAN_SCHEDULE === true &&
      String(operatorUsername || '').trim().toLowerCase() === targetUsername.toLowerCase());

    if (!isAdminOrSup && !isMemberSelf) {
      return { 
        success: false, 
        message: 'ท่านไม่มีสิทธิ์ลบตารางเข้าไซต์งาน (จำกัดเฉพาะสิทธิ์ Admin หรือ Supervisor เท่านั้น)' 
      };
    }

    sheet.deleteRow(targetRow);
    return { success: true, message: 'ลบรายการนัดหมายเข้าไซต์งานเรียบร้อยแล้ว' };
  } catch (error) {
    console.error('deleteSiteSchedule error: ' + error.message);
    return { success: false, message: error.message };
  }
}

/**
 * ส่งแจ้งเตือนทาง Telegram เมื่อมีการลงตารางนัดหมายเข้าไซต์งาน (Immediate Alert)
 */
function notifyTelegramSiteSchedule_(schedule, operatorUsername, skipGroup = false) {
  try {
    const settings = (typeof getSettings_ === 'function') ? getSettings_() : null;
    if (!settings || !settings.TELEGRAM_BOT_TOKEN) return;
    if (settings.CALENDAR_NOTIFY_ON_SCHEDULE === false) return;

    // หา Chat ID ส่วนตัวของผู้ปฏิบัติงาน
    const userChatId = (typeof getAssigneeTelegramChatId_ === 'function') 
      ? (getAssigneeTelegramChatId_(schedule.username) || getAssigneeTelegramChatId_(schedule.fullName)) 
      : null;

    const thaiDate = formatThaiDateLong_(schedule.date);
    const escape = (typeof escapeHtml_ === 'function') ? escapeHtml_ : (t => String(t || ''));

    const msg = (
      `🗓️ <b>แจ้งเตือนกำหนดการเข้าไซต์งานใหม่</b>\n` +
      `━━━━━━━━━━━━━━━━━━\n` +
      `🏢 <b>ไซต์งาน:</b> ${escape(schedule.site)}\n` +
      `👷 <b>ผู้รับผิดชอบ:</b> ${escape(schedule.fullName)} (@${escape(schedule.username)})\n` +
      `📅 <b>วันที่เข้าไซต์:</b> ${thaiDate}\n` +
      (schedule.purpose ? `📝 <b>วัตถุประสงค์:</b> ${escape(schedule.purpose)}\n` : '') +
      (schedule.taskId ? `🆔 <b>รหัสงาน/Plan:</b> <code>${escape(schedule.taskId)}</code>\n` : '') +
      `📢 <b>ผู้มอบหมาย:</b> ${escape(operatorUsername || 'ผู้ดูแลระบบ')}\n` +
      `━━━━━━━━━━━━━━━━━━\n` +
      `👉 <i>บันทึกผ่านระบบ Work Tracker & Site Calendar</i>`
    );

    // 1. ส่งตรงเข้า Telegram ส่วนตัวของ User
    if (userChatId && typeof sendTelegramRaw_ === 'function') {
      try {
        sendTelegramRaw_(settings.TELEGRAM_BOT_TOKEN, userChatId, msg);
      } catch (err) {
        console.error(`Send personal Telegram schedule alert failed: ${err.message}`);
      }
    }

    // หากระบุ skipGroup ให้ข้ามการส่งเข้ากลุ่มไซต์งานและกลุ่มส่วนกลาง (ใช้กรณีมีฟังก์ชัน combined ส่งให้แทน)
    if (skipGroup) return;

    // 2. หากเปิดส่งเข้ากลุ่มไซต์งาน และไซต์นี้มี Telegram Group ให้ส่งด้วย
    if (settings.CALENDAR_NOTIFY_SITE_GROUP !== false && settings.SITE_TELEGRAM_CHAT_IDS && settings.SITE_TELEGRAM_CHAT_IDS[schedule.site]) {
      const siteChatId = String(settings.SITE_TELEGRAM_CHAT_IDS[schedule.site]).trim();
      if (siteChatId && siteChatId !== userChatId && typeof sendTelegramRaw_ === 'function') {
        try {
          sendTelegramRaw_(settings.TELEGRAM_BOT_TOKEN, siteChatId, msg);
        } catch (err) {
          console.error(`Send site group Telegram schedule alert failed: ${err.message}`);
        }
      }
    }

    // 3. หากเปิดส่งเข้ากลุ่มหลักส่วนกลางด้วยเสมอ
    if (settings.CALENDAR_NOTIFY_MAIN_GROUP === true && settings.TELEGRAM_CHAT_ID && typeof sendTelegramRaw_ === 'function') {
      const mainChatId = String(settings.TELEGRAM_CHAT_ID).trim();
      if (mainChatId && mainChatId !== userChatId) {
        try {
          sendTelegramRaw_(settings.TELEGRAM_BOT_TOKEN, mainChatId, msg);
        } catch (err) {
          console.error(`Send main group Telegram schedule alert failed: ${err.message}`);
        }
      }
    }
  } catch (e) {
    console.error('notifyTelegramSiteSchedule_ error: ' + e.message);
  }
}

/**
 * ส่งแจ้งเตือนสรุปรายชื่อทีมงานเข้ากลุ่ม Telegram (Site Group & Main Group) เมื่อมีการลงตารางพร้อมกันหลายคน
 */
function notifyTelegramCombinedSiteSchedule_(schedules, operatorUsername) {
  if (!schedules || schedules.length === 0) return;
  if (schedules.length === 1) {
    notifyTelegramSiteSchedule_(schedules[0], operatorUsername, false);
    return;
  }

  try {
    const settings = (typeof getSettings_ === 'function') ? getSettings_() : null;
    if (!settings || !settings.TELEGRAM_BOT_TOKEN) return;
    if (settings.CALENDAR_NOTIFY_ON_SCHEDULE === false) return;

    const first = schedules[0];
    const thaiDate = formatThaiDateLong_(first.date);
    const escape = (typeof escapeHtml_ === 'function') ? escapeHtml_ : (t => String(t || ''));

    const teamList = schedules.map(s => `  • <b>${escape(s.fullName || s.username)}</b> (@${escape(s.username)})`).join('\n');

    const msg = (
      `🗓️ <b>แจ้งเตือนกำหนดการเข้าไซต์งานใหม่ (ทีมงาน ${schedules.length} คน)</b>\n` +
      `━━━━━━━━━━━━━━━━━━\n` +
      `🏢 <b>ไซต์งาน:</b> ${escape(first.site)}\n` +
      `👷 <b>ผู้รับผิดชอบ:</b>\n${teamList}\n` +
      `📅 <b>วันที่เข้าไซต์:</b> ${thaiDate}\n` +
      (first.purpose ? `📝 <b>วัตถุประสงค์:</b> ${escape(first.purpose)}\n` : '') +
      (first.taskId ? `🆔 <b>รหัสงาน/Plan:</b> <code>${escape(first.taskId)}</code>\n` : '') +
      `📢 <b>ผู้มอบหมาย:</b> ${escape(operatorUsername || 'ผู้ดูแลระบบ')}\n` +
      `━━━━━━━━━━━━━━━━━━\n` +
      `👉 <i>บันทึกผ่านระบบ Work Tracker & Site Calendar</i>`
    );

    // 1. ส่งเข้ากลุ่มไซต์งาน (ถ้ามี)
    if (settings.CALENDAR_NOTIFY_SITE_GROUP !== false && settings.SITE_TELEGRAM_CHAT_IDS && settings.SITE_TELEGRAM_CHAT_IDS[first.site]) {
      const siteChatId = String(settings.SITE_TELEGRAM_CHAT_IDS[first.site]).trim();
      if (siteChatId && typeof sendTelegramRaw_ === 'function') {
        try {
          sendTelegramRaw_(settings.TELEGRAM_BOT_TOKEN, siteChatId, msg);
        } catch (err) {
          console.error(`Send site group Telegram combined schedule alert failed: ${err.message}`);
        }
      }
    }

    // 2. ส่งเข้ากลุ่มหลักส่วนกลาง (ถ้าเปิดไว้)
    if (settings.CALENDAR_NOTIFY_MAIN_GROUP === true && settings.TELEGRAM_CHAT_ID && typeof sendTelegramRaw_ === 'function') {
      const mainChatId = String(settings.TELEGRAM_CHAT_ID).trim();
      if (mainChatId && typeof sendTelegramRaw_ === 'function') {
        try {
          sendTelegramRaw_(settings.TELEGRAM_BOT_TOKEN, mainChatId, msg);
        } catch (err) {
          console.error(`Send main group Telegram combined schedule alert failed: ${err.message}`);
        }
      }
    }
  } catch (e) {
    console.error('notifyTelegramCombinedSiteSchedule_ error: ' + e.message);
  }
}

/**
 * ฟังก์ชันทำงานอัตโนมัติ: ส่งแจ้งเตือนล่วงหน้าก่อนถึงวันเข้าไซต์ (รองรับทั้งแชตส่วนตัว, กลุ่มประจำไซต์ และกลุ่มส่วนกลาง)
 * รันผ่าน Google Apps Script Time-driven Trigger หรือ Admin ทดสอบยิงทันที
 * @param {Object} [options] ตัวเลือกเพิ่มเติม { forceResend: boolean, source: 'manual'|'trigger' }
 */
function sendDailySiteScheduleReminders(options) {
  const opts = options || {};
  const forceResend = opts.forceResend === true;

  try {
    const settings = (typeof getSettings_ === 'function') ? getSettings_() : null;
    if (!settings || !settings.TELEGRAM_BOT_TOKEN) {
      console.log('Skipping reminders: No Telegram bot token configured');
      return {
        success: false,
        message: 'ยังไม่ได้ตั้งค่า Telegram Bot Token ในระบบ',
        sentCount: 0,
        foundCount: 0,
        totalForDate: 0,
        alreadySentCount: 0
      };
    }

    if (settings.CALENDAR_ENABLE_ADVANCE_REMINDER === false) {
      console.log('Advance reminder is disabled in Settings');
      return {
        success: true,
        sentCount: 0,
        foundCount: 0,
        totalForDate: 0,
        alreadySentCount: 0,
        message: 'ระบบแจ้งเตือนล่วงหน้าถูกปิดใช้งานอยู่ในการตั้งค่า'
      };
    }

    const ss = getDb_();
    const sheet = ss.getSheetByName(SHEET_SITE_SCHEDULES);
    if (!sheet) {
      return {
        success: true,
        sentCount: 0,
        foundCount: 0,
        totalForDate: 0,
        alreadySentCount: 0,
        message: 'ยังไม่มีข้อมูลตารางเข้าไซต์ในระบบ'
      };
    }

    // คำนวณวันที่เป้าหมายตามจำนวนวันเตือนล่วงหน้าที่กำหนด (Default: 1 วัน)
    let daysBefore = parseInt(settings.CALENDAR_REMINDER_DAYS_BEFORE !== undefined ? settings.CALENDAR_REMINDER_DAYS_BEFORE : '1', 10);
    if (isNaN(daysBefore) || daysBefore < 0) daysBefore = 1;

    const now = new Date();
    const targetDateObj = new Date(now.getTime() + daysBefore * 24 * 60 * 60 * 1000);
    const targetIso = Utilities.formatDate(targetDateObj, 'Asia/Bangkok', 'yyyy-MM-dd');
    const targetThai = formatThaiDateLong_(targetIso);

    let reminderHeader = '⏰ <b>เตือนความจำ: กำหนดการเข้าไซต์งานวันพรุ่งนี้!</b>';
    let personalHeader = '⏰ <b>เตือนความจำ: คุณมีกำหนดการเข้าไซต์งานในวันพรุ่งนี้!</b>';
    let dateLine = `📅 <b>กำหนดเข้าไซต์:</b> พรุ่งนี้ (${targetThai})`;
    if (daysBefore === 0) {
      reminderHeader = '⏰ <b>เตือนความจำ: กำหนดการเข้าไซต์งานวันนี้!</b>';
      personalHeader = '⏰ <b>เตือนความจำ: คุณมีกำหนดการเข้าไซต์งานในวันนี้!</b>';
      dateLine = `📅 <b>กำหนดเข้าไซต์:</b> วันนี้ (${targetThai})`;
    } else if (daysBefore > 1) {
      reminderHeader = `⏰ <b>เตือนความจำ: กำหนดการเข้าไซต์งานในอีก ${daysBefore} วัน!</b>`;
      personalHeader = `⏰ <b>เตือนความจำ: คุณมีกำหนดการเข้าไซต์งานในอีก ${daysBefore} วัน!</b>`;
      dateLine = `📅 <b>กำหนดเข้าไซต์:</b> ${targetThai} (ในอีก ${daysBefore} วัน)`;
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return {
        success: true,
        targetDate: targetIso,
        targetThai: targetThai,
        daysBefore: daysBefore,
        sentCount: 0,
        foundCount: 0,
        totalForDate: 0,
        alreadySentCount: 0,
        message: `ไม่พบรายการตารางเข้าไซต์สำหรับวันที่ ${targetThai} (${targetIso})`
      };
    }

    const pendingItems = [];
    let totalForDate = 0;
    let alreadySentCount = 0;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0] && !row[1]) continue;

      const dateIso = formatDateIso_(row[1]);
      if (dateIso !== targetIso) continue;

      totalForDate++;
      const reminderSent = Boolean(row[9] === true || String(row[9]).toUpperCase() === 'TRUE');
      if (reminderSent && !forceResend) {
        alreadySentCount++;
        continue;
      }

      const site = String(row[2] || '').trim();
      const username = String(row[3] || '').trim();
      const fullName = String(row[4] || '').trim() || username;
      const purpose = String(row[5] || '').trim();
      const taskId = String(row[6] || '').trim();

      const userChatId = (typeof getAssigneeTelegramChatId_ === 'function')
        ? (getAssigneeTelegramChatId_(username) || getAssigneeTelegramChatId_(fullName))
        : null;

      pendingItems.push({
        rowIdx: i + 1,
        site: site,
        username: username,
        fullName: fullName,
        purpose: purpose,
        taskId: taskId,
        userChatId: userChatId,
        sentOk: false
      });
    }

    if (pendingItems.length === 0) {
      const msgText = totalForDate > 0
        ? `คิวงานวันที่ ${targetThai} (${totalForDate} รายการ) ได้ถูกส่งแจ้งเตือนล่วงหน้าไปแล้วทั้งหมด`
        : `ไม่พบคิวงานเข้าไซต์สำหรับวันที่ ${targetThai} (ล่วงหน้า ${daysBefore} วัน)`;

      return {
        success: true,
        targetDate: targetIso,
        targetThai: targetThai,
        daysBefore: daysBefore,
        foundCount: 0,
        totalForDate: totalForDate,
        alreadySentCount: alreadySentCount,
        sentCount: 0,
        personalSentCount: 0,
        siteGroupSentCount: 0,
        mainGroupSentCount: 0,
        missingChatIdUsers: [],
        message: msgText
      };
    }

    const escape = (typeof escapeHtml_ === 'function') ? escapeHtml_ : (t => String(t || ''));
    let personalSentCount = 0;
    let siteGroupSentCount = 0;
    let mainGroupSentCount = 0;
    const missingChatSet = new Set();
    const errorsList = [];

    // 1. ส่งแจ้งเตือนตรงเข้าแชตส่วนตัว (Personal DM)
    pendingItems.forEach(item => {
      if (!item.userChatId) {
        missingChatSet.add(`${item.fullName} (@${item.username})`);
        return;
      }

      const personalMsg = (
        `${personalHeader}\n` +
        `━━━━━━━━━━━━━━━━━━\n` +
        `🏢 <b>ไซต์งาน:</b> ${escape(item.site)}\n` +
        `${dateLine}\n` +
        `👷 <b>ผู้รับผิดชอบ:</b> ${escape(item.fullName)}\n` +
        (item.purpose ? `📝 <b>งานที่ต้องทำ:</b> ${escape(item.purpose)}\n` : '') +
        (item.taskId ? `🆔 <b>รหัสงาน/Plan:</b> <code>${escape(item.taskId)}</code>\n` : '') +
        `━━━━━━━━━━━━━━━━━━\n` +
        `👉 <i>กรุณาเตรียมอุปกรณ์และวางแผนการเดินทางล่วงหน้า</i>`
      );

      if (typeof sendTelegramRaw_ === 'function') {
        try {
          sendTelegramRaw_(settings.TELEGRAM_BOT_TOKEN, item.userChatId, personalMsg);
          item.sentOk = true;
          personalSentCount++;
        } catch (sendErr) {
          console.error(`Personal reminder row ${item.rowIdx} failed: ${sendErr.message}`);
          errorsList.push(`DM ${item.username}: ${sendErr.message}`);
        }
      }
    });

    // จัดกลุ่มรายการตามไซต์งาน (สำหรับการส่งเข้ากลุ่มประจำไซต์ และสรุปเข้ากลุ่มส่วนกลาง)
    const siteGroupsMap = {};
    pendingItems.forEach(item => {
      const siteKey = item.site || 'ไม่ระบุไซต์งาน';
      if (!siteGroupsMap[siteKey]) siteGroupsMap[siteKey] = [];
      siteGroupsMap[siteKey].push(item);
    });

    // 2. ส่งเข้ากลุ่ม Telegram ประจำไซต์งาน (หากเปิดใช้งาน)
    if (settings.CALENDAR_NOTIFY_SITE_GROUP !== false && settings.SITE_TELEGRAM_CHAT_IDS) {
      Object.keys(siteGroupsMap).forEach(siteName => {
        const siteChatId = String(settings.SITE_TELEGRAM_CHAT_IDS[siteName] || '').trim();
        if (!siteChatId || typeof sendTelegramRaw_ !== 'function') return;

        const siteItems = siteGroupsMap[siteName];
        const workerLines = siteItems.map(it => {
          const pTxt = it.purpose ? ` — <i>${escape(it.purpose)}</i>` : '';
          return `  • <b>${escape(it.fullName)}</b> (@${escape(it.username)})${pTxt}`;
        }).join('\n');

        const siteGroupMsg = (
          `${reminderHeader}\n` +
          `━━━━━━━━━━━━━━━━━━\n` +
          `🏢 <b>ไซต์งาน:</b> ${escape(siteName)}\n` +
          `${dateLine}\n` +
          `👷 <b>ผู้ปฏิบัติงาน (${siteItems.length} คน):</b>\n${workerLines}\n` +
          `━━━━━━━━━━━━━━━━━━\n` +
          `👉 <i>แจ้งเตือนอัตโนมัติจาก Work Tracker & Site Calendar</i>`
        );

        try {
          sendTelegramRaw_(settings.TELEGRAM_BOT_TOKEN, siteChatId, siteGroupMsg);
          siteGroupSentCount++;
          siteItems.forEach(it => { it.sentOk = true; });
        } catch (sgErr) {
          console.error(`Site group reminder (${siteName}) failed: ${sgErr.message}`);
          errorsList.push(`SiteGroup ${siteName}: ${sgErr.message}`);
        }
      });
    }

    // 3. ส่งสรุปภาพรวมเข้ากลุ่มส่วนกลาง (หากเปิดใช้งานและตั้งค่า TELEGRAM_CHAT_ID)
    if (settings.CALENDAR_NOTIFY_MAIN_GROUP === true && settings.TELEGRAM_CHAT_ID && typeof sendTelegramRaw_ === 'function') {
      const mainChatId = String(settings.TELEGRAM_CHAT_ID).trim();
      if (mainChatId) {
        const siteBlocks = Object.keys(siteGroupsMap).map(siteName => {
          const siteItems = siteGroupsMap[siteName];
          const wLines = siteItems.map(it => {
            const pTxt = it.purpose ? ` (${escape(it.purpose)})` : '';
            return `   • <b>${escape(it.fullName)}</b>${pTxt}`;
          }).join('\n');
          return `🏢 <b>${escape(siteName)}</b>\n${wLines}`;
        }).join('\n\n');

        const mainGroupMsg = (
          `${reminderHeader}\n` +
          `━━━━━━━━━━━━━━━━━━\n` +
          `${dateLine}\n` +
          `📋 <b>สรุปคิวเข้าไซต์งานทั้งหมด (${pendingItems.length} คน / ${Object.keys(siteGroupsMap).length} ไซต์):</b>\n\n` +
          `${siteBlocks}\n` +
          `━━━━━━━━━━━━━━━━━━\n` +
          `👉 <i>แจ้งเตือนอัตโนมัติจาก Work Tracker & Site Calendar</i>`
        );

        try {
          sendTelegramRaw_(settings.TELEGRAM_BOT_TOKEN, mainChatId, mainGroupMsg);
          mainGroupSentCount++;
          pendingItems.forEach(it => { it.sentOk = true; });
        } catch (mgErr) {
          console.error(`Main group reminder failed: ${mgErr.message}`);
          errorsList.push(`MainGroup: ${mgErr.message}`);
        }
      }
    }

    // อัปเดตสถานะ ReminderSent = true เฉพาะรายการที่ส่งออกได้อย่างน้อย 1 ช่องทาง
    let markedSentCount = 0;
    pendingItems.forEach(item => {
      if (item.sentOk) {
        sheet.getRange(item.rowIdx, 10).setValue(true);
        markedSentCount++;
      }
    });

    const summaryParts = [];
    if (personalSentCount > 0) summaryParts.push(`แชตส่วนตัว ${personalSentCount} คน`);
    if (siteGroupSentCount > 0) summaryParts.push(`กลุ่มไซต์ ${siteGroupSentCount} กลุ่ม`);
    if (mainGroupSentCount > 0) summaryParts.push(`กลุ่มส่วนกลาง ${mainGroupSentCount} ข้อความ`);

    let finalMsg = '';
    if (markedSentCount > 0) {
      finalMsg = `ส่งแจ้งเตือนสำหรับวันที่ ${targetThai} สำเร็จ ${markedSentCount}/${pendingItems.length} รายการ (${summaryParts.join(', ')})`;
    } else if (missingChatSet.size > 0 && settings.CALENDAR_NOTIFY_MAIN_GROUP !== true && !settings.SITE_TELEGRAM_CHAT_IDS) {
      finalMsg = `พบ ${pendingItems.length} รายการสำหรับวันที่ ${targetThai} แต่ไม่สามารส่งได้เนื่องจากผู้ปฏิบัติงานยังไม่มี Telegram Chat ID ส่วนตัว และไม่ได้เปิดส่งเข้ากลุ่มส่วนกลาง`;
    } else {
      finalMsg = `พบ ${pendingItems.length} รายการสำหรับวันที่ ${targetThai} แต่ไม่สามารถส่งแจ้งเตือนได้${errorsList.length ? ' (' + errorsList[0] + ')' : ''}`;
    }

    console.log(`✅ Daily site schedule reminders: ${finalMsg}`);
    return {
      success: true,
      targetDate: targetIso,
      targetThai: targetThai,
      daysBefore: daysBefore,
      foundCount: pendingItems.length,
      totalForDate: totalForDate,
      alreadySentCount: alreadySentCount,
      sentCount: markedSentCount,
      personalSentCount: personalSentCount,
      siteGroupSentCount: siteGroupSentCount,
      mainGroupSentCount: mainGroupSentCount,
      missingChatIdUsers: Array.from(missingChatSet),
      errors: errorsList,
      message: finalMsg
    };
  } catch (error) {
    console.error('sendDailySiteScheduleReminders error: ' + error.message);
    return { success: false, message: error.message, sentCount: 0, foundCount: 0 };
  }
}

/**
 * Dispatcher แบบเบาที่ทำงานทุก 5 นาทีเพื่อตรวจสอบและส่งแจ้งเตือนตรงเวลาประเทศไทย (HH:mm) เป๊ะทุกนาที
 */
function checkAndRunSiteCalendarReminders() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;

  try {
    const props = PropertiesService.getScriptProperties();
    let enabledStr = props.getProperty('CAL_REMINDER_ENABLED');
    let timeStr = props.getProperty('CAL_REMINDER_TIME');

    // หาก PropertiesService ยังว่าง ให้ดึงจากชีต Settings มาซิงค์ไว้
    if (enabledStr === null || !timeStr) {
      const settings = (typeof getSettings_ === 'function') ? getSettings_() : null;
      if (settings) {
        enabledStr = String(settings.CALENDAR_ENABLE_ADVANCE_REMINDER !== false);
        timeStr = (typeof normalizeThailandTime24h_ === 'function')
          ? normalizeThailandTime24h_(settings.CALENDAR_REMINDER_TIME)
          : (settings.CALENDAR_REMINDER_TIME || '07:00');
        props.setProperty('CAL_REMINDER_ENABLED', enabledStr);
        props.setProperty('CAL_REMINDER_TIME', timeStr);
        props.setProperty('CAL_REMINDER_DAYS_BEFORE', String(settings.CALENDAR_REMINDER_DAYS_BEFORE || '1'));
      }
    }

    if (enabledStr === 'false') return;
    if (!timeStr || !timeStr.includes(':')) timeStr = '07:00';

    const tParts = timeStr.split(':');
    const targetH = parseInt(tParts[0], 10);
    const targetM = parseInt(tParts[1], 10);
    if (isNaN(targetH) || isNaN(targetM)) return;
    const targetTotalMins = targetH * 60 + targetM;

    const now = new Date();
    const todayIso = Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM-dd');
    const nowTimeStr = Utilities.formatDate(now, 'Asia/Bangkok', 'HH:mm');
    const nParts = nowTimeStr.split(':');
    const nowTotalMins = parseInt(nParts[0], 10) * 60 + parseInt(nParts[1], 10);

    // ตรวจสอบว่าเข้าสู่ช่วงเวลาที่กำหนด (+0 ถึง +90 นาที เพื่อ Catch-up กรณี Trigger ดีเลย์เล็กน้อย)
    if (nowTotalMins < targetTotalMins || nowTotalMins > targetTotalMins + 90) {
      return;
    }

    const slotKey = `${todayIso}@${timeStr}`;
    const lastSlot = props.getProperty('CAL_LAST_REMINDER_SLOT');
    if (lastSlot === slotKey) {
      return; // ส่งแล้วสำหรับ slot เวลานี้ของวันนี้
    }

    // บันทึก slot ก่อนยิงเพื่อป้องกันซ้ำซ้อน
    const nowFullStr = Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
    props.setProperty('CAL_LAST_REMINDER_SLOT', slotKey);
    props.setProperty('CAL_LAST_REMINDER_DATE', todayIso);
    props.setProperty('CAL_LAST_REMINDER_AT', nowFullStr);

    const res = sendDailySiteScheduleReminders({ source: 'trigger' });
    props.setProperty('CAL_LAST_REMINDER_RESULT', (res && res.message) ? res.message : 'ดำเนินการแล้ว');
  } catch (err) {
    console.error('checkAndRunSiteCalendarReminders error: ' + err.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * ตรวจสอบว่าข้อความ Error เกี่ยวข้องกับสิทธิ์ ScriptApp OAuth หรือไม่
 */
function isScriptAppAuthError_(errMsg) {
  const msg = String(errMsg || '').toLowerCase();
  return msg.includes('script.scriptapp') ||
    msg.includes('getprojecttriggers') ||
    msg.includes('newtrigger') ||
    msg.includes('ไม่ได้รับอนุญาต') ||
    msg.includes('permission') ||
    msg.includes('authorization');
}

/**
 * สร้าง Response มาตรฐานเมื่อตรวจสอบหรือติดตั้ง Time Trigger ไม่สำเร็จ
 */
function buildTriggerErrorResponse_(error) {
  const errMsg = (error && error.message) ? error.message : String(error || 'Unknown error');
  const needsAuth = isScriptAppAuthError_(errMsg);
  return {
    success: false,
    installed: false,
    triggerCount: 0,
    needsAuth: needsAuth,
    authError: errMsg,
    scriptEditorUrl: (typeof SCRIPT_EDITOR_URL !== 'undefined')
      ? SCRIPT_EDITOR_URL
      : 'https://script.google.com/d/1cd4_bnt_nnuKTQzu3F6gDmMpwWoeBBcuNyqPCn3lDBsAfmguacW1aUBE/edit',
    message: errMsg
  };
}

/**
 * ตรวจสอบสถานะการติดตั้งและประวัติการทำงานของ Time Trigger ปฏิทินเข้าไซต์งาน
 */
function getCalendarTriggerStatus() {
  try {
    const props = PropertiesService.getScriptProperties();
    const settings = (typeof getSettings_ === 'function') ? getSettings_() : {};
    const configuredTime = settings.CALENDAR_REMINDER_TIME || props.getProperty('CAL_REMINDER_TIME') || '07:00';
    const enabled = settings.CALENDAR_ENABLE_ADVANCE_REMINDER !== false;

    const triggers = ScriptApp.getProjectTriggers();
    const matchingTriggers = triggers.filter(t => {
      const fn = t.getHandlerFunction();
      return fn === 'checkAndRunSiteCalendarReminders' || fn === 'sendDailySiteScheduleReminders';
    });

    const hasExactDispatcher = matchingTriggers.some(t => t.getHandlerFunction() === 'checkAndRunSiteCalendarReminders');

    return {
      success: true,
      installed: matchingTriggers.length > 0,
      isExactMode: hasExactDispatcher,
      triggerCount: matchingTriggers.length,
      needsAuth: false,
      configuredTime: configuredTime,
      enabled: enabled,
      lastRunAt: props.getProperty('CAL_LAST_REMINDER_AT') || '',
      lastRunResult: props.getProperty('CAL_LAST_REMINDER_RESULT') || ''
    };
  } catch (error) {
    console.error('getCalendarTriggerStatus error: ' + error.message);
    return buildTriggerErrorResponse_(error);
  }
}

/**
 * ติดตั้งหรือรีเซ็ต Trigger อัตโนมัติแบบละเอียดทุก 5 นาที (รองรับตรงเวลาทุกนาที HH:mm)
 */
function setupScheduleReminderTrigger() {
  try {
    const settings = (typeof getSettings_ === 'function') ? getSettings_() : null;
    const rawTime = (settings && settings.CALENDAR_REMINDER_TIME) ? settings.CALENDAR_REMINDER_TIME : '07:00';
    const timeStr = (typeof normalizeThailandTime24h_ === 'function') ? normalizeThailandTime24h_(rawTime) : '07:00';

    // ซิงค์ค่าลง ScriptProperties
    const props = PropertiesService.getScriptProperties();
    props.setProperty('CAL_REMINDER_ENABLED', String(!settings || settings.CALENDAR_ENABLE_ADVANCE_REMINDER !== false));
    props.setProperty('CAL_REMINDER_TIME', timeStr);
    props.setProperty('CAL_REMINDER_DAYS_BEFORE', String((settings && settings.CALENDAR_REMINDER_DAYS_BEFORE !== undefined) ? settings.CALENDAR_REMINDER_DAYS_BEFORE : '1'));

    const triggers = ScriptApp.getProjectTriggers();
    // ลบ trigger เก่าทั้งแบบเดิมและแบบใหม่เพื่อป้องกันซ้ำซ้อน
    triggers.forEach(t => {
      const fn = t.getHandlerFunction();
      if (fn === 'sendDailySiteScheduleReminders' || fn === 'checkAndRunSiteCalendarReminders') {
        ScriptApp.deleteTrigger(t);
      }
    });

    // ติดตั้ง Dispatcher ทุก 5 นาที เพื่อให้ยิงเตือนได้ตรงเวลาทุกชั่วโมงและนาทีตามที่ตั้งค่า
    ScriptApp.newTrigger('checkAndRunSiteCalendarReminders')
      .timeBased()
      .everyMinutes(5)
      .create();

    const triggerStatus = getCalendarTriggerStatus();
    const formattedTime = timeStr + ' น. (เวลาประเทศไทย 24 ชม.)';
    return {
      success: true,
      installed: Boolean(triggerStatus && triggerStatus.installed),
      triggerCount: triggerStatus ? triggerStatus.triggerCount : 1,
      needsAuth: false,
      triggerStatus: triggerStatus,
      message: `ติดตั้งระบบแจ้งเตือนอัตโนมัติเรียบร้อย (ตรวจสอบทุก 5 นาที พร้อมส่งเวลา ${formattedTime} ทุกวัน)`
    };
  } catch (error) {
    console.error('setupScheduleReminderTrigger error: ' + error.message);
    return buildTriggerErrorResponse_(error);
  }
}

/**
 * ฟังก์ชันสำหรับ Admin ทดสอบยิงแจ้งเตือนล่วงหน้ารอบเช้าทันที (ไม่ต้องรอเวลา Trigger)
 * @param {string} operatorUsername ผู้ทำรายการ
 * @param {Object} [options] เช่น { forceResend: true }
 */
function triggerDailySiteScheduleRemindersNow(operatorUsername, options) {
  const isAdminOrSup = checkIsAdminOrSupervisor_(operatorUsername);
  if (!isAdminOrSup) {
    return { success: false, message: 'เฉพาะ Admin หรือ Supervisor เท่านั้นที่สามารถทดสอบยิงได้' };
  }
  const opts = Object.assign({}, options || {}, { source: 'manual' });
  return sendDailySiteScheduleReminders(opts);
}
