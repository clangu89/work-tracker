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
    const isMemberSelfSchedule = (settings && settings.CALENDAR_MEMBER_CAN_SCHEDULE === true && 
      String(operatorUsername || '').trim().toLowerCase() === String(scheduleData.username || '').trim().toLowerCase());

    if (!isAdminOrSup && !isMemberSelfSchedule) {
      return { 
        success: false, 
        message: 'ท่านไม่มีสิทธิ์กำหนดหรือแก้ไขตารางเข้าไซต์งาน (จำกัดเฉพาะสิทธิ์ Admin หรือ Supervisor เท่านั้น)' 
      };
    }

    // 2. ตรวจสอบข้อมูลบังคับ
    const dateIso = formatDateIso_(scheduleData.date);
    const site = String(scheduleData.site || '').trim();
    const username = String(scheduleData.username || '').trim();

    if (!dateIso || !site || !username) {
      return { success: false, message: 'กรุณาระบุ วันที่, ไซต์งาน, และผู้ปฏิบัติงานให้ครบถ้วน' };
    }

    const purpose = String(scheduleData.purpose || '').trim();
    const taskId = String(scheduleData.taskId || '').trim();

    // ค้นหา FullName ของผู้ปฏิบัติงาน
    let fullName = String(scheduleData.fullName || '').trim();
    const ss = getDb_();
    if (!fullName) {
      const uSheet = ss.getSheetByName(SHEET_USERS);
      if (uSheet) {
        const uData = uSheet.getDataRange().getValues();
        for (let i = 1; i < uData.length; i++) {
          if (String(uData[i][0] || '').trim().toLowerCase() === username.toLowerCase()) {
            fullName = String(uData[i][1] || '').trim();
            break;
          }
        }
      }
    }
    if (!fullName) fullName = username;

    let sheet = ss.getSheetByName(SHEET_SITE_SCHEDULES);
    if (!sheet) {
      sheet = setupSiteSchedulesDatabase_(ss);
    }

    const data = sheet.getDataRange().getValues();
    let targetRow = -1;
    let scheduleId = scheduleData.id ? String(scheduleData.id).trim() : '';

    if (scheduleId) {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0] || '').trim() === scheduleId) {
          targetRow = i + 1;
          break;
        }
      }
    }

    const nowStr = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');

    if (targetRow > 0) {
      // อัปเดตรายการเดิม
      sheet.getRange(targetRow, 2).setValue(dateIso);
      sheet.getRange(targetRow, 3).setValue(site);
      sheet.getRange(targetRow, 4).setValue(username);
      sheet.getRange(targetRow, 5).setValue(fullName);
      sheet.getRange(targetRow, 6).setValue(purpose);
      sheet.getRange(targetRow, 7).setValue(taskId);
      sheet.getRange(targetRow, 8).setValue(operatorUsername || 'admin');
      sheet.getRange(targetRow, 9).setValue(nowStr);
      // รีเซ็ต ReminderSent เป็น FALSE หากมีการเปลี่ยนวัน
      sheet.getRange(targetRow, 10).setValue(false);
    } else {
      // เพิ่มรายการใหม่
      const randSuffix = Math.floor(1000 + Math.random() * 9000);
      scheduleId = `SCH-${dateIso.replace(/-/g, '')}-${randSuffix}`;
      sheet.appendRow([
        scheduleId,
        dateIso,
        site,
        username,
        fullName,
        purpose,
        taskId,
        operatorUsername || 'admin',
        nowStr,
        false
      ]);
    }

    const savedRecord = {
      id: scheduleId,
      date: dateIso,
      site: site,
      username: username,
      fullName: fullName,
      purpose: purpose,
      taskId: taskId,
      createdBy: operatorUsername || 'admin'
    };

    // 3. ส่งการแจ้งเตือนทันทีทาง Telegram (Immediate Alert)
    if (scheduleData.notifyTelegram !== false) {
      try {
        notifyTelegramSiteSchedule_(savedRecord, operatorUsername);
      } catch (tgErr) {
        console.error('Telegram notification failed: ' + tgErr.message);
      }
    }

    return { 
      success: true, 
      schedule: savedRecord, 
      message: 'บันทึกกำหนดการเข้าไซต์งานเรียบร้อยแล้ว' 
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
function notifyTelegramSiteSchedule_(schedule, operatorUsername) {
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
 * ฟังก์ชันทำงานอัตโนมัติรอบเช้า (07:00 น.): ส่งแจ้งเตือนล่วงหน้า 1 วัน ก่อนถึงวันเข้าไซต์
 * รันผ่าน Google Apps Script Time-driven Trigger
 */
function sendDailySiteScheduleReminders() {
  try {
    const settings = (typeof getSettings_ === 'function') ? getSettings_() : null;
    if (!settings || !settings.TELEGRAM_BOT_TOKEN) {
      console.log('Skipping reminders: No Telegram bot token configured');
      return { success: false, message: 'No Bot Token' };
    }

    if (settings.CALENDAR_ENABLE_ADVANCE_REMINDER === false) {
      console.log('Advance reminder is disabled in Settings');
      return { success: true, count: 0, message: 'การแจ้งเตือนล่วงหน้าถูกปิดใช้งานในการตั้งค่า' };
    }

    const ss = getDb_();
    const sheet = ss.getSheetByName(SHEET_SITE_SCHEDULES);
    if (!sheet) return { success: true, count: 0 };

    // คำนวณวันที่เป้าหมายตามจำนวนวันเตือนล่วงหน้าที่กำหนด (Default: 1 วัน)
    const daysBefore = parseInt(settings.CALENDAR_REMINDER_DAYS_BEFORE !== undefined ? settings.CALENDAR_REMINDER_DAYS_BEFORE : '1', 10);
    const now = new Date();
    const targetDateObj = new Date(now.getTime() + daysBefore * 24 * 60 * 60 * 1000);
    const targetIso = Utilities.formatDate(targetDateObj, 'Asia/Bangkok', 'yyyy-MM-dd');
    const targetThai = formatThaiDateLong_(targetIso);

    let reminderHeader = '⏰ <b>เตือนความจำ: คุณมีกำหนดการเข้าไซต์งานในวันพรุ่งนี้!</b>';
    let dateLine = `📅 <b>กำหนดเข้าไซต์:</b> พรุ่งนี้ (${targetThai})`;
    if (daysBefore === 0) {
      reminderHeader = '⏰ <b>เตือนความจำ: คุณมีกำหนดการเข้าไซต์งานในวันนี้!</b>';
      dateLine = `📅 <b>กำหนดเข้าไซต์:</b> วันนี้ (${targetThai})`;
    } else if (daysBefore > 1) {
      reminderHeader = `⏰ <b>เตือนความจำ: คุณมีกำหนดการเข้าไซต์งานในอีก ${daysBefore} วัน!</b>`;
      dateLine = `📅 <b>กำหนดเข้าไซต์:</b> ${targetThai} (ในอีก ${daysBefore} วัน)`;
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, count: 0 };

    let sentCount = 0;
    const escape = (typeof escapeHtml_ === 'function') ? escapeHtml_ : (t => String(t || ''));

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const dateIso = formatDateIso_(row[1]);
      const reminderSent = Boolean(row[9] === true || String(row[9]).toUpperCase() === 'TRUE');

      // ตรวจสอบว่าตรงกับวันที่เป้าหมาย และยังไม่เคยส่งเตือน
      if (dateIso === targetIso && !reminderSent) {
        const site = String(row[2] || '').trim();
        const username = String(row[3] || '').trim();
        const fullName = String(row[4] || '').trim() || username;
        const purpose = String(row[5] || '').trim();

        const userChatId = (typeof getAssigneeTelegramChatId_ === 'function')
          ? (getAssigneeTelegramChatId_(username) || getAssigneeTelegramChatId_(fullName))
          : null;

        if (userChatId && typeof sendTelegramRaw_ === 'function') {
          const reminderMsg = (
            `${reminderHeader}\n` +
            `━━━━━━━━━━━━━━━━━━\n` +
            `🏢 <b>ไซต์งาน:</b> ${escape(site)}\n` +
            `${dateLine}\n` +
            `👷 <b>ผู้รับผิดชอบ:</b> ${escape(fullName)}\n` +
            (purpose ? `📝 <b>งานที่ต้องทำ:</b> ${escape(purpose)}\n` : '') +
            `━━━━━━━━━━━━━━━━━━\n` +
            `👉 <i>กรุณาเตรียมอุปกรณ์และวางแผนการเดินทางล่วงหน้า</i>`
          );

          try {
            sendTelegramRaw_(settings.TELEGRAM_BOT_TOKEN, userChatId, reminderMsg);
            sheet.getRange(i + 1, 10).setValue(true);
            sentCount++;
          } catch (sendErr) {
            console.error(`Send reminder for row ${i + 1} failed: ${sendErr.message}`);
          }
        }
      }
    }

    console.log(`✅ Daily site schedule reminders processed: sent ${sentCount} notices for ${targetIso}`);
    return { success: true, sentCount: sentCount, targetDate: targetIso };
  } catch (error) {
    console.error('sendDailySiteScheduleReminders error: ' + error.message);
    return { success: false, message: error.message };
  }
}

/**
 * ติดตั้งหรือรีเซ็ต Trigger อัตโนมัติสำหรับส่งแจ้งเตือนรอบเช้าตามเวลาที่ตั้งค่าไว้
 */
function setupScheduleReminderTrigger() {
  try {
    const settings = (typeof getSettings_ === 'function') ? getSettings_() : null;
    const rawTime = (settings && settings.CALENDAR_REMINDER_TIME) ? settings.CALENDAR_REMINDER_TIME : '07:00';
    const timeStr = (typeof normalizeThailandTime24h_ === 'function') ? normalizeThailandTime24h_(rawTime) : '07:00';

    let targetHour = 7;
    let targetMinute = 0;
    if (timeStr && timeStr.includes(':')) {
      const parts = timeStr.split(':');
      const h = parseInt(parts[0], 10);
      const m = parseInt(parts[1], 10);
      if (!isNaN(h) && h >= 0 && h <= 23) targetHour = h;
      if (!isNaN(m) && m >= 0 && m <= 59) targetMinute = m;
    }

    const triggers = ScriptApp.getProjectTriggers();
    // ลบ trigger เดิมเพื่อรีเซ็ตชั่วโมงใหม่ให้ตรงกับการตั้งค่า
    triggers.forEach(t => {
      if (t.getHandlerFunction() === 'sendDailySiteScheduleReminders') {
        ScriptApp.deleteTrigger(t);
      }
    });

    ScriptApp.newTrigger('sendDailySiteScheduleReminders')
      .timeBased()
      .atHour(targetHour)
      .everyDays(1)
      .inTimezone('Asia/Bangkok')
      .create();

    const formattedTime = timeStr + ' น. (เวลาประเทศไทย 24 ชม.)';
    return { 
      success: true, 
      message: `ติดตั้งระบบแจ้งเตือนอัตโนมัติรอบเช้าเวลา ${formattedTime} ทุกวัน เรียบร้อยแล้ว` 
    };
  } catch (error) {
    console.error('setupScheduleReminderTrigger error: ' + error.message);
    return { success: false, message: error.message };
  }
}

/**
 * ฟังก์ชันสำหรับ Admin ทดสอบยิงแจ้งเตือนล่วงหน้ารอบเช้าทันที (ไม่ต้องรอ 07:00 น.)
 */
function triggerDailySiteScheduleRemindersNow(operatorUsername) {
  const isAdminOrSup = checkIsAdminOrSupervisor_(operatorUsername);
  if (!isAdminOrSup) {
    return { success: false, message: 'เฉพาะ Admin หรือ Supervisor เท่านั้นที่สามารถทดสอบยิงได้' };
  }
  return sendDailySiteScheduleReminders();
}
