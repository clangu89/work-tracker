/**
 * =========================================================================
 * RDP MANAGER & ADMIN BACKEND MODULE (RdpBackend.js)
 * Google Apps Script Backend for Multi-Machine Remote Desktop & Admin Center
 * =========================================================================
 */

const SHEET_RDP_SITES = 'RDP_Sites';
const SHEET_RDP_GROUPS = 'RDP_Groups';
const SHEET_RDP_LOGS = 'RDP_AuditLogs';
const LEGACY_RDP_SPREADSHEET_ID = '1r-mSATonJBdsFUqGzPlOi6oDzp87XbeI7YFd3jP7VCk';

/**
 * คำนวณ SHA-256 Hash สำหรับรหัสผ่าน (Hex format)
 */
function hashPassword_(password) {
  if (!password) return '';
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(password), Utilities.Charset.UTF_8);
  let txtHash = '';
  for (let i = 0; i < rawHash.length; i++) {
    let byteVal = rawHash[i];
    if (byteVal < 0) byteVal += 256;
    let byteHex = byteVal.toString(16);
    if (byteHex.length === 1) byteHex = '0' + byteHex;
    txtHash += byteHex;
  }
  return txtHash;
}

/**
 * สร้างและจัดฟอร์แมตโครงสร้างชีตสำหรับระบบ RDP Manager (RDP_Sites, RDP_Groups, RDP_AuditLogs)
 */
function setupRdpDatabase_(ss) {
  if (!ss) ss = getDb_();

  // --- 1. RDP_Sites ---
  let sitesSheet = ss.getSheetByName(SHEET_RDP_SITES);
  if (!sitesSheet) {
    sitesSheet = ss.insertSheet(SHEET_RDP_SITES);
    const headers = [
      'SiteId',
      'SiteName',
      'IpAddress',
      'Port',
      'Username',
      'Password',
      'GroupName',
      'Description',
      'UpdatedAt',
      'MachinesJson'
    ];
    sitesSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    formatHeaderRow_(sitesSheet, '#1e293b');
  }

  // --- 2. RDP_Groups ---
  let groupsSheet = ss.getSheetByName(SHEET_RDP_GROUPS);
  if (!groupsSheet) {
    groupsSheet = ss.insertSheet(SHEET_RDP_GROUPS);
    const gHeaders = ['GroupId', 'GroupName', 'Description', 'CreatedAt'];
    groupsSheet.getRange(1, 1, 1, gHeaders.length).setValues([gHeaders]);
    formatHeaderRow_(groupsSheet, '#334155');

    const defaultGroups = [
      ['GRP-001', 'Headquarters', 'สำนักงานใหญ่ / ศูนย์กลางระบบ', new Date()],
      ['GRP-002', 'Branch', 'สาขาย่อยและไซต์ภูมิภาค', new Date()],
      ['GRP-003', 'Datacenter', 'ศูนย์จัดเก็บข้อมูลและคลาวด์เซิร์ฟเวอร์', new Date()]
    ];
    groupsSheet.getRange(2, 1, defaultGroups.length, gHeaders.length).setValues(defaultGroups);
    groupsSheet.autoResizeColumn(1);
    groupsSheet.autoResizeColumn(2);
    groupsSheet.autoResizeColumn(3);
    groupsSheet.autoResizeColumn(4);
  }

  // --- 3. RDP_AuditLogs ---
  let logsSheet = ss.getSheetByName(SHEET_RDP_LOGS);
  if (!logsSheet) {
    logsSheet = ss.insertSheet(SHEET_RDP_LOGS);
    const lHeaders = ['Timestamp', 'Username', 'Action', 'TargetId', 'Details'];
    logsSheet.getRange(1, 1, 1, lHeaders.length).setValues([lHeaders]);
    formatHeaderRow_(logsSheet, '#0f172a');
  }
}

/**
 * บันทึก Audit Log ลงในตาราง RDP_AuditLogs
 */
function writeRdpAuditLog_(username, action, targetId, details) {
  try {
    const ss = getDb_();
    let sheet = ss.getSheetByName(SHEET_RDP_LOGS);
    if (!sheet) {
      setupRdpDatabase_(ss);
      sheet = ss.getSheetByName(SHEET_RDP_LOGS);
    }
    const nowStr = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
    sheet.appendRow([nowStr, String(username || '-').trim(), String(action || '-').trim(), String(targetId || '-').trim(), String(details || '').trim()]);
  } catch (err) {
    console.error('writeRdpAuditLog_ error: ' + err.message);
  }
}

/**
 * ดึงรายการ Audit Logs ล่าสุด
 */
function getRdpAuditLogs(limit = 100) {
  try {
    const ss = getDb_();
    const sheet = ss.getSheetByName(SHEET_RDP_LOGS);
    if (!sheet) return { success: true, logs: [] };

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, logs: [] };

    const logs = [];
    const maxRows = Math.min(data.length - 1, limit);
    for (let i = data.length - 1; i >= data.length - maxRows; i--) {
      const row = data[i];
      if (!row[0] && !row[1]) continue;
      logs.push({
        timestamp: row[0] ? String(row[0]) : '',
        username: String(row[1] || ''),
        action: String(row[2] || ''),
        targetId: String(row[3] || ''),
        details: String(row[4] || '')
      });
    }

    return { success: true, logs: logs };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * ดึงข้อมูลกลุ่มทั้งหมดในระบบ RDP Manager
 */
function getRdpGroups() {
  try {
    const ss = getDb_();
    let sheet = ss.getSheetByName(SHEET_RDP_GROUPS);
    if (!sheet) {
      setupRdpDatabase_(ss);
      sheet = ss.getSheetByName(SHEET_RDP_GROUPS);
    }

    const data = sheet.getDataRange().getValues();
    const groups = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0] && !data[i][1]) continue;
      groups.push({
        group_id: String(data[i][0] || '').trim(),
        group_name: String(data[i][1] || '').trim(),
        description: String(data[i][2] || '').trim(),
        created_at: data[i][3] ? String(data[i][3]) : ''
      });
    }

    return { success: true, groups: groups };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * บันทึกกลุ่มไซต์งาน (เพิ่มใหม่หรือแก้ไข)
 */
function saveRdpGroup(groupData, reqUsername) {
  try {
    if (!groupData || !groupData.group_name) {
      return { success: false, message: 'กรุณาระบุชื่อกลุ่มงาน' };
    }

    const ss = getDb_();
    let sheet = ss.getSheetByName(SHEET_RDP_GROUPS);
    if (!sheet) {
      setupRdpDatabase_(ss);
      sheet = ss.getSheetByName(SHEET_RDP_GROUPS);
    }

    const data = sheet.getDataRange().getValues();
    const gName = String(groupData.group_name).trim();
    const gDesc = String(groupData.description || '').trim();
    const gId = groupData.group_id ? String(groupData.group_id).trim() : '';

    let targetRow = -1;
    if (gId) {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === gId) {
          targetRow = i + 1;
          break;
        }
      }
    }

    const nowStr = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');

    if (targetRow > 0) {
      sheet.getRange(targetRow, 2, 1, 2).setValues([[gName, gDesc]]);
      writeRdpAuditLog_(reqUsername || 'admin', 'GROUP_UPDATE', gId, `แก้ไขกลุ่ม: ${gName}`);
      return { success: true, message: `บันทึกการแก้ไขกลุ่ม "${gName}" เรียบร้อยแล้ว` };
    } else {
      const newGroupId = 'GRP-' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd-HHmmss');
      sheet.appendRow([newGroupId, gName, gDesc, nowStr]);
      writeRdpAuditLog_(reqUsername || 'admin', 'GROUP_CREATE', newGroupId, `สร้างกลุ่มใหม่: ${gName}`);
      return { success: true, message: `เพิ่มกลุ่มงาน "${gName}" สำเร็จ`, group_id: newGroupId };
    }
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * ลบกลุ่มงาน RDP (ตรวจสอบว่ามีไซต์งานผูกอยู่หรือไม่)
 */
function deleteRdpGroup(groupId, reqUsername) {
  try {
    if (!groupId) return { success: false, message: 'ไม่พบรหัสกลุ่ม' };

    const ss = getDb_();
    const gSheet = ss.getSheetByName(SHEET_RDP_GROUPS);
    const sSheet = ss.getSheetByName(SHEET_RDP_SITES);
    if (!gSheet) return { success: false, message: 'ไม่พบตาราง Groups' };

    const gData = gSheet.getDataRange().getValues();
    let groupRow = -1;
    let groupName = '';

    for (let i = 1; i < gData.length; i++) {
      if (String(gData[i][0]).trim() === String(groupId).trim()) {
        groupRow = i + 1;
        groupName = String(gData[i][1]).trim();
        break;
      }
    }

    if (groupRow === -1) {
      return { success: false, message: 'ไม่พบกลุ่มงานนี้ในระบบ' };
    }

    // ตรวจสอบว่ามีไซต์ที่ใช้กลุ่มนี้อยู่หรือไม่
    if (sSheet) {
      const sData = sSheet.getDataRange().getValues();
      const boundSites = [];
      for (let j = 1; j < sData.length; j++) {
        const siteGroup = String(sData[j][6] || '').trim();
        if (siteGroup.toLowerCase() === groupName.toLowerCase()) {
          boundSites.push(String(sData[j][1] || sData[j][0]));
        }
      }
      if (boundSites.length > 0) {
        return {
          success: false,
          message: `ไม่สามารถลบกลุ่ม "${groupName}" ได้ เนื่องจากยังมีไซต์งานผูกอยู่ ${boundSites.length} ไซต์ (${boundSites.slice(0, 3).join(', ')}${boundSites.length > 3 ? '...' : ''}) กรุณาย้ายกลุ่มของไซต์งานก่อนลบ`
        };
      }
    }

    gSheet.deleteRow(groupRow);
    writeRdpAuditLog_(reqUsername || 'admin', 'GROUP_DELETE', groupId, `ลบกลุ่ม: ${groupName}`);
    return { success: true, message: `ลบกลุ่ม "${groupName}" เรียบร้อยแล้ว` };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * ดึงข้อมูลเริ่มต้นของ RDP Manager (ไซต์งาน, กลุ่ม, ผู้ใช้ออนไลน์, สิทธิ์)
 */
function getRdpInitialData(reqUsername) {
  try {
    const ss = getDb_();
    setupRdpDatabase_(ss);

    // ดึงโปรไฟล์และสิทธิ์ของผู้ใช้
    const userProfile = getUserProfile_(reqUsername);
    const isAdmin = userProfile && userProfile.role === 'Admin';
    let allowedRdpSites = 'ALL';

    if (userProfile && !isAdmin) {
      allowedRdpSites = userProfile.allowedRdpSites || 'ALL';
    }

    // 1. อ่านข้อมูลกลุ่ม
    const groupsRes = getRdpGroups();
    const groups = groupsRes.groups || [];

    // 2. อ่านข้อมูลไซต์งาน
    const sSheet = ss.getSheetByName(SHEET_RDP_SITES);
    const sites = [];

    if (sSheet && sSheet.getLastRow() > 1) {
      const sData = sSheet.getDataRange().getValues();

      for (let i = 1; i < sData.length; i++) {
        const row = sData[i];
        if (!row[0] && !row[1]) continue;

        const siteId = String(row[0]).trim();
        const siteName = String(row[1] || '').trim();

        // ตรวจสอบสิทธิ์การเข้าถึงไซต์
        if (!isAdmin && allowedRdpSites !== 'ALL') {
          const allowedList = Array.isArray(allowedRdpSites) ? allowedRdpSites : [allowedRdpSites];
          if (!allowedList.includes('*') && !allowedList.includes(siteId) && !allowedList.includes(siteName)) {
            continue; // ข้ามไซต์ที่ไม่มีสิทธิ์
          }
        }

        const defaultIp = String(row[2] || '').trim();
        const defaultPort = parseInt(row[3], 10) || 3389;
        const defaultUser = String(row[4] || '').trim();
        const defaultPass = String(row[5] || '');
        const groupName = String(row[6] || 'General').trim();
        const description = String(row[7] || '').trim();
        const updatedAt = row[8] ? String(row[8]) : '';

        // แยกรายการเครื่อง (Machines)
        let machines = [];
        if (row[9]) {
          try {
            const parsed = JSON.parse(String(row[9]));
            if (Array.isArray(parsed) && parsed.length > 0) {
              machines = parsed;
            }
          } catch (e) {}
        }

        if (machines.length === 0) {
          machines = [{
            machine_id: 'MCH-1',
            machine_name: 'เครื่องหลัก',
            ip_address: defaultIp,
            port: defaultPort,
            username: defaultUser,
            password: defaultPass,
            note: ''
          }];
        }

        // จัดเตรียม Machine แต่ละเครื่อง
        const processedMachines = machines.map((m, idx) => {
          const mId = m.machine_id || `MCH-${idx + 1}`;
          const mName = String(m.machine_name || `เครื่องที่ ${idx + 1}`).trim();
          const mIp = String(m.ip_address || defaultIp).trim();
          const mPort = parseInt(m.port, 10) || defaultPort;
          const mUser = String(m.username || defaultUser).trim();
          const mPass = String(m.password || '');
          const mNote = String(m.note || '').trim();

          // สร้าง remote_payload สำหรับ 1-Click Zero-Prompt Connection
          const connectionObj = {
            id: `${siteId}_${mId}`,
            site_id: siteId,
            machine_id: mId,
            name: `${siteName} - ${mName}`,
            ip: mIp,
            port: mPort,
            user: mUser,
            pass: mPass
          };

          const remotePayload = Utilities.base64Encode(
            JSON.stringify(connectionObj),
            Utilities.Charset.UTF_8
          );

          return {
            machine_id: mId,
            machine_name: mName,
            ip_address: mIp,
            port: mPort,
            username: mUser,
            password: isAdmin ? mPass : '••••••••',
            note: mNote,
            remote_payload: remotePayload
          };
        });

        const primary = processedMachines[0] || {};

        sites.push({
          site_id: siteId,
          site_name: siteName,
          ip_address: primary.ip_address || defaultIp,
          port: primary.port || defaultPort,
          username: primary.username || defaultUser,
          password: primary.password || '••••••••',
          group_name: groupName,
          description: description,
          updated_at: updatedAt,
          machines: processedMachines
        });
      }
    }

    // 3. ข้อมูลผู้ใช้ออนไลน์
    const onlineInfo = getOnlineUsers();

    return {
      success: true,
      sites: sites,
      groups: groups,
      online_users: onlineInfo.users || [],
      onlineCount: onlineInfo.count || 0,
      currentUserProfile: userProfile,
      isAdmin: isAdmin
    };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * บันทึกไซต์งาน RDP (Admin Only) พร้อมข้อมูลหลายเครื่อง (Multi-Machine)
 */
function saveRdpSite(siteData, reqUsername) {
  try {
    const userProfile = getUserProfile_(reqUsername);
    if (!userProfile || userProfile.role !== 'Admin') {
      return { success: false, message: 'ไม่มีสิทธิ์ดำเนินการ ต้องเป็น Admin เท่านั้น' };
    }

    if (!siteData || !siteData.site_name || !siteData.site_name.trim()) {
      return { success: false, message: 'กรุณาระบุชื่อไซต์งาน' };
    }

    const ss = getDb_();
    setupRdpDatabase_(ss);
    const sheet = ss.getSheetByName(SHEET_RDP_SITES);
    const data = sheet.getDataRange().getValues();

    const siteId = siteData.site_id ? String(siteData.site_id).trim() : '';
    let targetRow = -1;
    let oldMachines = [];

    if (siteId) {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === siteId) {
          targetRow = i + 1;
          if (data[i][9]) {
            try { oldMachines = JSON.parse(String(data[i][9])) || []; } catch (e) {}
          }
          if (oldMachines.length === 0) {
            oldMachines = [{
              machine_id: 'MCH-1',
              machine_name: 'เครื่องหลัก',
              ip_address: data[i][2],
              port: data[i][3],
              username: data[i][4],
              password: data[i][5],
              note: ''
            }];
          }
          break;
        }
      }
    }

    // จัดการรายการเครื่อง
    let rawMachines = Array.isArray(siteData.machines) && siteData.machines.length > 0
      ? siteData.machines
      : [{
          machine_id: 'MCH-1',
          machine_name: 'เครื่องหลัก',
          ip_address: siteData.ip_address || '',
          port: parseInt(siteData.port, 10) || 3389,
          username: siteData.username || 'Administrator',
          password: siteData.password || '',
          note: ''
        }];

    // ตรวจสอบและรักษารหัสผ่านเดิมหากไม่มีการเปลี่ยน (Password Preservation)
    const finalMachines = rawMachines.map((m, idx) => {
      const mId = m.machine_id || `MCH-${Date.now() + idx}`;
      let mPass = m.password;

      if (!mPass || mPass === '••••••••') {
        const match = oldMachines.find(om => om.machine_id === m.machine_id);
        if (match && match.password) {
          mPass = match.password;
        } else if (idx === 0 && targetRow > 0) {
          mPass = String(data[targetRow - 1][5] || '');
        } else {
          mPass = '';
        }
      }

      return {
        machine_id: mId,
        machine_name: String(m.machine_name || `เครื่องที่ ${idx + 1}`).trim(),
        ip_address: String(m.ip_address || '').trim(),
        port: parseInt(m.port, 10) || 3389,
        username: String(m.username || 'Administrator').trim(),
        password: mPass,
        note: String(m.note || '').trim()
      };
    });

    if (finalMachines.length === 0 || !finalMachines[0].ip_address) {
      return { success: false, message: 'กรุณากรอกข้อมูลอย่างน้อย 1 เครื่อง (ต้องมี IP Address และ Username)' };
    }

    const primary = finalMachines[0];
    const machinesJson = JSON.stringify(finalMachines);
    const nowStr = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
    const sName = String(siteData.site_name).trim();
    const gName = String(siteData.group_name || 'General').trim();
    const desc = String(siteData.description || '').trim();

    if (targetRow > 0) {
      // แก้ไขไซต์เดิม
      sheet.getRange(targetRow, 2, 1, 9).setValues([[
        sName,
        primary.ip_address,
        primary.port,
        primary.username,
        primary.password,
        gName,
        desc,
        nowStr,
        machinesJson
      ]]);
      writeRdpAuditLog_(reqUsername || 'admin', 'SITE_UPDATE', siteId, `แก้ไขไซต์: ${sName} (${finalMachines.length} เครื่อง)`);
      return { success: true, message: `บันทึกการแก้ไขไซต์งาน "${sName}" สำเร็จ` };
    } else {
      // เพิ่มไซต์ใหม่
      const newSiteId = 'SITE-' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd-HHmmss');
      sheet.appendRow([
        newSiteId,
        sName,
        primary.ip_address,
        primary.port,
        primary.username,
        primary.password,
        gName,
        desc,
        nowStr,
        machinesJson
      ]);
      writeRdpAuditLog_(reqUsername || 'admin', 'SITE_CREATE', newSiteId, `เพิ่มไซต์ใหม่: ${sName} (${finalMachines.length} เครื่อง)`);
      return { success: true, message: `เพิ่มไซต์งาน "${sName}" สำเร็จ`, site_id: newSiteId };
    }
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * ลบไซต์งาน RDP (Admin Only)
 */
function deleteRdpSite(siteId, reqUsername) {
  try {
    const userProfile = getUserProfile_(reqUsername);
    if (!userProfile || userProfile.role !== 'Admin') {
      return { success: false, message: 'ไม่มีสิทธิ์ดำเนินการ ต้องเป็น Admin เท่านั้น' };
    }

    if (!siteId) return { success: false, message: 'ไม่พบรหัสไซต์งาน' };

    const ss = getDb_();
    const sheet = ss.getSheetByName(SHEET_RDP_SITES);
    if (!sheet) return { success: false, message: 'ไม่พบตาราง RDP_Sites' };

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(siteId).trim()) {
        const siteName = String(data[i][1]).trim();
        sheet.deleteRow(i + 1);
        writeRdpAuditLog_(reqUsername || 'admin', 'SITE_DELETE', siteId, `ลบไซต์งาน: ${siteName}`);
        return { success: true, message: `ลบไซต์งาน "${siteName}" สำเร็จ` };
      }
    }

    return { success: false, message: 'ไม่พบไซต์งานที่ต้องการลบ' };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * บันทึกประวัติการเชื่อมต่อ Remote Desktop และอัปเดตสถานะ Presence
 */
function logRdpRemoteAction(username, siteId, siteName, machineId, machineName) {
  try {
    const sLabel = String(siteName || '').trim();
    const mLabel = String(machineName || machineId || '').trim();
    const combined = (sLabel && mLabel) ? `${sLabel}/${mLabel}` : (sLabel || mLabel || 'เครื่องรีโมต');

    writeRdpAuditLog_(username, 'REMOTE_CONNECT', siteId, `เปิดการเชื่อมต่อ Remote Desktop: ${combined}`);

    // อัปเดต Online Presence ให้แสดงสถานะ Active Remote
    const activeRemoteObj = {
      siteId: siteId,
      siteName: siteName || '',
      machineId: machineId || '',
      machineName: machineName || '',
      connectedAt: Date.now()
    };

    updateUserActiveRemote_(username, activeRemoteObj);

    return {
      success: true,
      online_users: getOnlineUsers().users || []
    };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * ดึงข้อมูลรวมศูนย์สำหรับหน้าจอระบบ Admin Center
 */
function getAdminOverviewData(reqUsername) {
  try {
    const userProfile = getUserProfile_(reqUsername);
    if (!userProfile || userProfile.role !== 'Admin') {
      return { success: false, message: 'สิทธิ์ไม่เพียงพอ ต้องเป็น Admin เท่านั้น' };
    }

    const ss = getDb_();
    setupRdpDatabase_(ss);

    // 1. ดึงรายชื่อผู้ใช้ทั้งหมด (รวมคอลัมน์ใหม่ Password, AllowedApps, AllowedRdpSites)
    const usersRes = getUsersList();
    const users = usersRes.users || [];

    // 2. ดึงการตั้งค่าทั้งหมด
    const settings = getSettings_();

    // 3. ดึงไซต์งาน Work Tracker
    const wtSites = settings.SITES || DEFAULT_SITES;

    // 4. ดึงข้อมูล RDP Sites & Groups
    const rdpSitesData = [];
    const rdpSheet = ss.getSheetByName(SHEET_RDP_SITES);
    if (rdpSheet && rdpSheet.getLastRow() > 1) {
      const sVals = rdpSheet.getDataRange().getValues();
      for (let i = 1; i < sVals.length; i++) {
        if (!sVals[i][0] && !sVals[i][1]) continue;
        let machines = [];
        if (sVals[i][9]) {
          try { machines = JSON.parse(String(sVals[i][9])); } catch (e) {}
        }
        rdpSitesData.push({
          site_id: String(sVals[i][0]).trim(),
          site_name: String(sVals[i][1]).trim(),
          ip_address: String(sVals[i][2]).trim(),
          port: sVals[i][3] || 3389,
          group_name: String(sVals[i][6] || 'General').trim(),
          machinesCount: machines.length || 1
        });
      }
    }

    const rdpGroupsRes = getRdpGroups();
    const rdpGroups = rdpGroupsRes.groups || [];

    // 5. ดึง Audit Logs ล่าสุด 50 รายการ
    const logsRes = getRdpAuditLogs(50);

    return {
      success: true,
      users: users,
      workTrackerSites: wtSites,
      rdpSites: rdpSitesData,
      rdpGroups: rdpGroups,
      settings: settings,
      auditLogs: logsRes.logs || [],
      sheetUrl: ss.getUrl()
    };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * โอนย้ายข้อมูลจาก Google Spreadsheet เดิมของ RDP Manager (1-Click Auto-Migration)
 */
function migrateFromLegacyRdpSpreadsheet(reqUsername) {
  try {
    const userProfile = getUserProfile_(reqUsername);
    if (reqUsername && (!userProfile || userProfile.role !== 'Admin')) {
      return { success: false, message: 'สิทธิ์ไม่เพียงพอ ต้องเป็น Admin เท่านั้น' };
    }

    const currentSs = getDb_();
    setupRdpDatabase_(currentSs);

    let legacySs;
    try {
      legacySs = SpreadsheetApp.openById(LEGACY_RDP_SPREADSHEET_ID);
    } catch (e) {
      return { success: false, message: `ไม่สามารถเปิดสเปรดชีต RDP เดิมได้ (${LEGACY_RDP_SPREADSHEET_ID}): ${e.message}` };
    }

    let sitesImported = 0;
    let groupsImported = 0;
    let usersUpdated = 0;

    // --- 1. โอนย้าย Groups ---
    const legacyGroupsSheet = legacySs.getSheetByName('Groups');
    const curGroupsSheet = currentSs.getSheetByName(SHEET_RDP_GROUPS);
    if (legacyGroupsSheet && curGroupsSheet) {
      const lData = legacyGroupsSheet.getDataRange().getValues();
      const cData = curGroupsSheet.getDataRange().getValues();
      const curGroupNames = new Set(cData.slice(1).map(r => String(r[1]).trim().toLowerCase()));

      for (let i = 1; i < lData.length; i++) {
        const gName = String(lData[i][1] || '').trim();
        if (gName && !curGroupNames.has(gName.toLowerCase())) {
          curGroupsSheet.appendRow([
            String(lData[i][0] || `GRP-${Date.now()}`),
            gName,
            String(lData[i][2] || ''),
            lData[i][3] ? lData[i][3] : new Date()
          ]);
          curGroupNames.add(gName.toLowerCase());
          groupsImported++;
        }
      }
    }

    // --- 2. โอนย้าย Sites ---
    const legacySitesSheet = legacySs.getSheetByName('Sites');
    const curSitesSheet = currentSs.getSheetByName(SHEET_RDP_SITES);
    if (legacySitesSheet && curSitesSheet) {
      const lData = legacySitesSheet.getDataRange().getValues();
      const cData = curSitesSheet.getDataRange().getValues();
      const curSiteNames = new Set(cData.slice(1).map(r => String(r[1]).trim().toLowerCase()));

      // ค้นหาคอลัมน์ machines_json ของ legacy sheet
      const lHeaders = lData[0].map(h => String(h).trim());
      const colMach = lHeaders.indexOf('machines_json');

      for (let i = 1; i < lData.length; i++) {
        const sName = String(lData[i][1] || '').trim();
        if (sName && !curSiteNames.has(sName.toLowerCase())) {
          const sId = String(lData[i][0] || `SITE-${Date.now()}`);
          const ip = String(lData[i][2] || '');
          const port = lData[i][3] || 3389;
          const user = String(lData[i][4] || '');
          const pass = String(lData[i][5] || '');
          const grp = String(lData[i][6] || 'General');
          const desc = String(lData[i][7] || '');
          const updated = lData[i][8] ? lData[i][8] : new Date();
          const machJson = (colMach !== -1 && lData[i][colMach]) ? String(lData[i][colMach]) : '';

          curSitesSheet.appendRow([sId, sName, ip, port, user, pass, grp, desc, updated, machJson]);
          curSiteNames.add(sName.toLowerCase());
          sitesImported++;
        }
      }
    }

    // --- 3. โอนย้าย Users & Permissions ---
    const legacyUsersSheet = legacySs.getSheetByName('Users');
    const legacyPermSheet = legacySs.getSheetByName('Permissions');
    const curUsersSheet = currentSs.getSheetByName(SHEET_USERS);

    if (legacyUsersSheet && curUsersSheet) {
      // อ่าน permissions ของ legacy
      const userPermMap = {};
      if (legacyPermSheet) {
        const pData = legacyPermSheet.getDataRange().getValues();
        for (let p = 1; p < pData.length; p++) {
          const uId = String(pData[p][1] || '').trim();
          const sId = String(pData[p][2] || '').trim();
          if (uId && sId) {
            if (!userPermMap[uId]) userPermMap[uId] = [];
            userPermMap[uId].push(sId);
          }
        }
      }

      const lUData = legacyUsersSheet.getDataRange().getValues();
      const lUHeaders = lUData[0].map(h => String(h).trim());
      const lColChatId = lUHeaders.indexOf('telegram_chat_id');

      const cUData = curUsersSheet.getDataRange().getValues();
      const curUserMap = {};
      for (let c = 1; c < cUData.length; c++) {
        const uName = String(cUData[c][0] || '').trim().toLowerCase();
        if (uName) curUserMap[uName] = c + 1; // row index in Sheet
      }

      for (let i = 1; i < lUData.length; i++) {
        const rawUId = String(lUData[i][0] || '').trim();
        const uName = String(lUData[i][1] || '').trim().toLowerCase();
        if (!uName) continue;

        const passHash = String(lUData[i][2] || '').trim();
        const dispName = String(lUData[i][3] || uName).trim();
        const role = String(lUData[i][4] || '').toLowerCase() === 'admin' ? 'Admin' : 'Member';
        const status = String(lUData[i][5] || 'Active').trim();
        const chatId = (lColChatId !== -1 && lUData[i][lColChatId]) ? String(lUData[i][lColChatId]).trim() : '';

        // ดึงสิทธิ์ไซต์ RDP
        const permList = userPermMap[rawUId] || (role === 'Admin' ? ['*'] : []);
        const rdpPermVal = (permList.includes('*') || role === 'Admin') ? 'ALL' : JSON.stringify(permList);

        if (curUserMap[uName]) {
          // มีผู้ใช้นี้อยู่แล้วใน Users -> อัปเดตคอลัมน์รหัสผ่านสำรอง และ allowedRdpSites
          const rIndex = curUserMap[uName];
          if (passHash) {
            curUsersSheet.getRange(rIndex, 8).setValue(passHash);
          }
          if (chatId && !curUsersSheet.getRange(rIndex, 3).getValue()) {
            curUsersSheet.getRange(rIndex, 3).setValue(chatId);
          }
          curUsersSheet.getRange(rIndex, 10).setValue(rdpPermVal);
          usersUpdated++;
        } else {
          // เพิ่มผู้ใช้ใหม่
          curUsersSheet.appendRow([
            uName,
            dispName,
            chatId,
            role,
            status,
            '',
            'ALL',
            passHash,
            JSON.stringify(['work_tracker', 'rdp_manager']),
            rdpPermVal
          ]);
          curUserMap[uName] = curUsersSheet.getLastRow();
          usersUpdated++;
        }
      }
    }

    writeRdpAuditLog_(reqUsername || 'admin', 'MIGRATION', 'LEGACY_IMPORT', `โอนย้ายข้อมูลจาก RDP Manager เดิม: ไซต์ +${sitesImported}, กลุ่ม +${groupsImported}, ผู้ใช้ +${usersUpdated}`);

    return {
      success: true,
      message: `โอนย้ายข้อมูลสำเร็จ! นำเข้าไซต์งานใหม่ ${sitesImported} ไซต์, กลุ่มงานใหม่ ${groupsImported} กลุ่ม, และอัปเดตผู้ใช้ ${usersUpdated} บัญชี`,
      sitesCount: sitesImported,
      groupsCount: groupsImported,
      usersCount: usersUpdated
    };
  } catch (err) {
    console.error('migrateFromLegacyRdpSpreadsheet error:', err);
    return { success: false, message: `เกิดข้อผิดพลาดในการโอนย้าย: ${err.message}` };
  }
}

/**
 * ดึงโปรไฟล์และสิทธิ์ของผู้ใช้งานจาก Username
 */
function getUserProfile_(username) {
  if (!username) return null;
  try {
    const ss = getDb_();
    let sheet = ss.getSheetByName(SHEET_USERS);
    if (!sheet) return null;

    const data = sheet.getDataRange().getValues();
    const target = String(username).trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === target) {
        // AllowedSites (col 7)
        let allowedSites = 'ALL';
        if (data[i][6] !== undefined && String(data[i][6]).trim() !== '') {
          const raw = String(data[i][6]).trim();
          if (raw !== 'ALL') {
            try { allowedSites = JSON.parse(raw); } catch (e) { allowedSites = 'ALL'; }
          }
        }

        // AllowedApps (col 9)
        let allowedApps = ['work_tracker', 'rdp_manager'];
        if (data[i][8] !== undefined && String(data[i][8]).trim() !== '') {
          try {
            const parsedApps = JSON.parse(String(data[i][8]));
            if (Array.isArray(parsedApps)) allowedApps = parsedApps;
          } catch (e) {}
        }
        if (String(data[i][3]).trim() === 'Admin' && !allowedApps.includes('admin')) {
          allowedApps.push('admin');
        }

        // AllowedRdpSites (col 10)
        let allowedRdpSites = 'ALL';
        if (data[i][9] !== undefined && String(data[i][9]).trim() !== '') {
          const rawR = String(data[i][9]).trim();
          if (rawR !== 'ALL') {
            try { allowedRdpSites = JSON.parse(rawR); } catch (e) { allowedRdpSites = 'ALL'; }
          }
        }

        return {
          username: String(data[i][0]).trim(),
          fullName: String(data[i][1] || data[i][0]).trim(),
          telegramChatId: String(data[i][2] || '').trim(),
          role: String(data[i][3] || 'Member').trim(),
          status: String(data[i][4] || 'Active').trim(),
          allowedSites: allowedSites,
          hasPassword: Boolean(data[i][7]),
          allowedApps: allowedApps,
          allowedRdpSites: allowedRdpSites
        };
      }
    }
  } catch (err) {
    console.error('getUserProfile_ error:', err);
  }
  return null;
}

/**
 * อัปเดตสถานะ Active Remote ของผู้ใช้ใน Presence Cache
 */
function updateUserActiveRemote_(username, activeRemoteObj) {
  if (!username) return;
  try {
    const lock = LockService.getScriptLock();
    try { lock.tryLock(2000); } catch (e) {}

    const cache = CacheService.getScriptCache();
    const raw = cache.get(CACHE_KEY_ONLINE_USERS);
    let map = {};
    if (raw) {
      try { map = JSON.parse(raw) || {}; } catch (e) { map = {}; }
    }

    const uName = String(username).trim().toLowerCase();
    if (map[uName]) {
      map[uName].activeRemote = activeRemoteObj;
      map[uName].lastActive = Date.now();
      cache.put(CACHE_KEY_ONLINE_USERS, JSON.stringify(map), 600);
    }

    try { lock.releaseLock(); } catch (e) {}
  } catch (err) {
    console.error('updateUserActiveRemote_ error:', err);
  }
}
