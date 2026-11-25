// ============================================================================
// NGI LEADER'S & ADMIN PORTAL - OFFLINE SYNC & ATTENDANCE BACKEND (Code.gs)
// ============================================================================

// Admin password is stored as a salted SHA-256 digest (never plaintext).
// The matching password is provided separately to the portal owner.
const NGI_ADMIN_EMAIL = "quickngiteam@gmail.com";
const NGI_ADMIN_PASSWORD_SALT = "Uma0kwIQU4myKlDl9eya94VjPcH9o5XPw5g1LkpK9W0";
const NGI_ADMIN_PASSWORD_HASH = "jE5qrwgSXeVnONiZ_T3uHEPLhQtgJA6rrfkReShLC1g";
const NGI_OTP_TTL_MS = 5 * 60 * 1000;
const NGI_ADMIN_DEVICE_LIMIT = 3;
const NGI_ADMIN_SESSION_MS = 12 * 60 * 60 * 1000;

const DEFAULT_LEADERS = [
  {
    fullName: "Zain Ul Abdin",
    position: "Founder & President",
    role: "PRESIDENT",
    memberId: "NGI-001",
    password: "LeaderPassword123!",
    email: "zain@ngiportal.com",
    status: "Active",
    avatar: "https://drive.google.com/thumbnail?id=1k5DRLmp4i9syfFFmfRM36kF3P2_u-qiT&sz=w400"
  },
  {
    fullName: "Mehak Kumbhar",
    position: "Media Director",
    role: "MEDIA_DIRECTOR",
    memberId: "NGI-002",
    password: "LeaderPassword123!",
    email: "mehak@ngiportal.com",
    status: "Active",
    avatar: "https://drive.google.com/thumbnail?id=19yL7iHRIHVraJ_0dmVhLEJ2ixosKkfLe&sz=w400"
  },
  {
    fullName: "Farooq Ahmed",
    position: "General Secretary",
    role: "GEN_SECRETARY",
    memberId: "NGI-003",
    password: "LeaderPassword123!",
    email: "farooq@ngiportal.com",
    status: "Active",
    avatar: "https://drive.google.com/thumbnail?id=1U-UltejCAiLrARiSRvKxgkpHkQ471JIX&sz=w400"
  }
];

function getSpreadsheet() {
  let active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;

  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty("AUTO_NGI_SHEET_ID");
  if (savedId) {
    try {
      return SpreadsheetApp.openById(savedId);
    } catch (e) {}
  }

  const newSs = SpreadsheetApp.create("NGI Portal Database (Auto-Generated)");
  props.setProperty("AUTO_NGI_SHEET_ID", newSs.getId());
  return newSs;
}

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle("NGI Leader's & Admin Portal")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const result = handleApiAction(payload);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: "Server Error: " + err.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function getOrCreateSheet(ss, sheetName, defaultHeaders) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (defaultHeaders && defaultHeaders.length > 0) {
      sheet.appendRow(defaultHeaders);
    }
  }
  return sheet;
}

function initializeSystemSheets(ss) {
  const membersSheet = getOrCreateSheet(ss, "Members", [
    "memberId", "fullName", "position", "role", "password", "email", "status", "avatar", "lastLogin", "isOnline"
  ]);
  if (membersSheet.getLastRow() <= 1) {
    DEFAULT_LEADERS.forEach(m => {
      membersSheet.appendRow([
        m.memberId, m.fullName, m.position, m.role, m.password, m.email, m.status, m.avatar, "", "false"
      ]);
    });
  }

  getOrCreateSheet(ss, "Attendance", [
    "attendanceId", "memberId", "memberName", "position", "date", "time", "status", "location", "notes", "syncedAt"
  ]);
  getOrCreateSheet(ss, "Notices", ["noticeId", "title", "content", "priority", "authorName", "timestamp", "emailSentCount"]);
  getOrCreateSheet(ss, "DirectNotes", ["noteId", "memberId", "memberName", "position", "subject", "message", "status", "adminReply", "timestamp"]);
  getOrCreateSheet(ss, "Complaints", ["complaintId", "memberId", "memberName", "position", "subject", "priority", "details", "status", "assignedTo", "feedback", "timestamp"]);
  getOrCreateSheet(ss, "Applications", ["trackingId", "applicantName", "positionApplied", "contactInfo", "notes", "status", "submittedAt"]);
  getOrCreateSheet(ss, "ActivityFeed", ["timestamp", "memberId", "memberName", "position", "section", "action", "summary"]);
  getOrCreateSheet(ss, "AdminOTP", ["otpId","otpHash","salt","createdAt","expiresAt","attempts","status"]);
  getOrCreateSheet(ss, "AdminDevices", ["deviceId","deviceName","deviceType","isMaster","status","createdAt","lastSeen"]);
  getOrCreateSheet(ss, "AdminSessions", ["sessionId","deviceId","tokenHash","salt","status","createdAt","expiresAt","lastSeen"]);
  getOrCreateSheet(ss, "EmailLog", ["emailId","recipient","subject","sentAt","status","error"]);
  getOrCreateSheet(ss, "PushTokens", ["tokenId","memberId","token","platform","status","createdAt","lastSeen"]);
  getOrCreateSheet(ss, "UserSessions", ["sessionId","memberId","tokenHash","salt","status","createdAt","expiresAt","lastSeen"]);
  getOrCreateSheet(ss, "PairRequests", ["pairId","pairHash","pairSalt","deviceName","createdAt","expiresAt","status"]);
  const devSheet = getOrCreateSheet(ss, "AdminDevices", ["deviceId","deviceName","deviceType","isMaster","status","createdAt","lastSeen","deviceTokenHash","deviceTokenSalt"]);
  const devHeaders = devSheet.getRange(1,1,1,Math.max(1,devSheet.getLastColumn())).getValues()[0].map(String);
  ["deviceTokenHash","deviceTokenSalt"].forEach(h => { if (devHeaders.indexOf(h) < 0) devSheet.getRange(1,devSheet.getLastColumn()+1).setValue(h); });

  const settingsSheet = getOrCreateSheet(ss, "Settings", ["key", "value"]);
  if (settingsSheet.getLastRow() <= 1) {
    settingsSheet.appendRow(["maintenanceMode", "false"]);
    settingsSheet.appendRow(["locked", "false"]);
    settingsSheet.appendRow(["loginEnabled", "true"]);
    settingsSheet.appendRow(["registrationEnabled", "true"]);
    settingsSheet.appendRow(["homeBannerText", ""]);
  }

  const defaultSheet = ss.getSheetByName("Sheet1");
  if (defaultSheet && ss.getSheets().length > 1) {
    try { ss.deleteSheet(defaultSheet); } catch (e) {}
  }
}

function sheetToObjects(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  const headers = values[0].map(h => String(h).trim());
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const obj = {};
    headers.forEach((h, idx) => {
      obj[h] = values[i][idx];
    });
    rows.push(obj);
  }
  return rows;
}

function getSettingsMap(ss) {
  const sheet = getOrCreateSheet(ss, "Settings", ["key", "value"]);
  const rows = sheetToObjects(sheet);
  const map = {
    maintenanceMode: false,
    locked: false,
    loginEnabled: true,
    registrationEnabled: true,
    homeBannerText: "",
    brandName: "NGI Leader's Portal",
    brandLogoUrl: ""
  };
  rows.forEach(r => {
    const k = String(r.key);
    const v = String(r.value);
    if (k === "maintenanceMode" || k === "locked") map[k] = (v === "true");
    else if (k === "loginEnabled" || k === "registrationEnabled") map[k] = (v !== "false");
    else map[k] = v;
  });
  return map;
}

function logAudit(ss, memberId, memberName, position, section, action, summary) {
  const sheet = getOrCreateSheet(ss, "ActivityFeed", ["timestamp", "memberId", "memberName", "position", "section", "action", "summary"]);
  sheet.appendRow([new Date().toISOString(), memberId, memberName, position, section, action, summary]);
}



function saveActivityImageToDrive_(dataUrl, memberId) {
  const match = String(dataUrl).match(/^data:(image\/(?:png|jpeg|webp));base64,([A-Za-z0-9+/=]+)$/);
  if (!match) throw new Error("Invalid image format. Use JPG, PNG or WEBP.");
  const bytes = Utilities.base64Decode(match[2]);
  if (bytes.length > 5 * 1024 * 1024) throw new Error("Image must be 5 MB or smaller.");
  const mime = match[1];
  const ext = mime === "image/jpeg" ? "jpg" : mime.split("/")[1];
  const props = PropertiesService.getScriptProperties();
  let folderId = props.getProperty("NGI_ACTIVITY_UPLOAD_FOLDER_ID");
  let folder = null;
  if (folderId) { try { folder = DriveApp.getFolderById(folderId); } catch (e) {} }
  if (!folder) {
    folder = DriveApp.createFolder("NGI Portal Activity Photos");
    props.setProperty("NGI_ACTIVITY_UPLOAD_FOLDER_ID", folder.getId());
  }
  const safeId = String(memberId).replace(/[^A-Za-z0-9_-]/g, "_");
  const blob = Utilities.newBlob(bytes, mime, "NGI_" + safeId + "_" + Date.now() + "." + ext);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return "https://drive.google.com/uc?export=view&id=" + file.getId();
}

function saveBrandLogo_(dataUrl) {
  if (!dataUrl || typeof dataUrl !== "string" || dataUrl.indexOf("data:image/") !== 0) return "";
  const match = dataUrl.match(/^data:(image\/(?:png|jpeg|webp));base64,(.+)$/i);
  if (!match) throw new Error("Unsupported logo image format.");
  const bytes = Utilities.base64Decode(match[2]);
  if (bytes.length > 2 * 1024 * 1024) throw new Error("Logo must be 2 MB or smaller.");
  const ext = match[1].toLowerCase().indexOf("png") >= 0 ? "png" : (match[1].toLowerCase().indexOf("webp") >= 0 ? "webp" : "jpg");
  const blob = Utilities.newBlob(bytes, match[1], "NGI_Portal_Logo." + ext);
  const props = PropertiesService.getScriptProperties();
  const oldId = props.getProperty("NGI_BRAND_LOGO_FILE_ID");
  const file = DriveApp.createFile(blob);
  file.setName("NGI Portal Brand Logo");
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  props.setProperty("NGI_BRAND_LOGO_FILE_ID", file.getId());
  if (oldId) { try { DriveApp.getFileById(oldId).setTrashed(true); } catch (e) {} }
  return "https://drive.google.com/uc?export=view&id=" + file.getId();
}

function handleApiAction(payload) {
  const ss = getSpreadsheet();
  initializeSystemSheets(ss);
  if (!payload || typeof payload !== "object" || Array.isArray(payload)) return {success:false,message:"Invalid request."};
  const action = String(payload.action || "");
  if (action === "requestMasterOtp") return ngiRequestMasterOtp_(ss, payload);
  if (action === "verifyMasterOtp") return ngiVerifyMasterOtp_(ss, payload);
  if (action === "adminDeviceStatus") return ngiAdminDeviceStatus_(ss, payload);
  if (action === "adminDeviceLogin") return ngiAdminDeviceLogin_(ss, payload);
  if (action === "completeAdminPairing") return ngiCompleteAdminPairing_(ss, payload);
  if (action === "createAdminPairRequest") return ngiCreateAdminPairRequest_(ss, payload);
  if (action === "adminListDevices") return ngiAdminListDevices_(ss, payload);
  if (action === "adminValidateSession") return ngiValidateAdminSession_(ss, payload);
  if (action === "adminLogout") return ngiAdminLogout_(ss, payload);
  if (action === "adminRegisterDevice") return ngiRegisterAdminDevice_(ss, payload);
  if (action === "adminRevokeDevice" || action === "adminRevokeDevice") return ngiRevokeAdminDevice_(ss, payload);
  if (action === "adminSendEmail") return ngiAdminSendEmail_(ss, payload);
  if (action === "registerPushToken" || action === "registerPush") return ngiRegisterPushToken_(ss, payload);
  if (action === "sendAnnouncement") return ngiSendAnnouncement_(ss, payload);
  if (action === "changeAdminPassword") return ngiChangeAdminPassword_(ss, payload);
  const adminActions = new Set(["deleteRecord","adminCreateMember","adminUpdateMemberProfile","adminDeleteMember","saveNotice","deleteNotice","updateNotice","replyDirectNote","updateComplaint","updateTicker","updatePortalSettings","saveApplication","updateApplicationStatus","adminExportApplication","adminDeleteApplication","adminUpdateMemberField","adminUpdateBranding"]);
  if (adminActions.has(action)) { const a=ngiValidateAdminSession_(ss,payload); if(!a.success)return {success:false,authExpired:true,message:a.message||"Admin session expired."}; }
  if (["markAttendance","saveRecord","updateOwnRecord","submitDirectNote","submitComplaint","userSubmitApplication"].includes(action)) { const m=ngiRequireMember_(ss,payload); if(!m.success)return m; payload.memberId=m.memberId; }
  if (action === "getDashboardData") {
    if (String(payload.role||"")==="ADMIN") { const a=ngiValidateAdminSession_(ss,payload); if(!a.success)return {success:false,authExpired:true,message:a.message||"Admin session expired."}; }
    else { const m=ngiRequireMember_(ss,payload); if(!m.success)return m; payload.memberId=m.memberId; }
  }
  const settings = getSettingsMap(ss);

  // 1. LOGIN
  if (action === "login") {
    const memberId = String(payload.memberId || "").trim();
    const password = String(payload.password || "");

    if (memberId.toUpperCase() === "NGI-ADMIN") {
      return { success: false, message: "Administrator password login has been retired. Use email OTP verification." };
    }

    if (settings.locked || settings.maintenanceMode) {
      return { success: false, locked: true, message: "Portal is currently paused or in maintenance mode by Administration." };
    }
    if (!settings.loginEnabled) {
      return { success: false, message: "Member login system is temporarily disabled by Admin." };
    }

    const membersSheet = ss.getSheetByName("Members");
    const data = membersSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === memberId.toLowerCase()) {
        const storedPass = String(data[i][4]);
        const status = String(data[i][6] || "Active");
        if (storedPass !== password) {
          return { success: false, message: "Incorrect password. Please try again." };
        }
        if (status !== "Active") {
          return { success: false, message: "Account is " + status + ". Contact Administration." };
        }

        const nowIso = new Date().toISOString();
        membersSheet.getRange(i + 1, 9).setValue(nowIso);
        membersSheet.getRange(i + 1, 10).setValue("true");

        const userObj = {
          memberId: String(data[i][0]),
          fullName: String(data[i][1]),
          position: String(data[i][2]),
          role: String(data[i][3]),
          email: String(data[i][5] || ""),
          status: status,
          avatar: String(data[i][7] || "")
        };

        const memberToken=ngiToken_(), memberSalt=ngiToken_(), issued=new Date(), expiry=new Date(issued.getTime()+12*60*60*1000);
        ss.getSheetByName("UserSessions").appendRow([Utilities.getUuid(),userObj.memberId,ngiDigest_(memberToken,memberSalt),memberSalt,"Active",issued.toISOString(),expiry.toISOString(),issued.toISOString()]);
        userObj.token=memberToken;
        logAudit(ss, userObj.memberId, userObj.fullName, userObj.position, "Auth", "Member Login", "Signed into portal");
        return { success: true, user: userObj };
      }
    }
    return { success: false, message: "Unique Member ID not found." };
  }

  // 2. GET DASHBOARD DATA (ADMIN SEES EVERYTHING; MEMBER SEES SYSTEM ANNOUNCEMENTS)
  if (action === "getDashboardData") {
    const role = String(payload.role || "");
    const memberId = String(payload.memberId || "");

    if (role !== "ADMIN" && (settings.locked || settings.maintenanceMode)) {
      return { success: false, locked: true, message: "Portal is currently locked by Administration." };
    }

    const systemSheets = new Set(["Members", "Attendance", "Notices", "DirectNotes", "Complaints", "Applications", "ActivityFeed", "Settings"]);
    const recordsMap = {};
    ss.getSheets().forEach(sh => {
      const name = sh.getName();
      if (systemSheets.has(name)) return;
      const rows = sheetToObjects(sh);
      // Admin receives all records; members receive only their own records.
      recordsMap[name] = role === "ADMIN" ? rows : rows.filter(r =>
        String(r.MemberID || r.memberId || "").trim().toLowerCase() === memberId.toLowerCase()
      );
    });

    const members = sheetToObjects(ss.getSheetByName("Members")).map(m => ({
      memberId: m.memberId,
      fullName: m.fullName,
      position: m.position,
      role: m.role,
      email: m.email,
      status: m.status || "Active",
      avatar: m.avatar,
      lastLogin: m.lastLogin,
      isOnline: String(m.isOnline) === "true"
    }));

    const allAttendance = sheetToObjects(ss.getSheetByName("Attendance")).reverse();
    const attendance = role === "ADMIN" ? allAttendance : allAttendance.filter(a => String(a.memberId || "").toLowerCase() === memberId.toLowerCase());
    const notices = sheetToObjects(ss.getSheetByName("Notices")).reverse();
    let directNotes = sheetToObjects(ss.getSheetByName("DirectNotes")).reverse();
    let complaints = sheetToObjects(ss.getSheetByName("Complaints")).reverse();

    if (role !== "ADMIN") {
      directNotes = directNotes.filter(n => String(n.memberId).toLowerCase() === memberId.toLowerCase());
      complaints = complaints.filter(c => String(c.memberId).toLowerCase() === memberId.toLowerCase());
    }

    const applications = sheetToObjects(ss.getSheetByName("Applications")).reverse();
    const activityFeed = sheetToObjects(ss.getSheetByName("ActivityFeed")).reverse().slice(0, 120);

    return {
      success: true,
      data: {
        records: recordsMap,
        attendance: attendance,
        members: role === "ADMIN" ? members : [],
        notices: notices,
        directNotes: directNotes,
        complaints: complaints,
        applications: role === "ADMIN" ? applications : [],
        activityFeed: role === "ADMIN" ? activityFeed : activityFeed.filter(a => String(a.memberId || "").toLowerCase() === memberId.toLowerCase()),
        settings: settings
      }
    };
  }

  // 3. MARK / SYNC ATTENDANCE
  if (action === "markAttendance") {
    const memberId = String(payload.memberId || "");
    const members = sheetToObjects(ss.getSheetByName("Members"));
    const memberInfo = members.find(m => String(m.memberId).toLowerCase() === memberId.toLowerCase()) || {
      fullName: payload.memberName || memberId,
      position: payload.position || "Leader"
    };

    const sheet = getOrCreateSheet(ss, "Attendance", [
      "attendanceId", "memberId", "memberName", "position", "date", "time", "status", "location", "notes", "syncedAt"
    ]);
    const attId = "ATT-" + Date.now() + "-" + Math.floor(Math.random() * 1000);
    const nowIso = new Date().toISOString();

    sheet.appendRow([
      attId,
      memberId,
      memberInfo.fullName,
      memberInfo.position,
      payload.date || nowIso.slice(0, 10),
      payload.time || new Date().toLocaleTimeString(),
      payload.status || "Present",
      payload.location || "Field / Office",
      payload.notes || "",
      nowIso
    ]);

    logAudit(ss, memberId, memberInfo.fullName, memberInfo.position, "Attendance", "Synced Attendance", (payload.status || "Present") + " (" + (payload.location || "Field") + ")");
    return { success: true, attendanceId: attId, message: "Attendance saved to Sheet." };
  }

  // 4. SAVE OR SYNC RECORD (DYNAMIC COLUMNS)
  if (action === "saveRecord") {
    const memberId = String(payload.memberId || "");
    const section = String(payload.section || "").trim();
    const fields = payload.fields || {};
    const protectedSheets = new Set(["members","attendance","notices","directnotes","complaints","applications","activityfeed","settings","adminotp","admindevices","adminsessions","usersessions","pairrequests","pushtokens","emaillog"]);
    if (!section || section.length > 80 || /[\\/\?*\[\]:]/.test(section) || protectedSheets.has(section.toLowerCase())) return {success:false,message:"This section is reserved or invalid."};
    if (!fields || typeof fields !== "object" || Array.isArray(fields)) return {success:false,message:"Invalid record fields."};
    const nowIso = new Date().toISOString();

    const members = sheetToObjects(ss.getSheetByName("Members"));
    const memberInfo = members.find(m => String(m.memberId).toLowerCase() === memberId.toLowerCase()) || {
      fullName: memberId,
      position: payload.role || "Leader"
    };

    const fieldKeys = Object.keys(fields);
    Object.keys(fields).forEach(key => {
      const value = fields[key];
      if (typeof value === "string" && value.indexOf("data:image/") === 0) {
        fields[key] = saveActivityImageToDrive_(value, memberId);
      }
    });
    const baseHeaders = ["RecordID", "MemberID", "MemberName", "CreatedAt", "SyncedAt"];
    const sheet = getOrCreateSheet(ss, section, baseHeaders.concat(fieldKeys));

    const data = sheet.getDataRange().getValues();
    let headers = data[0].map(h => String(h).trim());

    fieldKeys.forEach(fk => {
      if (!headers.includes(fk)) {
        headers.push(fk);
        sheet.getRange(1, headers.length).setValue(fk);
      }
    });

    const finalRecordId = "REC-" + Date.now() + "-" + Math.floor(Math.random() * 1000);
    const createdAt = payload.createdAt || nowIso;

    const newRow = headers.map(h => {
      if (h === "RecordID") return finalRecordId;
      if (h === "MemberID") return memberId;
      if (h === "MemberName") return memberInfo.fullName;
      if (h === "CreatedAt") return createdAt;
      if (h === "SyncedAt") return nowIso;
      return fields[h] !== undefined ? fields[h] : "";
    });
    sheet.appendRow(newRow);

    const firstVal = fieldKeys.length > 0 ? fields[fieldKeys[0]] : finalRecordId;
    logAudit(ss, memberId, memberInfo.fullName, memberInfo.position, section, "Synced Field Entry", String(firstVal));
    return { success: true, recordId: finalRecordId, message: "Record synced to Google Sheet." };
  }

  // Members may edit only their own record for 10 minutes after creation.
  if (action === "updateOwnRecord") {
    const memberId = String(payload.memberId || "").trim();
    const section = String(payload.section || "").trim();
    const recordId = String(payload.recordId || "").trim();
    const fields = payload.fields || {};
    const sheet = ss.getSheetByName(section);
    if (!memberId || !recordId || !sheet) return { success: false, message: "Invalid record request." };
    const values = sheet.getDataRange().getValues();
    if (!values.length) return { success: false, message: "Record not found." };
    const headers = values[0].map(h => String(h).trim());
    const idCol = headers.indexOf("RecordID");
    const ownerCol = headers.indexOf("MemberID");
    const createdCol = headers.indexOf("CreatedAt");
    if (idCol < 0 || ownerCol < 0 || createdCol < 0) return { success: false, message: "Record format is not editable." };
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][idCol]) !== recordId) continue;
      if (String(values[i][ownerCol]).toLowerCase() !== memberId.toLowerCase()) return { success: false, message: "You can edit only your own record." };
      const created = new Date(values[i][createdCol]).getTime();
      if (!created || Date.now() - created > 10 * 60 * 1000 || Date.now() < created) return { success: false, message: "The 10-minute editing period has expired." };
      Object.keys(fields).forEach(key => {
        if (["RecordID", "MemberID", "MemberName", "CreatedAt", "SyncedAt"].includes(key)) return;
        let value = fields[key];
        if (typeof value === "string" && value.indexOf("data:image/") === 0) value = saveActivityImageToDrive_(value, memberId);
        let col = headers.indexOf(key);
        if (col < 0) { col = headers.length; headers.push(key); sheet.getRange(1, col + 1).setValue(key); }
        sheet.getRange(i + 1, col + 1).setValue(value);
      });
      const syncCol = headers.indexOf("SyncedAt");
      if (syncCol >= 0) sheet.getRange(i + 1, syncCol + 1).setValue(new Date().toISOString());
      logAudit(ss, memberId, String(values[i][headers.indexOf("MemberName")] || memberId), "Member", section, "Edited Own Record", recordId);
      return { success: true, message: "Record updated." };
    }
    return { success: false, message: "Record not found." };
  }

  // 5. ADMIN: DELETE RECORD OR ATTENDANCE
  if (action === "deleteRecord") {
    const section = String(payload.section || "");
    const recordId = String(payload.recordId || "");
    const sheet = ss.getSheetByName(section);
    if (!sheet) return { success: false, message: "Sheet not found." };

    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]) === recordId) {
        sheet.deleteRow(i + 1);
        logAudit(ss, "NGI-ADMIN", "NGI System Admin", "Administrator", section, "Deleted Record", "ID: " + recordId);
        return { success: true, message: "Deleted permanently from Google Sheet." };
      }
    }
    return { success: false, message: "Record not found." };
  }

  // 6. ADMIN: USER MANAGEMENT
  if (action === "adminCreateMember") {
    if (payload.actorRole !== "ADMIN") return { success: false, message: "Unauthorized." };
    const membersSheet = ss.getSheetByName("Members");
    const data = membersSheet.getDataRange().getValues();
    const newId = String(payload.memberId || "").trim();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === newId.toLowerCase()) {
        return { success: false, message: "Member ID already exists." };
      }
    }

    membersSheet.appendRow([
      newId,
      String(payload.fullName || "").trim(),
      String(payload.position || "").trim(),
      String(payload.schemaRole || "GENERAL_LEADER"),
      String(payload.password || ""),
      String(payload.email || "").trim(),
      "Active",
      String(payload.avatarUrl || "").trim(),
      "",
      "false"
    ]);
    return { success: true, message: "User created successfully." };
  }

  if (action === "adminUpdateMemberProfile") {
    if (payload.actorRole !== "ADMIN") return { success: false, message: "Unauthorized." };
    const membersSheet = ss.getSheetByName("Members");
    const data = membersSheet.getDataRange().getValues();
    const origId = String(payload.originalMemberId || "").trim();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === origId.toLowerCase()) {
        membersSheet.getRange(i + 1, 1).setValue(String(payload.newMemberId || origId).trim());
        membersSheet.getRange(i + 1, 2).setValue(String(payload.fullName || "").trim());
        membersSheet.getRange(i + 1, 3).setValue(String(payload.position || "").trim());
        membersSheet.getRange(i + 1, 4).setValue(String(payload.schemaRole || "GENERAL_LEADER"));
        if (payload.newPassword && String(payload.newPassword).trim() !== "") {
          membersSheet.getRange(i + 1, 5).setValue(String(payload.newPassword).trim());
        }
        membersSheet.getRange(i + 1, 6).setValue(String(payload.email || "").trim());
        membersSheet.getRange(i + 1, 7).setValue(String(payload.status || "Active"));
        membersSheet.getRange(i + 1, 8).setValue(String(payload.avatarUrl || "").trim());
        return { success: true, message: "User updated successfully." };
      }
    }
    return { success: false, message: "User not found." };
  }

  if (action === "adminDeleteMember") {
    if (payload.actorRole !== "ADMIN") return { success: false, message: "Unauthorized." };
    const membersSheet = ss.getSheetByName("Members");
    const data = membersSheet.getDataRange().getValues();
    const targetId = String(payload.targetMemberId || "").trim();
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]).trim().toLowerCase() === targetId.toLowerCase()) {
        membersSheet.deleteRow(i + 1);
        return { success: true, message: "User deleted." };
      }
    }
    return { success: false, message: "User not found." };
  }

  // Applications and profile/branding actions used by the current portal UI.
  if (action === "userSubmitApplication") {
    const sheet=ss.getSheetByName("Applications"), trackingId="NGI-APP-"+Utilities.getUuid().slice(0,8).toUpperCase(), now=new Date().toISOString();
    sheet.appendRow([trackingId,String(payload.applicantName||"").trim(),String(payload.positionApplied||"").trim(),String(payload.contactInfo||"").trim(),String(payload.notes||"").trim(),"Pending",now]);
    return {success:true,trackingCode:trackingId,message:"Application submitted."};
  }
  if (action === "saveApplication") {
    const sheet=ss.getSheetByName("Applications"), trackingId="NGI-APP-"+Utilities.getUuid().slice(0,8).toUpperCase(), now=new Date().toISOString();
    sheet.appendRow([trackingId,String(payload.applicantName||"").trim(),String(payload.positionApplied||"").trim(),String(payload.contactInfo||"").trim(),String(payload.notes||"").trim(),"Pending",now]);
    return {success:true,trackingId:trackingId,message:"Application created: "+trackingId};
  }
  if (action === "updateApplicationStatus") {
    const sh=ss.getSheetByName("Applications"), rows=sh.getDataRange().getValues();
    for(let i=1;i<rows.length;i++) if(String(rows[i][0])===String(payload.trackingId)){ sh.getRange(i+1,6).setValue(String(payload.status||"Pending")); return {success:true,message:"Application status updated."}; }
    return {success:false,message:"Application not found."};
  }
  if (action === "adminExportApplication") {
    const rec=sheetToObjects(ss.getSheetByName("Applications")).find(x=>String(x.trackingId)===String(payload.trackingId));
    return rec?{success:true,record:rec}:{success:false,message:"Application not found."};
  }
  if (action === "adminDeleteApplication") {
    const sh=ss.getSheetByName("Applications"), rows=sh.getDataRange().getValues();
    for(let i=rows.length-1;i>=1;i--) if(String(rows[i][0])===String(payload.trackingId)){sh.deleteRow(i+1);return {success:true,message:"Application deleted."};}
    return {success:false,message:"Application not found."};
  }
  if (action === "adminUpdateMemberField") {
    const sh=ss.getSheetByName("Members"), rows=sh.getDataRange().getValues(), id=String(payload.memberId||"").trim(), field=String(payload.field||""), allowed={fullName:2,position:3,role:4,memberId:1,status:7,email:6,avatar:8,password:5};
    if(!allowed[field]) return {success:false,message:"Invalid member field."};
    for(let i=1;i<rows.length;i++) if(String(rows[i][0]).toLowerCase()===id.toLowerCase()) { sh.getRange(i+1,allowed[field]).setValue(String(payload.value||"")); return {success:true,message:"Member "+field+" updated."}; }
    return {success:false,message:"Member not found."};
  }
  if (action === "adminUpdateBranding") {
    const sh=ss.getSheetByName("Settings"), key=String(payload.field)==="brandLogo"?"brandLogoUrl":"brandName";
    let value=String(payload.value||""); if(key==="brandLogoUrl"&&value.indexOf("data:image/")===0)value=saveBrandLogo_(value);
    const rows=sh.getDataRange().getValues(); let found=false;
    for(let i=1;i<rows.length;i++)if(String(rows[i][0])===key){sh.getRange(i+1,2).setValue(value);found=true;break;}
    if(!found)sh.appendRow([key,value]); return {success:true,message:"Branding updated."};
  }

  // 7. NOTICES, NOTES, COMPLAINTS, SETTINGS
  if (action === "saveNotice") {
    if (payload.actorRole !== "ADMIN") return { success: false, message: "Unauthorized." };
    const sheet = ss.getSheetByName("Notices");
    sheet.appendRow(["NTC-" + Date.now(), payload.title, payload.content, payload.priority, "NGI Administration", new Date().toISOString(), 0]);
    return { success: true, message: "Notice published." };
  }

  if (action === "deleteNotice") {
    if (payload.actorRole !== "ADMIN") return { success: false, message: "Unauthorized." };
    const sheet = ss.getSheetByName("Notices");
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]) === String(payload.noticeId)) {
        sheet.deleteRow(i + 1);
        return { success: true, message: "Notice deleted." };
      }
    }
    return { success: false, message: "Notice not found." };
  }

  if (action === "updateNotice") {
    if (payload.actorRole !== "ADMIN") return { success: false, message: "Unauthorized." };
    const sheet = ss.getSheetByName("Notices");
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(payload.noticeId)) {
        sheet.getRange(i + 1, 2).setValue(String(payload.title || ""));
        sheet.getRange(i + 1, 3).setValue(String(payload.content || ""));
        sheet.getRange(i + 1, 4).setValue(String(payload.priority || "Normal"));
        return { success: true, message: "Announcement updated." };
      }
    }
    return { success: false, message: "Notice not found." };
  }

  if (action === "submitDirectNote") {
    const members = sheetToObjects(ss.getSheetByName("Members"));
    const m = members.find(x => String(x.memberId).toLowerCase() === String(payload.memberId).toLowerCase()) || { fullName: payload.memberId, position: "Leader" };
    ss.getSheetByName("DirectNotes").appendRow(["NOTE-" + Date.now(), payload.memberId, m.fullName, m.position, payload.subject, payload.message, "Unread", "", new Date().toISOString()]);
    return { success: true, message: "Note sent." };
  }

  if (action === "replyDirectNote") {
    if (payload.actorRole !== "ADMIN") return { success: false, message: "Unauthorized." };
    const sheet = ss.getSheetByName("DirectNotes");
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(payload.noteId)) {
        sheet.getRange(i + 1, 7).setValue("Answered");
        sheet.getRange(i + 1, 8).setValue(payload.replyText);
        return { success: true, message: "Reply sent." };
      }
    }
    return { success: false, message: "Not found." };
  }

  if (action === "submitComplaint") {
    const members = sheetToObjects(ss.getSheetByName("Members"));
    const m = members.find(x => String(x.memberId).toLowerCase() === String(payload.memberId).toLowerCase()) || { fullName: payload.memberId, position: "Leader" };
    ss.getSheetByName("Complaints").appendRow(["CMP-" + Date.now(), payload.memberId, m.fullName, m.position, payload.subject, payload.priority, payload.details, "New", "NGI System Admin", "", new Date().toISOString()]);
    return { success: true, message: "Complaint submitted." };
  }

  if (action === "updateComplaint") {
    if (payload.actorRole !== "ADMIN") return { success: false, message: "Unauthorized." };
    const sheet = ss.getSheetByName("Complaints");
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(payload.complaintId)) {
        sheet.getRange(i + 1, 8).setValue(payload.status);
        sheet.getRange(i + 1, 9).setValue(payload.assignedTo);
        sheet.getRange(i + 1, 10).setValue(payload.feedback);
        return { success: true, message: "Updated." };
      }
    }
    return { success: false, message: "Not found." };
  }

  if (action === "updateTicker") {
    if (payload.actorRole !== "ADMIN") return { success: false, message: "Unauthorized." };
    const sheet = ss.getSheetByName("Settings");
    const values = sheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === "homeBannerText") {
        sheet.getRange(i + 1, 2).setValue(String(payload.homeBannerText || ""));
        found = true;
        break;
      }
    }
    if (!found) sheet.appendRow(["homeBannerText", String(payload.homeBannerText || "")]);
    return { success: true, message: payload.homeBannerText ? "Scrolling ticker updated." : "Scrolling ticker deleted." };
  }

  if (action === "updatePortalSettings") {
    if (payload.actorRole !== "ADMIN") return { success: false, message: "Unauthorized." };
    const sheet = ss.getSheetByName("Settings");
    sheet.clearContents();
    sheet.appendRow(["key", "value"]);
    sheet.appendRow(["maintenanceMode", String(Boolean(payload.maintenanceMode))]);
    sheet.appendRow(["locked", String(Boolean(payload.locked))]);
    sheet.appendRow(["loginEnabled", String(Boolean(payload.loginEnabled))]);
    sheet.appendRow(["registrationEnabled", String(Boolean(payload.registrationEnabled))]);
    sheet.appendRow(["homeBannerText", String(payload.homeBannerText || "")]);
    sheet.appendRow(["brandName", String(payload.brandName || "NGI Leader's Portal")]);
    let brandLogoUrl = String(payload.brandLogoUrl || settings.brandLogoUrl || "");
    if (payload.brandLogoDataUrl) brandLogoUrl = saveBrandLogo_(payload.brandLogoDataUrl);
    if (brandLogoUrl) sheet.appendRow(["brandLogoUrl", brandLogoUrl]);
    return { success: true, message: "Settings updated." };
  }

  return { success: false, message: "Unknown action: " + action };
}

/* ===== NGI ADMIN EMAIL OTP / DEVICE TRUST ===== */
function ngiDigest_(value, salt) {
  return Utilities.base64EncodeWebSafe(Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, String(salt) + "|" + String(value), Utilities.Charset.UTF_8));
}
function ngiToken_() { return Utilities.getUuid().replace(/-/g,"") + Utilities.getUuid().replace(/-/g,""); }
function ngiRows_(ss, name) { return sheetToObjects(ss.getSheetByName(name)); }
function ngiUpdateBy_(ss, name, key, value, changes) {
  const sh=ss.getSheetByName(name), rows=ngiRows_(ss,name), r=rows.find(x=>String(x[key])===String(value));
  if(!r) return false;
  const headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String);
  Object.keys(changes).forEach(k=>{const c=headers.indexOf(k);if(c>=0)sh.getRange(r._row||rows.indexOf(r)+2,c+1).setValue(changes[k]);});
  return true;
}
function ngiAdminPasswordCredentials_() {
  const props = PropertiesService.getScriptProperties();
  return {
    salt: props.getProperty("NGI_ADMIN_PASSWORD_SALT") || NGI_ADMIN_PASSWORD_SALT,
    hash: props.getProperty("NGI_ADMIN_PASSWORD_HASH") || NGI_ADMIN_PASSWORD_HASH
  };
}
function ngiRequestMasterOtp_(ss, p) {
  const cache=CacheService.getScriptCache();
  if(cache.get("ngi_master_otp_cooldown")) return {success:false,message:"A verification code was sent recently. Please wait before requesting another."};
  cache.put("ngi_master_otp_cooldown","1",60);
  const otp = String(Math.floor(100000 + Math.random() * 900000));
  const salt = ngiToken_(), now = new Date();
  ngiRows_(ss,"AdminOTP").filter(r=>r.status==="Pending").forEach(r=>ngiUpdateBy_(ss,"AdminOTP","otpId",r.otpId,{status:"Expired"}));
  ss.getSheetByName("AdminOTP").appendRow([Utilities.getUuid(),ngiDigest_(otp,salt),salt,now.toISOString(),new Date(now.getTime()+NGI_OTP_TTL_MS).toISOString(),0,"Pending"]);
  MailApp.sendEmail({to:NGI_ADMIN_EMAIL,subject:"NGI Portal — Admin Verification Code",body:"Your NGI Portal OTP is "+otp+". It expires in five minutes. Do not share it."});
  return {success:true,message:"Password verified. OTP sent to the registered administrator email."};
}
function ngiChangeAdminPassword_(ss, p) {
  const auth = ngiRequireAdmin_(ss, p);
  const currentPassword = String(p.currentPassword || "");
  const newPassword = String(p.newPassword || "");
  const credentials = ngiAdminPasswordCredentials_();
  if (!currentPassword || ngiDigest_(currentPassword, credentials.salt) !== credentials.hash) {
    return {success:false, message:"Current administrator password is incorrect."};
  }
  if (newPassword.length < 12 || newPassword.length > 128) {
    return {success:false, message:"New password must be between 12 and 128 characters."};
  }
  if (newPassword === currentPassword) {
    return {success:false, message:"New password must be different from the current password."};
  }
  const newSalt = ngiToken_();
  const props = PropertiesService.getScriptProperties();
  props.setProperties({NGI_ADMIN_PASSWORD_SALT:newSalt, NGI_ADMIN_PASSWORD_HASH:ngiDigest_(newPassword,newSalt)});

  // Keep the current authenticated session; revoke every other active admin session.
  const sessions = ngiRows_(ss,"AdminSessions");
  const current = sessions.find(x => x.deviceId === p.deviceId && x.status === "Active" && ngiDigest_(p.adminToken || "",x.salt) === x.tokenHash);
  sessions.filter(x => x.status === "Active" && (!current || x.sessionId !== current.sessionId))
    .forEach(x => ngiUpdateBy_(ss,"AdminSessions","sessionId",x.sessionId,{status:"Revoked"}));
  logAudit(ss,"NGI-ADMIN","NGI System Admin","Administrator","Security","Administrator Password Changed","Other admin sessions revoked");
  return {success:true,message:"Administrator password updated. Other admin sessions have been signed out."};
}
function ngiVerifyMasterOtp_(ss,p) {
  const pending=ngiRows_(ss,"AdminOTP").filter(r=>r.status==="Pending").sort((a,b)=>new Date(b.createdAt)-new Date(a.createdAt));
  if(!pending.length) return {success:false,message:"No active OTP. Request a new code."};
  const r=pending[0], attempts=Number(r.attempts||0)+1;
  ngiUpdateBy_(ss,"AdminOTP","otpId",r.otpId,{attempts:attempts});
  if(Date.now()>new Date(r.expiresAt).getTime()){ngiUpdateBy_(ss,"AdminOTP","otpId",r.otpId,{status:"Expired"});return {success:false,message:"OTP expired. Request a new code."};}
  if(attempts>5){ngiUpdateBy_(ss,"AdminOTP","otpId",r.otpId,{status:"Locked"});return {success:false,message:"Too many attempts. Request a new OTP."};}
  if(ngiDigest_(String(p.otp||"").trim(),r.salt)!==r.otpHash) return {success:false,message:"Incorrect OTP."};
  ngiUpdateBy_(ss,"AdminOTP","otpId",r.otpId,{status:"Used"});

  // A successful password + email OTP is the authorized master-device recovery.
  // Revoke the previous master and its sessions before assigning this device.
  const oldDevices = ngiRows_(ss,"AdminDevices").filter(d=>String(d.status)==="Active");
  oldDevices.forEach(d=>{
    ngiUpdateBy_(ss,"AdminDevices","deviceId",d.deviceId,{status:"Revoked"});
    ngiRows_(ss,"AdminSessions").filter(x=>x.deviceId===d.deviceId && x.status==="Active").forEach(x=>ngiUpdateBy_(ss,"AdminSessions","sessionId",x.sessionId,{status:"Revoked"}));
  });
  const deviceId=Utilities.getUuid(), now=new Date(), token=ngiToken_(), salt=ngiToken_(), deviceToken=ngiToken_(), deviceSalt=ngiToken_();
  ss.getSheetByName("AdminDevices").appendRow([deviceId,String(p.deviceName||"Master Device").slice(0,80),String(p.deviceType||"Unknown"),true,"Active",now.toISOString(),now.toISOString(),ngiDigest_(deviceToken,deviceSalt),deviceSalt]);
  const expiry=new Date(now.getTime()+NGI_ADMIN_SESSION_MS);
  ss.getSheetByName("AdminSessions").appendRow([Utilities.getUuid(),deviceId,ngiDigest_(token,salt),salt,"Active",now.toISOString(),expiry.toISOString(),now.toISOString()]);
  logAudit(ss,"NGI-ADMIN","NGI System Admin","Administrator","Security",oldDevices.length ? "All Previous Devices Revoked; New Master Registered" : "Master Device Registered",deviceId);
  return {success:true,isMaster:true,deviceId:deviceId,deviceToken:deviceToken,adminToken:token,token:token,expiresAt:expiry.toISOString(),user:{memberId:"NGI-ADMIN",fullName:"NGI System Admin",position:"Full System Administrator",role:"ADMIN",avatar:"",deviceId:deviceId,adminToken:token,token:token,adminExpiresAt:expiry.toISOString()}};
}
function ngiValidateAdminSession_(ss,p) {
  const d=ngiRows_(ss,"AdminDevices").find(x=>x.deviceId===p.deviceId&&x.status==="Active");
  if(!d)return {success:false,message:"Device is not authorized."};
  const suppliedToken=String(p.adminToken||p.token||"");
  const s=ngiRows_(ss,"AdminSessions").find(x=>x.deviceId===p.deviceId&&x.status==="Active"&&Date.now()<=new Date(x.expiresAt).getTime()&&ngiDigest_(suppliedToken,x.salt)===x.tokenHash);
  if(!s)return {success:false,message:"Admin session expired."};
  ngiUpdateBy_(ss,"AdminSessions","sessionId",s.sessionId,{lastSeen:new Date().toISOString()});
  ngiUpdateBy_(ss,"AdminDevices","deviceId",d.deviceId,{lastSeen:new Date().toISOString()});
  return {success:true,isMaster:String(d.isMaster)==="true",deviceId:d.deviceId};
}

function ngiRequireMember_(ss,p) {
  const memberId=String(p.memberId||"").trim(), token=String(p.memberSessionToken||p.token||"");
  if(!memberId||!token)return {success:false,authExpired:true,message:"Member session missing. Please sign in again."};
  const member=ngiRows_(ss,"Members").find(m=>String(m.memberId).toLowerCase()===memberId.toLowerCase()&&String(m.status||"Active")==="Active");
  if(!member)return {success:false,authExpired:true,message:"Member account is unavailable."};
  const sess=ngiRows_(ss,"UserSessions").find(x=>String(x.memberId).toLowerCase()===memberId.toLowerCase()&&x.status==="Active"&&Date.now()<=new Date(x.expiresAt).getTime()&&ngiDigest_(token,x.salt)===x.tokenHash);
  if(!sess)return {success:false,authExpired:true,message:"Member session expired. Please sign in again."};
  ngiUpdateBy_(ss,"UserSessions","sessionId",sess.sessionId,{lastSeen:new Date().toISOString()});
  return {success:true,memberId:member.memberId};
}
function ngiAdminDeviceStatus_(ss,p) {
  const t=String(p.deviceToken||""); if(!t)return {success:true,authorized:false};
  const d=ngiRows_(ss,"AdminDevices").find(x=>x.status==="Active"&&x.deviceTokenSalt&&ngiDigest_(t,x.deviceTokenSalt)===x.deviceTokenHash);
  return {success:true,authorized:!!d,isMaster:!!(d&&String(d.isMaster)==="true")};
}
function ngiAdminDeviceLogin_(ss,p) {
  const t=String(p.deviceToken||"");
  const d=ngiRows_(ss,"AdminDevices").find(x=>x.status==="Active"&&x.deviceTokenSalt&&ngiDigest_(t,x.deviceTokenSalt)===x.deviceTokenHash);
  if(!d)return {success:false,revoked:true,message:"This device is no longer authorized. Register it again."};
  const now=new Date(), token=ngiToken_(), salt=ngiToken_(), exp=new Date(now.getTime()+NGI_ADMIN_SESSION_MS);
  ss.getSheetByName("AdminSessions").appendRow([Utilities.getUuid(),d.deviceId,ngiDigest_(token,salt),salt,"Active",now.toISOString(),exp.toISOString(),now.toISOString()]);
  ngiUpdateBy_(ss,"AdminDevices","deviceId",d.deviceId,{lastSeen:now.toISOString()});
  return {success:true,isMaster:String(d.isMaster)==="true",deviceId:d.deviceId,token:token,adminToken:token,expiresAt:exp.toISOString()};
}
function ngiAdminListDevices_(ss,p) {
  const a=ngiRequireAdmin_(ss,p); if(!a)return {success:false,message:"Unauthorized."};
  const rows=ngiRows_(ss,"AdminDevices").filter(d=>d.status==="Active").map(d=>({deviceId:d.deviceId,name:d.deviceName,isMaster:String(d.isMaster)==="true",createdAt:d.createdAt,lastSeen:d.lastSeen}));
  return {success:true,devices:rows,currentDeviceId:p.deviceId,isMaster:!!(rows.find(d=>d.deviceId===p.deviceId)||{}).isMaster};
}
function ngiCreateAdminPairRequest_(ss,p) {
  const a=ngiRequireAdmin_(ss,p); if(!a.isMaster)return {success:false,message:"Only the master device can authorize new devices."};
  if(ngiRows_(ss,"AdminDevices").filter(d=>d.status==="Active").length>=NGI_ADMIN_DEVICE_LIMIT)return {success:false,message:"Maximum three active admin devices reached."};
  const code=String(Math.floor(10000000+Math.random()*90000000)), salt=ngiToken_(), now=new Date();
  ss.getSheetByName("PairRequests").appendRow([Utilities.getUuid(),ngiDigest_(code,salt),salt,String(p.deviceName||"Admin Device").slice(0,80),now.toISOString(),new Date(now.getTime()+10*60*1000).toISOString(),"Pending"]);
  return {success:true,pairCode:code,expiresInMinutes:10};
}
function ngiCompleteAdminPairing_(ss,p) {
  const lock=LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const code=String(p.pairCode||"").trim(), rows=ngiRows_(ss,"PairRequests");
    const r=rows.find(x=>x.status==="Pending"&&Date.now()<=new Date(x.expiresAt).getTime()&&ngiDigest_(code,x.pairSalt)===x.pairHash);
    if(!r)return {success:false,message:"Invalid or expired authorization code."};
    if(ngiRows_(ss,"AdminDevices").filter(d=>d.status==="Active").length>=NGI_ADMIN_DEVICE_LIMIT)return {success:false,message:"Maximum three active admin devices reached."};
    ngiUpdateBy_(ss,"PairRequests","pairId",r.pairId,{status:"Used"});
    const id=Utilities.getUuid(), now=new Date(), token=ngiToken_(), deviceSalt=ngiToken_(), sessionSalt=ngiToken_(), exp=new Date(now.getTime()+NGI_ADMIN_SESSION_MS);
    ss.getSheetByName("AdminDevices").appendRow([id,String(p.deviceName||r.deviceName||"Admin Device").slice(0,80),String(p.deviceType||"web"),false,"Active",now.toISOString(),now.toISOString(),ngiDigest_(token,deviceSalt),deviceSalt]);
    const session=ngiToken_(); ss.getSheetByName("AdminSessions").appendRow([Utilities.getUuid(),id,ngiDigest_(session,sessionSalt),sessionSalt,"Active",now.toISOString(),exp.toISOString(),now.toISOString()]);
    return {success:true,isMaster:false,deviceId:id,deviceToken:token,token:session,adminToken:session,expiresAt:exp.toISOString()};
  } finally { lock.releaseLock(); }
}

function ngiRequireAdmin_(ss,p) { const r=ngiValidateAdminSession_(ss,p); if(!r.success) throw new Error("Admin authorization required: "+r.message); return r; }
function ngiAdminLogout_(ss,p) {
  const auth = ngiValidateAdminSession_(ss,p);
  if (!auth.success) return {success:true};
  ngiRows_(ss,"AdminSessions").filter(x=>x.deviceId===p.deviceId && x.status==="Active").forEach(x=>ngiUpdateBy_(ss,"AdminSessions","sessionId",x.sessionId,{status:"Revoked"}));
  return {success:true};
}
function ngiRegisterAdminDevice_(ss,p) {
  const auth=ngiRequireAdmin_(ss,p); if(!auth.isMaster)return {success:false,message:"Only master device can authorize devices."};
  const active=ngiRows_(ss,"AdminDevices").filter(d=>d.status==="Active");
  if(active.length>=NGI_ADMIN_DEVICE_LIMIT)return {success:false,message:"Maximum three active admin devices reached."};
  const id=Utilities.getUuid(),token=ngiToken_(),salt=ngiToken_(),now=new Date(),exp=new Date(now.getTime()+NGI_ADMIN_SESSION_MS);
  const deviceSalt=ngiToken_();
  ss.getSheetByName("AdminDevices").appendRow([id,String(p.deviceName||"Admin Device"),String(p.deviceType||"Unknown"),false,"Active",now.toISOString(),now.toISOString(),ngiDigest_(token,deviceSalt),deviceSalt]);
  ss.getSheetByName("AdminSessions").appendRow([Utilities.getUuid(),id,ngiDigest_(token,salt),salt,"Active",now.toISOString(),exp.toISOString(),now.toISOString()]);
  return {success:true,deviceId:id,deviceToken:token,token:token,adminToken:token,expiresAt:exp.toISOString()};
}
function ngiRevokeAdminDevice_(ss,p) {
  ngiRequireAdmin_(ss,p); const id=String(p.targetDeviceId||p.deviceId||"");
  const d=ngiRows_(ss,"AdminDevices").find(x=>x.deviceId===id);
  if(!d||String(d.isMaster)==="true")return {success:false,message:"Device not found or master device cannot be revoked here."};
  ngiUpdateBy_(ss,"AdminDevices","deviceId",id,{status:"Revoked"});
  ngiRows_(ss,"AdminSessions").filter(s=>s.deviceId===id&&s.status==="Active").forEach(s=>ngiUpdateBy_(ss,"AdminSessions","sessionId",s.sessionId,{status:"Revoked"}));
  return {success:true};
}
function ngiAdminSendEmail_(ss,p) {
  ngiRequireAdmin_(ss,p); const to=String(p.to||"").trim(),subject=String(p.subject||"").trim(),body=String(p.body||"").trim();
  if(!/^[^\\s@]+@[^\\s@]+\\.[^\\s@]+$/.test(to)||!subject||!body)return {success:false,message:"Valid email, subject and message are required."};
  try {MailApp.sendEmail({to:to,subject:subject,body:body,name:"NGI Team"});ss.getSheetByName("EmailLog").appendRow([Utilities.getUuid(),to,subject,new Date().toISOString(),"Sent",""]);return {success:true};}
  catch(e){ss.getSheetByName("EmailLog").appendRow([Utilities.getUuid(),to,subject,new Date().toISOString(),"Failed",e.message]);return {success:false,message:e.message};}
}
function ngiRegisterPushToken_(ss,p) {
  const auth=ngiRequireMember_(ss,p); if(!auth.success)return auth;
  const m=ngiRows_(ss,"Members").find(x=>String(x.memberId).toLowerCase()===String(auth.memberId).toLowerCase());
  const pushToken=p.fcmToken||p.token;
  if(!m||m.status!=="Active"||!pushToken)return {success:false,message:"Active member and push token required."};
  ss.getSheetByName("PushTokens").appendRow([Utilities.getUuid(),m.memberId,String(pushToken),String(p.platform||"web"),"Active",new Date().toISOString(),new Date().toISOString()]);
  return {success:true};
}
function ngiSendAnnouncement_(ss,p) {
  ngiRequireAdmin_(ss,p); const subject=String(p.subject||"").trim(),body=String(p.body||"").trim();
  if(!subject||!body)return {success:false,message:"Subject and message required."};
  let sent=0;
  if(p.emailUsers){ngiRows_(ss,"Members").filter(m=>m.status==="Active"&&m.email).forEach(m=>{try{MailApp.sendEmail({to:m.email,subject:subject,body:body,name:"NGI Team"});sent++;}catch(e){}});}
  return {success:true,emailCount:sent,push:{success:false,message:"Push delivery requires Firebase Cloud Messaging configuration and frontend service worker."}};
}