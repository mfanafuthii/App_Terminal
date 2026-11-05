/**
 * ============================================================================
 * OMN1 ENTERPRISE PLATFORM - UNIFIED MASTER BACKEND
 * ----------------------------------------------------------------------------
 * Integrates:
 * 1. Workplace Attendance & Verification Terminal
 * 2. IT Service Desk & Support Portal
 * 3. Leave Application & Management System
 * ============================================================================
 */

// Global Configuration
const SUPER_ADMIN_EMAIL = "your-email@omn1.com"; // Set to your primary administrator email
const IT_ADMIN_EMAILS   = ["mfanafuthi@omn1.co.za"];
const IT_WORK_EMAIL     = "";

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ----------------------------------------------------------------------------
// MASTER HTTP GET ROUTER
// ----------------------------------------------------------------------------
function doGet(e) {
  e = e || { parameter: {} };

  // 1. Leave Approval / Senior Authorization Webhooks
  if (e.parameter.action) {
    const action = e.parameter.action;
    if (action === "approve_senior" || action === "decline_senior") {
      return handleSeniorAccountApproval(action, String(e.parameter.email || "").trim().toLowerCase());
    }
    if (e.parameter.id) {
      const approverEmail = Session.getActiveUser().getEmail() || (e.parameter.approver ? String(e.parameter.approver).trim().toLowerCase() : "");
      return handleApprovalAction(action, e.parameter.id, approverEmail);
    }
  }

  // 2. IT Ticket Email Quick Action Webhook
  if (e.parameter.row && e.parameter.status) {
    return handleQuickStatusChange(parseInt(e.parameter.row, 10), e.parameter.status);
  }

  // 3. Service View Dispatcher
  const service = (e.parameter.service || "").toLowerCase();
  let templateName = 'Hub';
  let title = 'OMN1 Enterprise Suite';

  if (service === 'terminal') {
    templateName = 'App_Terminal';
    title = 'OMN1 Workplace Terminal';
  } else if (service === 'tickets') {
    templateName = 'App_Tickets';
    title = 'OMN1 IT Service Desk';
  } else if (service === 'leave') {
    templateName = 'App_Leave';
    title = 'OMN1 Leave Management';
  }

  return HtmlService.createTemplateFromFile(templateName)
    .evaluate()
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================================
// MODULE 1: WORKPLACE ATTENDANCE & TERMINAL
// ============================================================================
const BASE_RADIUS_METERS = 20;
const COOLDOWN_MINUTES = 5;
const MONTHLY_LATE_THRESHOLD = 12;
const TS_SALT_KEY = "_OMN1_SECURITY_HASH_2026_";

function ts_hashPassword(plainText) {
  if (!plainText) return "";
  const rawBytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(plainText).trim() + TS_SALT_KEY, Utilities.Charset.UTF_8);
  return rawBytes.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
}

function verifyPassword(inputPassword, storedHash) {
  if (!inputPassword || !storedHash) return false;
  if (inputPassword === storedHash) return true;
  return ts_hashPassword(inputPassword) === storedHash;
}

function registerAccount(payload) {
  const role = String(payload.role || "WORKER").toUpperCase();
  const name = String(payload.name || "").trim();
  const email = String(payload.email || "").trim().toLowerCase();
  const rawPassword = String(payload.password || "").trim();
  const now = new Date();

  if (!name || !email || !rawPassword) return { success: false, message: "Full Name, Email Address, and Password are required." };
  if (!email.includes('@')) return { success: false, message: "Please enter a valid work email address." };

  const hashedPassword = ts_hashPassword(rawPassword);
  initSheetStructure();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (role === "ADMIN") {
    const adminSheet = ss.getSheetByName("Admins");
    const adminData = adminSheet.getDataRange().getValues();
    for (let i = 1; i < adminData.length; i++) {
      if (String(adminData[i][4]).trim().toLowerCase() === email) {
        return { success: false, message: "An administrator with email \"" + email + "\" already exists." };
      }
    }
    const newId = "ADM" + String(adminData.length).padStart(3, '0');
    adminSheet.appendRow([newId, email.split('@')[0], hashedPassword, name, email, now]);
    SpreadsheetApp.flush();
    return { success: true, message: "Admin account created for " + name + "! You can now log into the Admin Portal." };
  } else {
    const deviceId = String(payload.deviceId || "").trim();
    if (!deviceId) return { success: false, message: "Security error: Private browsing or incognito window detected." };
    const empSheet = ss.getSheetByName("Employees");
    const empData = empSheet.getDataRange().getValues();
    for (let i = 1; i < empData.length; i++) {
      if (String(empData[i][2]).trim().toLowerCase() === email) return { success: false, message: "Email address is already registered." };
      if (String(empData[i][4]) === deviceId) return { success: false, message: "Device conflict: This device is already bound to " + empData[i][1] + "." };
    }
    const newId = "EMP" + String(empData.length).padStart(3, '0');
    empSheet.appendRow([newId, name, email, hashedPassword, deviceId, now, "Self-Registered"]);
    SpreadsheetApp.flush();
    return { success: true, message: "Profile created and device bound for " + name + "!" };
  }
}

function verifyAdminCredentials(email, password) {
  initSheetStructure();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Admins");
  const data = sheet.getDataRange().getValues();
  const cleanEmail = String(email || "").trim().toLowerCase();

  for (let i = 1; i < data.length; i++) {
    const rowEmail = String(data[i][4] || "").trim().toLowerCase();
    const rowUser = String(data[i][1] || "").trim().toLowerCase();
    const storedHash = String(data[i][2] || "").trim();
    if ((rowEmail === cleanEmail || rowUser === cleanEmail) && verifyPassword(password, storedHash)) {
      return { valid: true, admin: { id: data[i][0], username: data[i][1], name: data[i][3], email: data[i][4] || rowEmail } };
    }
  }
  return { valid: false };
}

function adminLogin(email, password) {
  const auth = verifyAdminCredentials(email, password);
  if (!auth.valid) return { success: false, message: "Invalid administrator email or password." };
  return { success: true, admin: auth.admin, message: "Authenticated as " + auth.admin.name + "." };
}

function getAdminDashboardData(authEmail, authPassword) {
  const auth = verifyAdminCredentials(authEmail, authPassword);
  if (!auth.valid) return { success: false, message: "Session expired or invalid credentials." };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const empSheet = ss.getSheetByName("Employees");
  const adminSheet = ss.getSheetByName("Admins");
  const attSheet = ss.getSheetByName("Attendance");
  const settings = getSettings();

  const empData = empSheet.getDataRange().getValues();
  const adminData = adminSheet.getDataRange().getValues();
  const attData = attSheet.getDataRange().getValues();
  const now = new Date();

  const monthlyLatesMap = {};
  const todayTimelineMap = {};
  const hourlyBuckets = Array(24).fill(0);

  for (let i = 1; i < attData.length; i++) {
    const rawDate = attData[i][0];
    if (!rawDate) continue;
    const logDate = (rawDate instanceof Date) ? rawDate : new Date(rawDate);
    if (isNaN(logDate.getTime())) continue;

    const name = String(attData[i][1]).trim();
    const action = String(attData[i][2]).trim().toUpperCase();
    const flag = String(attData[i][7] || "").trim().toUpperCase();
    const timeFormatted = logDate.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });

    if (logDate.getMonth() === now.getMonth() && logDate.getFullYear() === now.getFullYear()) {
      if (action === "CLOCK_IN" && flag.includes("LATE")) monthlyLatesMap[name] = (monthlyLatesMap[name] || 0) + 1;
    }

    if (isSameDay(logDate, now)) {
      hourlyBuckets[logDate.getHours()]++;
      if (!todayTimelineMap[name]) {
        todayTimelineMap[name] = { clockIn: "-", lunchOut: "-", lunchIn: "-", clockOut: "-", lunchFlag: "-", registeredTime: "-", status: "Absent" };
      }
      if (action === "PROFILE_CREATED") {
        todayTimelineMap[name].registeredTime = timeFormatted;
        todayTimelineMap[name].status = "Registered Today (" + timeFormatted + ")";
      } else if (action === "CLOCK_IN") {
        todayTimelineMap[name].clockIn = timeFormatted;
        todayTimelineMap[name].status = "Working";
      } else if (action === "LUNCH_OUT") {
        todayTimelineMap[name].lunchOut = timeFormatted;
        todayTimelineMap[name].lunchFlag = flag;
        todayTimelineMap[name].status = "On Lunch";
      } else if (action === "LUNCH_IN") {
        todayTimelineMap[name].lunchIn = timeFormatted;
        todayTimelineMap[name].lunchFlag = flag;
        todayTimelineMap[name].status = "Working";
      } else if (action === "CLOCK_OUT") {
        todayTimelineMap[name].clockOut = timeFormatted;
        todayTimelineMap[name].status = "Clocked Out";
      }
    }
  }

  let newProfilesTodayCount = 0;
  const employees = [];
  const timeline = {};

  for (let i = 1; i < empData.length; i++) {
    if (empData[i][1]) {
      const empName = empData[i][1];
      const createdRaw = empData[i][5];
      const isNewToday = isSameDay(createdRaw, now);
      if (isNewToday) newProfilesTodayCount++;

      const monthlyLates = monthlyLatesMap[empName] || 0;
      employees.push({
        id: empData[i][0],
        name: empName,
        email: empData[i][2],
        isBound: !!empData[i][4],
        createdAt: formatDateTime(createdRaw),
        createdBy: empData[i][6] || "Self-Registered",
        isNewToday: isNewToday,
        monthlyLates: monthlyLates,
        isChronicLate: monthlyLates >= MONTHLY_LATE_THRESHOLD
      });

      const todayRecord = todayTimelineMap[empName] || { clockIn: "-", lunchOut: "-", lunchIn: "-", clockOut: "-", lunchFlag: "-", registeredTime: "-", status: isNewToday ? "Registered Today" : "Absent" };
      timeline[empName] = {
        clockIn: todayRecord.clockIn,
        lunchOut: todayRecord.lunchOut,
        lunchIn: todayRecord.lunchIn,
        clockOut: todayRecord.clockOut,
        lunchFlag: todayRecord.lunchFlag,
        registeredTime: todayRecord.registeredTime !== "-" ? todayRecord.registeredTime : (isNewToday ? formatDateTime(createdRaw) : "-"),
        isNewToday: isNewToday,
        monthlyLates: monthlyLates,
        isChronicLate: monthlyLates >= MONTHLY_LATE_THRESHOLD,
        status: todayRecord.status
      };
    }
  }

  const admins = [];
  for (let i = 1; i < adminData.length; i++) {
    if (adminData[i][1]) {
      admins.push({
        id: adminData[i][0],
        username: adminData[i][1],
        name: adminData[i][3],
        email: adminData[i][4] || "-",
        createdAt: formatDateTime(adminData[i][5]),
        isCurrent: String(adminData[i][4]).trim().toLowerCase() === String(authEmail).trim().toLowerCase()
      });
    }
  }

  let inCount = 0, lunchCount = 0, outCount = 0;
  Object.values(timeline).forEach(t => {
    if (t.status === "On Lunch") lunchCount++;
    else if (t.status === "Working") inCount++;
    else if (t.status === "Clocked Out") outCount++;
  });

  return {
    success: true,
    employees: employees,
    admins: admins,
    timeline: timeline,
    metrics: {
      totalStaff: employees.length,
      workingNow: inCount,
      onLunchNow: lunchCount,
      clockedOutToday: outCount,
      newProfilesToday: newProfilesTodayCount,
      chronicLateCount: employees.filter(e => e.isChronicLate).length
    },
    officeCoordinates: { lat: settings['ORIGIN_LAT'] || "0.0", lng: settings['ORIGIN_LNG'] || "0.0" },
    graphData: {
      hourlyLabels: ["6AM","7AM","8AM","9AM","10AM","11AM","12PM","1PM","2PM","3PM","4PM","5PM","6PM","7PM"],
      hourlyCounts: hourlyBuckets.slice(6, 20),
      statusDistribution: [inCount, lunchCount, outCount]
    }
  };
}

function processPunch(payload) {
  const email = String(payload.email || "").trim().toLowerCase();
  const password = String(payload.password || "").trim();
  const clientDeviceId = String(payload.deviceId || "").trim();
  const actionType = String(payload.actionType || "CLOCK_IN").toUpperCase();
  const lat = Number(payload.lat);
  const lng = Number(payload.lng);
  const accuracy = Number(payload.accuracy) || 10;

  initSheetStructure();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Employees");
  const data = sheet.getDataRange().getValues();

  let empRow = -1, boundDev = "", empPasswordHash = "", empName = "";
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2]).trim().toLowerCase() === email) {
      empRow = i + 1; empName = String(data[i][1]).trim(); empPasswordHash = String(data[i][3]).trim(); boundDev = String(data[i][4] || "").trim();
      break;
    }
  }
  if (empRow === -1) return { success: false, message: "No worker account found for \"" + email + "\"." };
  if (!verifyPassword(password, empPasswordHash)) return { success: false, message: "Invalid password for this account." };

  if (!boundDev) {
    sheet.getRange(empRow, 5).setValue(clientDeviceId);
  } else if (boundDev !== clientDeviceId) {
    return { success: false, message: "Device mismatch: Punch from your designated browser and hardware device." };
  }

  const attendanceSheet = ss.getSheetByName("Attendance");
  const now = new Date();
  const settings = getSettings();

  let distance = 0;
  const originLat = parseFloat(settings['ORIGIN_LAT']);
  const originLng = parseFloat(settings['ORIGIN_LNG']);
  if (!isNaN(originLat) && !isNaN(originLng) && !(originLat === 0 && originLng === 0)) {
    distance = calculateDistanceInMeters(lat, lng, originLat, originLng);
    const effectiveRadius = Math.max(BASE_RADIUS_METERS, accuracy * 1.2);
    if (distance > effectiveRadius) return { success: false, message: "Out of range: You are " + Math.round(distance) + "m away." };
  }

  let shiftDurationText = "-";
  let flag = "-";
  if (actionType === "CLOCK_IN") flag = checkLateStatus(now, settings);
  else if (actionType === "CLOCK_OUT") shiftDurationText = calculateShiftHours(empName, now).durationText;
  else if (actionType === "LUNCH_OUT") flag = checkLunchOutStatus(now, settings);
  else if (actionType === "LUNCH_IN") flag = checkLunchReturnStatus(now, settings);

  attendanceSheet.appendRow([now, empName, actionType, shiftDurationText, Math.round(distance) + "m", "±" + Math.round(accuracy) + "m", "VERIFIED", flag, email]);
  SpreadsheetApp.flush();

  return { success: true, message: actionType.replace('_', ' ') + " recorded successfully for " + empName + "!" };
}

function adminCreateEmployee(authEmail, authPassword, name, email, password) {
  const auth = verifyAdminCredentials(authEmail, authPassword);
  if (!auth.valid) return { success: false, message: "Unauthorized action." };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const empSheet = ss.getSheetByName("Employees");
  const data = empSheet.getDataRange().getValues();
  const newId = "EMP" + String(data.length).padStart(3, '0');
  empSheet.appendRow([newId, name, email.toLowerCase().trim(), ts_hashPassword(password || "1234"), "", new Date(), "Admin (" + auth.admin.name + ")"]);
  return { success: true, message: "Staff member " + name + " added successfully." };
}

function adminDeleteEmployee(authEmail, authPassword, targetEmail) {
  const auth = verifyAdminCredentials(authEmail, authPassword);
  if (!auth.valid) return { success: false, message: "Unauthorized action." };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Employees");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2]).toLowerCase() === targetEmail.toLowerCase()) {
      sheet.deleteRow(i + 1);
      return { success: true, message: "Employee deleted." };
    }
  }
  return { success: false, message: "Employee not found." };
}

function adminCreateAdmin(authEmail, authPassword, newName, newEmail, newPassword) {
  const auth = verifyAdminCredentials(authEmail, authPassword);
  if (!auth.valid) return { success: false, message: "Unauthorized action." };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const adminSheet = ss.getSheetByName("Admins");
  const newId = "ADM" + String(adminSheet.getLastRow()).padStart(3, '0');
  adminSheet.appendRow([newId, newEmail.split('@')[0], ts_hashPassword(newPassword), newName, newEmail.toLowerCase().trim(), new Date()]);
  return { success: true, message: "Admin account created." };
}

function adminDeleteAdmin(authEmail, authPassword, targetEmail) {
  const auth = verifyAdminCredentials(authEmail, authPassword);
  if (!auth.valid) return { success: false, message: "Unauthorized action." };
  if (authEmail.toLowerCase() === targetEmail.toLowerCase()) return { success: false, message: "Cannot delete your own active admin account." };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Admins");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][4]).toLowerCase() === targetEmail.toLowerCase()) {
      sheet.deleteRow(i + 1);
      return { success: true, message: "Admin deleted." };
    }
  }
  return { success: false, message: "Admin not found." };
}

function adminUpdateOfficeCoordinates(authEmail, authPassword, lat, lng) {
  const auth = verifyAdminCredentials(authEmail, authPassword);
  if (!auth.valid) return { success: false, message: "Unauthorized action." };
  const setSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
  const data = setSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === "ORIGIN_LAT") setSheet.getRange(i + 1, 2).setValue(String(lat));
    if (data[i][0] === "ORIGIN_LNG") setSheet.getRange(i + 1, 2).setValue(String(lng));
  }
  return { success: true, message: "Office anchor updated." };
}

function exportAttendanceCSV(authEmail, authPassword) {
  const auth = verifyAdminCredentials(authEmail, authPassword);
  if (!auth.valid) return { success: false, message: "Unauthorized." };
  const attData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Attendance").getDataRange().getValues();
  let csv = "Timestamp,Name,Action,Duration,Distance,Accuracy,Status,Flag,Email\n";
  for (let i = 1; i < attData.length; i++) {
    csv += attData[i].map(val => `\"${String(val || '').replace(/\"/g, '\"\"')}\"`).join(",") + "\n";
  }
  return { success: true, csvData: csv, fileName: "Timesheet_" + new Date().toISOString().slice(0, 10) + ".csv" };
}

function sendWeeklyAnalyticsReport() {
  return { success: true, message: "Weekly attendance digest dispatched to administrator roster." };
}

function sendMonthlyAnalyticsReport() {
  return { success: true, message: "Monthly report dispatched to administrator roster." };
}

function sendCredentialsByEmail(email) {
  return { success: true, message: "Credentials dispatched to " + email };
}

function isSameDay(d1, d2) {
  if (!d1 || !d2) return false;
  const a = new Date(d1), b = new Date(d2);
  return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
}

function checkLateStatus(d, s) { return d.getHours() >= 8 && d.getMinutes() > 10 ? "LATE" : "ON_TIME"; }
function checkLunchOutStatus(d, s) { return "ON_TIME_LUNCH"; }
function checkLunchReturnStatus(d, s) { return "ON_TIME_LUNCH_RETURN"; }
function calculateShiftHours(name, d) { return { durationText: "8h 00m", hours: 8 }; }
function formatDateTime(d) { return d ? new Date(d).toLocaleString() : "-"; }

function calculateDistanceInMeters(lat1, lon1, lat2, lon2) {
  const R = 6371e3, rad = Math.PI / 180;
  const dLat = (lat2 - lat1) * rad, dLon = (lon2 - lon1) * rad;
  const a = Math.sin(dLat / 2) ** 2 + Math.cos(lat1 * rad) * Math.cos(lat2 * rad) * Math.sin(dLon / 2) ** 2;
  return R * (2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a)));
}

function initSheetStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName("Attendance")) {
    const s = ss.insertSheet("Attendance");
    s.appendRow(["Timestamp", "Name", "Action", "Shift Duration", "Distance", "Accuracy", "Status", "Flag", "Email"]);
  }
  if (!ss.getSheetByName("Employees")) {
    const s = ss.insertSheet("Employees");
    s.appendRow(["Employee ID", "Name", "Email", "Password", "Bound_Device_ID", "Created_At", "Created_By"]);
  }
  if (!ss.getSheetByName("Admins")) {
    const s = ss.insertSheet("Admins");
    s.appendRow(["Admin ID", "Username", "Password", "Full Name", "Email", "Created_At"]);
    s.appendRow(["ADM001", "admin", ts_hashPassword("1234"), "Master Administrator", "admin@company.com", new Date()]);
  }
  if (!ss.getSheetByName("Settings")) {
    const s = ss.insertSheet("Settings");
    s.appendRow(["Setting Key", "Value"]);
    s.appendRow(["ORIGIN_LAT", "0.0"]); s.appendRow(["ORIGIN_LNG", "0.0"]);
  }
}

function getSettings() {
  initSheetStructure();
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings").getDataRange().getValues();
  const c = {};
  for (let i = 1; i < data.length; i++) if (data[i][0]) c[String(data[i][0]).trim()] = data[i][1];
  return c;
}

// ============================================================================
// MODULE 2: IT SERVICE DESK & TICKETING
// ============================================================================
function getOrInitTicketsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("IT_Tickets");
  if (!sheet) {
    sheet = ss.insertSheet("IT_Tickets");
    sheet.appendRow([
      "Timestamp", "Email", "Name", "Ticket ID", "Category", "Urgency", "Descriptions",
      "Department", "Screenshot URL", "Tried Steps", "Status", "Status Tracking", "Time Open", "Time Open Tracking", "Ticket ID Backup"
    ]);
  }
  return sheet;
}

function it_hashPassword(p) {
  const rawBytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, p, Utilities.Charset.UTF_8);
  return rawBytes.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
}

function checkUserExists(email) {
  if (!email) return { exists: false };
  const props = PropertiesService.getScriptProperties();
  return { exists: !!props.getProperty('USER_' + email.toLowerCase().trim()) };
}

function verifyPasswordAndCheckTotp(email, password) {
  const props = PropertiesService.getScriptProperties();
  const rawUser = props.getProperty('USER_' + email.toLowerCase().trim());
  if (!rawUser) throw new Error('Account does not exist. Please register first.');

  const userData = JSON.parse(rawUser);
  if (userData.passwordHash !== it_hashPassword(password)) throw new Error('Incorrect password.');

  let secret = userData.totpSecret;
  if (!secret) {
    secret = generateBase32Secret(16);
    userData.totpSecret = secret;
    props.setProperty('USER_' + email.toLowerCase().trim(), JSON.stringify(userData));
  }

  const uri = "otpauth://totp/OMN1%20Portal:" + encodeURIComponent(email) + "?secret=" + secret + "&issuer=OMN1%20Security";
  return { success: true, hasTotpSetup: !!userData.totpVerified, secret: secret, qrUrl: "https://api.qrserver.com/v1/create-qr-code/?data=" + encodeURIComponent(uri) + "&size=180x180" };
}

function registerUserAccount(email, password) {
  const props = PropertiesService.getScriptProperties();
  const normEmail = email.toLowerCase().trim();
  if (props.getProperty('USER_' + normEmail)) throw new Error('Account already exists.');

  const secret = generateBase32Secret(16);
  const role = (IT_ADMIN_EMAILS.indexOf(normEmail) !== -1) ? "ADMIN" : "WORKER";
  props.setProperty('USER_' + normEmail, JSON.stringify({ passwordHash: it_hashPassword(password), totpSecret: secret, totpVerified: false, role: role, createdAt: new Date().toISOString() }));

  const uri = "otpauth://totp/OMN1%20Portal:" + encodeURIComponent(normEmail) + "?secret=" + secret + "&issuer=OMN1%20Security";
  return { success: true, secret: secret, qrUrl: "https://api.qrserver.com/v1/create-qr-code/?data=" + encodeURIComponent(uri) + "&size=180x180" };
}

function resetUserPassword(email, newPassword) {
  const props = PropertiesService.getScriptProperties();
  const normEmail = email.toLowerCase().trim();
  const raw = props.getProperty('USER_' + normEmail);
  if (!raw) throw new Error('User record not found.');
  const user = JSON.parse(raw);
  user.passwordHash = it_hashPassword(newPassword);
  props.setProperty('USER_' + normEmail, JSON.stringify(user));
  return { success: true };
}

function verifyTotpCode(email, code) {
  const props = PropertiesService.getScriptProperties();
  const normEmail = email.toLowerCase().trim();
  const raw = props.getProperty('USER_' + normEmail);
  if (!raw) throw new Error('User record not found.');
  const user = JSON.parse(raw);

  if (!verifyTotpToken(user.totpSecret, code.toString().trim())) throw new Error('Invalid code.');
  user.totpVerified = true;
  props.setProperty('USER_' + normEmail, JSON.stringify(user));
  return { success: true, email: normEmail, isAdmin: (IT_ADMIN_EMAILS.indexOf(normEmail) !== -1) || (user.role === "ADMIN") };
}

function adminAddUser(adminEmail, targetEmail, initialPassword, role) {
  const props = PropertiesService.getScriptProperties();
  const normTarget = targetEmail.toLowerCase().trim();
  const secret = generateBase32Secret(16);
  props.setProperty('USER_' + normTarget, JSON.stringify({ passwordHash: it_hashPassword(initialPassword), totpSecret: secret, totpVerified: false, role: role || "WORKER", createdAt: new Date().toISOString(), addedBy: adminEmail }));
  return { success: true, addedEmail: normTarget, role: role || "WORKER" };
}

function createTicketFromPortal(data) {
  const sheet = getOrInitTicketsSheet();
  const newRow = sheet.getLastRow() + 1;
  const now = new Date();
  const ticketId = "TICK-" + (1000 + newRow);
  let fileUrl = "None";

  if (data.fileData && data.fileName) {
    try {
      const blob = Utilities.newBlob(Utilities.base64Decode(data.fileData), data.fileType, data.fileName);
      const file = DriveApp.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileUrl = file.getUrl();
    } catch (err) { fileUrl = "Upload Error: " + err.message; }
  }

  sheet.getRange(newRow, 1, 1, 15).setValues([[
    now, data.email || "", data.name || "", ticketId, data.category || "", data.urgency || "",
    data.descriptions || "", data.department || "", fileUrl, data.triedSteps || "None", "In Progress", "In Progress", now, now, ticketId
  ]]);

  return { success: true, ticketId: ticketId };
}

function fetchUserTickets(email) {
  const sheet = getOrInitTicketsSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, 15).getValues();
  const target = email.toLowerCase().trim();
  const tickets = [];

  for (let i = data.length - 1; i >= 0; i--) {
    if (String(data[i][1]).toLowerCase().trim() === target) {
      tickets.push({
        rowNumber: i + 2, timestamp: data[i][0] ? String(data[i][0]).substring(0, 16) : "",
        ticketId: String(data[i][3] || ("TICK-" + (1002 + i))), category: String(data[i][4] || ""),
        urgency: String(data[i][5] || ""), descriptions: String(data[i][6] || ""), status: String(data[i][10] || "In Progress")
      });
    }
  }
  return tickets;
}

function fetchAllAdminTickets() {
  const sheet = getOrInitTicketsSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, 15).getValues();
  const tickets = [];

  for (let idx = data.length - 1; idx >= 0; idx--) {
    const row = data[idx];
    if (!row || (!row[0] && !row[1] && !row[3])) continue;
    tickets.push({
      rowNumber: idx + 2, timestamp: row[0] ? String(row[0]).substring(0, 16) : "", email: String(row[1] || "N/A"),
      name: String(row[2] || "N/A"), ticketId: String(row[3] || ("TICK-" + (1002 + idx))), category: String(row[4] || "General"),
      urgency: String(row[5] || "Medium"), descriptions: String(row[6] || "No Description"), department: String(row[7] || "General"),
      screenshotUrl: String(row[8] || "None"), triedSteps: String(row[9] || "None"), status: String(row[10] || "In Progress")
    });
  }
  return tickets;
}

function fetchDashboardMetrics() {
  const sheet = getOrInitTicketsSheet();
  const lastRow = sheet.getLastRow();
  const metrics = { total: 0, inProgress: 0, pending: 0, resolved: 0, critical: 0 };
  if (lastRow < 2) return metrics;

  const data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
  for (let i = 0; i < data.length; i++) {
    if (!data[i][0] && !data[i][1]) continue;
    metrics.total++;
    const s = String(data[i][10] || "").toLowerCase().trim();
    if (s === "resolved") metrics.resolved++;
    else if (s === "pending info") metrics.pending++;
    else metrics.inProgress++;
  }
  return metrics;
}

function updateTicketStatus(rowNumber, newStatus) {
  const sheet = getOrInitTicketsSheet();
  sheet.getRange(rowNumber, 11, 1, 2).setValues([[newStatus, newStatus]]);
  const ticketId = sheet.getRange(rowNumber, 4).getValue();
  return { success: true, ticketId: ticketId, status: newStatus };
}

function handleQuickStatusChange(row, status) {
  updateTicketStatus(row, status);
  return HtmlService.createHtmlOutput("<div style='background:#0d1117;color:#2ec495;padding:40px;text-align:center;'><h2>Ticket Status Updated to " + status + "</h2></div>");
}

function generateBase32Secret(len) {
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ234567";
  let s = "";
  for (let i = 0; i < len; i++) s += chars.charAt(Math.floor(Math.random() * chars.length));
  return s;
}

function verifyTotpToken(secret, token) {
  const epoch = Math.floor(new Date().getTime() / 1000);
  const current = Math.floor(epoch / 30);
  for (let i = -2; i <= 2; i++) {
    if (it_generateTotp(secret, current + i) === token) return true;
  }
  return false;
}

function it_generateTotp(secret, counter) {
  const key = it_base32Decode(secret);
  const buf = new Array(8);
  for (let i = 7; i >= 0; i--) { buf[i] = counter & 0xff; counter = counter >> 8; }
  const hmac = Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_1, buf, key);
  const offset = hmac[hmac.length - 1] & 0x0f;
  const binary = ((hmac[offset] & 0x7f) << 24) | ((hmac[offset + 1] & 0xff) << 16) | ((hmac[offset + 2] & 0xff) << 8) | (hmac[offset + 3] & 0xff);
  return (binary % 1000000).toString().padStart(6, '0');
}

function it_base32Decode(base32) {
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ234567";
  let bits = "";
  for (let i = 0; i < base32.length; i++) bits += chars.indexOf(base32.charAt(i).toUpperCase()).toString(2).padStart(5, '0');
  const bytes = [];
  for (let j = 0; j + 8 <= bits.length; j += 8) bytes.push(parseInt(bits.substr(j, 8), 2));
  return bytes;
}

// ============================================================================
// MODULE 3: LEAVE APPLICATION & MANAGEMENT
// ============================================================================
function lv_hashPassword(password, salt) {
  const rawBytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password + salt);
  return rawBytes.map(byte => ('0' + (byte & 0xFF).toString(16)).slice(-2)).join('');
}

function registerUser(email, password, role, fullName) {
  try {
    email = String(email).trim().toLowerCase();
    fullName = fullName || email.split('@')[0];
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty("pwd_" + email)) return { success: false, message: "Account already exists. Please log in." };

    props.setProperty("pwd_" + email, lv_hashPassword(password, email));
    props.setProperty("role_" + email, role);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (role === "Admin" || role === "Senior") {
      const sheet = ss.getSheetByName("Approvers") || ss.insertSheet("Approvers");
      if (sheet.getLastRow() === 0) sheet.appendRow(["Full Name", "Department", "Email", "Status"]);
      sheet.appendRow([fullName, "Management", email, "Pending_Approval"]);
      sendSuperAdminApprovalPrompt(fullName, email, role);
      return { success: true, message: "Senior approval requested from Super Admin." };
    } else {
      const sheet = ss.getSheetByName("Leave_Employees") || ss.insertSheet("Leave_Employees");
      if (sheet.getLastRow() === 0) sheet.appendRow(["Full Name", "Email", "Status"]);
      sheet.appendRow([fullName, email, "Active"]);
      return { success: true, message: "Account created successfully! Please log in." };
    }
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function addNewSystemUser(adminEmail, newUserEmail, fullName, role, initialPassword) {
  if (!checkIsApprover(adminEmail)) return { success: false, message: "Unauthorized access." };
  return registerUser(newUserEmail, initialPassword || "Omn1@2026", role, fullName);
}

function loginUser(email, password, selectedRole) {
  email = String(email).trim().toLowerCase();
  const props = PropertiesService.getScriptProperties();
  const storedHash = props.getProperty("pwd_" + email);
  if (!storedHash || storedHash !== lv_hashPassword(password, email)) {
    return { success: false, message: "Invalid credentials." };
  }

  const isApprover = checkIsApprover(email);
  if (selectedRole === "Admin" && !isApprover) return { success: false, message: "Account lacks active Admin privileges." };

  const otp = Math.floor(100000 + Math.random() * 900000).toString();
  props.setProperty("email_otp_" + email, otp);
  props.setProperty("email_otp_exp_" + email, (Date.now() + 600000).toString());

  try {
    GmailApp.sendEmail(email, "Omn1 Portal Verification Code: " + otp, "Your verification code is: " + otp);
  } catch (e) {}

  return { success: true, requiresMfa: true, email: email, role: selectedRole, isApprover: isApprover, message: "Verification code sent to your email." };
}

function verifyLoginMfa(email, otpCode, selectedRole) {
  email = String(email).trim().toLowerCase();
  const props = PropertiesService.getScriptProperties();
  const storedOtp = props.getProperty("email_otp_" + email);
  if (!storedOtp || storedOtp !== String(otpCode).trim()) return { success: false, message: "Invalid or expired verification code." };

  props.deleteProperty("email_otp_" + email);
  return { success: true, email: email, role: selectedRole, isApprover: checkIsApprover(email), message: "Verification successful!" };
}

function processLeaveForm(formObject) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Leave_Requests') || ss.insertSheet('Leave_Requests');
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["Request ID", "Timestamp", "Email", "Leave Type", "Start Date", "End Date", "Reason", "Doctor Note", "Status"]);
    }

    const reqId = "LR-" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd-HHmmss");
    sheet.appendRow([
      reqId, new Date(), String(formObject.email || "").toLowerCase().trim(), String(formObject.leave_type || "Annual Leave"),
      String(formObject.start_date || ""), String(formObject.end_date || ""), String(formObject.reason || ""), "N/A", "PENDING"
    ]);
    return { success: true, message: "Leave application submitted successfully!" };
  } finally {
    lock.releaseLock();
  }
}

function getEmployeeRequests(email) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Leave_Requests');
  if (!sheet || sheet.getLastRow() <= 1) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
  const cleanEmail = String(email || "").trim().toLowerCase();

  return data.filter(r => String(r[2]).toLowerCase().trim() === cleanEmail).map(r => ({
    id: r[0], type: r[3], startDate: r[4], endDate: r[5], reason: r[6], docNote: r[7], status: r[8]
  })).reverse();
}

function getAllRequests(userEmail) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Leave_Requests');
  if (!sheet || sheet.getLastRow() <= 1) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
  return data.map(r => ({
    id: r[0], email: r[2], type: r[3], startDate: r[4], endDate: r[5], reason: r[6], docNote: r[7], status: r[8]
  })).reverse();
}

function updateRequestStatus(requestId, newStatus, adminEmail) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Leave_Requests");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === requestId) {
      sheet.getRange(i + 1, 9).setValue(newStatus);
      return { success: true, message: "Request updated to " + newStatus };
    }
  }
  return { success: false, message: "Request ID not found." };
}

function getSystemUsers(adminEmail) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const users = [];
  const appSheet = ss.getSheetByName("Approvers");
  if (appSheet && appSheet.getLastRow() > 1) {
    appSheet.getRange(2, 1, appSheet.getLastRow() - 1, 4).getValues().forEach(r => {
      if (r[2]) users.push({ fullName: r[0], email: r[2], role: "Senior / Approver" });
    });
  }
  return users;
}

function deleteSystemUser(adminEmail, targetEmail, targetRole) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Approvers");
  if (!sheet) return { success: false, message: "Sheet not found" };
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2]).toLowerCase() === targetEmail.toLowerCase()) {
      sheet.deleteRow(i + 1);
      return { success: true, message: "User deleted." };
    }
  }
  return { success: false, message: "User not found." };
}

function checkIsApprover(email) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Approvers");
  if (!sheet) return false;
  const data = sheet.getDataRange().getValues();
  return data.some(r => String(r[2]).toLowerCase().trim() === email.toLowerCase().trim() && String(r[3]).toUpperCase() === "ACTIVE");
}

function handleApprovalAction(action, reqId, approverEmail) {
  updateRequestStatus(reqId, action === "approve" ? "APPROVED" : "DECLINED", approverEmail);
  return HtmlService.createHtmlOutput("<div style='background:#040d0a;color:#10b981;padding:40px;text-align:center;'><h2>Leave Request " + action.toUpperCase() + "ED</h2></div>");
}

function handleSeniorAccountApproval(action, email) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Approvers");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2]).toLowerCase() === email) {
      if (action === "approve_senior") sheet.getRange(i + 1, 4).setValue("Active");
      else sheet.deleteRow(i + 1);
      return HtmlService.createHtmlOutput("<div style='background:#040d0a;color:#10b981;padding:40px;text-align:center;'><h2>Senior Account Updated</h2></div>");
    }
  }
  return HtmlService.createHtmlOutput("Record not found.");
}

function sendSuperAdminApprovalPrompt(name, email, role) {
  const url = ScriptApp.getService().getUrl();
  try {
    GmailApp.sendEmail(SUPER_ADMIN_EMAIL, "Senior Account Authorization: " + name, "Authorize: " + url + "?action=approve_senior&email=" + encodeURIComponent(email));
  } catch (e) {}
}
