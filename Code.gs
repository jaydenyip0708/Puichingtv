// Puiching TV 專案管理系統 - 後端 Code.gs

var SPREADSHEET_ID = "18Dae4Bw0hJwwHVXOWTYpvWML0fI0Sjx5n44aAoMjJeg";

function sheetTimeZone_() {
  return "Asia/Hong_Kong";
}

function getSpreadsheet() {
  if (SPREADSHEET_ID) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Puiching TV 影片管理系統')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function logRevisionHistory(actionType, title, versionType, details, notes) {
  try {
    var ss = getSpreadsheet();
    var revSheet = ss.getSheetByName("Revision History") || ss.getSheetByName("修訂紀錄");
    if (!revSheet) {
      revSheet = ss.insertSheet("Revision History");
      revSheet.appendRow(["異動時間", "操作者  Email", "操作者名字", "動作類型", "節目名稱", "版本類型", "變更詳情", "備註"]);
    }
    var userEmail = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || "未知使用者";
    var userName = "查無此帳號對應資料";
    var memberSheet = ss.getSheetByName("Member") || ss.getSheetByName("成員");
    if (memberSheet && userEmail && userEmail !== "未知使用者") {
      var memberData = memberSheet.getDataRange().getValues();
      for (var i = 1; i < memberData.length; i++) {
        var account = String(memberData[i][1]).trim().toLowerCase();
        if (account && account === String(userEmail).trim().toLowerCase()) {
          userName = memberData[i][0];
          break;
        }
      }
    }
    var timestamp = Utilities.formatDate(new Date(), sheetTimeZone_(), "yyyy-MM-dd HH:mm:ss");
    revSheet.insertRowBefore(2);
    revSheet.getRange(2, 1, 1, 8).setValues([[timestamp, userEmail, userName, actionType || "", title || "", versionType || "", details || "", notes || ""]]);
  } catch (e) {
    Logger.log("logRevisionHistory 錯誤 :" + e.toString());
  }
}

function recordLogin() {
  try {
    var ss = getSpreadsheet();
    var loginSheet = ss.getSheetByName("Login Record") || ss.getSheetByName("登入紀錄");
    if (!loginSheet) {
      loginSheet = ss.insertSheet("Login Record");
      loginSheet.appendRow(["登入時間", "使用者  Email", "名字"]);
    }
    var userEmail = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || "未知使用者";
    var matchedAValue = "查無此帳號對應資料";
    var memberSheet = ss.getSheetByName("Member") || ss.getSheetByName("成員");
    if (memberSheet && userEmail && userEmail !== "未知使用者") {
      var memberData = memberSheet.getDataRange().getValues();
      for (var i = 1; i < memberData.length; i++) {
        var memberAccount = String(memberData[i][1]).trim();
        if (memberAccount && memberAccount.toLowerCase() === String(userEmail).trim().toLowerCase()) {
          matchedAValue = memberData[i][0];
          break;
        }
      }
    }
    var timestamp = Utilities.formatDate(new Date(), sheetTimeZone_(), "yyyy-MM-dd HH:mm:ss");
    if (loginSheet.getLastRow() < 1) {
      loginSheet.appendRow(["登入時間", "使用者  Email", "名字"]);
    }
    loginSheet.insertRowAfter(1);
    loginSheet.getRange(2, 1, 1, 3).setNumberFormats([["@", "@", "@"]]);
    loginSheet.getRange(2, 1, 1, 3).setValues([[timestamp, userEmail, matchedAValue]]);
    SpreadsheetApp.flush();
    return { success: true, email: userEmail, memberInfo: matchedAValue };
  } catch (error) {
    Logger.log("recordLogin 錯誤 :" + error.toString());
    return { success: false, message: error.toString() };
  }
}


function getAdminNames() {
  try {
    var sheet = getAdminSheet_();
    if (!sheet || sheet.getLastRow() < 2) return [];
    var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    return values.map(function(r) { return String(r[0] || "").trim(); }).filter(String);
  } catch (e) {
    return [];
  }
}

function getInitialData() {
  var result = { categories: [], members: [], projects: [], login: {}, admin: {}, activities: [], adminNames: [] };
  try { result.categories = getCategories(); } catch (e) { Logger.log("categories: " + e); }
  try { result.members = getProjectManagers(); } catch (e) { Logger.log("members: " + e); }
  try { result.projects = getProjects(); } catch (e) { Logger.log("projects: " + e); }
  try { result.login = recordLogin(); } catch (e) { Logger.log("login: " + e); }
  try { result.admin = checkAdminAccess(); } catch (e) { Logger.log("admin: " + e); }
  try {
    if (result.admin && result.admin.isAdmin) {
      var act = getActivities();
      result.activities = (act && act.success) ? (act.data || []) : [];
      result.adminNames = getAdminNames();
    }
  } catch (e) { Logger.log("activities: " + e); }
  return result;
}

function getCategories() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("Categories") || ss.getSheetByName("類別");
  if (!sheet) return ["小城遊記", "培正 Talkshow", "年少有你", "走進中華大觀園", "紅藍檔案", "校園我 say", "案件重演", "動力「培」FUN", "講呢 D", "1分鐘就有60秒過去"];
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  return values.map(function(r) { return r[0]; }).filter(String);
}

function getProjectManagers() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("Member") || ss.getSheetByName("成員");
  if (!sheet) return [];
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  return values.map(function(r) { return r[0]; }).filter(String);
}

function getMembers() {
  return getProjectManagers();
}

function getProjects() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("Projects") || ss.getSheetByName("專案");
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
  var result = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    result.push({
      rowIndex: i + 2,
      title: row[0] || "",
      category: row[1] || "",
      leader: row[2] || "",
      editor: row[3] || "",
      driveUrl: row[4] || "",
      normalStatus: row[5] || "待剪接",
      normalTargetDate: formatDate(row[6]),
      needTdm: row[7] || "否",
      tdmStatus: row[8] || "待剪接",
      tdmTargetDate: formatDate(row[9]),
      notes: row[10] || ""
    });
  }
  return result;
}

function getProjectsByMonth(yearMonth) {
  var all = getProjects();
  if (!yearMonth) return all;
  return all.filter(function(p) {
    var nd = (p.normalTargetDate || "").substring(0, 7);
    var td = (p.tdmTargetDate || "").substring(0, 7);
    return nd === yearMonth || td === yearMonth;
  });
}

function updateEventDate(rowIndex, versionType, newDate) {
  try {
    var sheet = getSpreadsheet().getSheetByName("Projects") || getSpreadsheet().getSheetByName("專案");
    if (!sheet) return { success: false, message: "找不到工作表" };
    rowIndex = Number(rowIndex);
    if (!rowIndex || rowIndex < 2) return { success: false, message: "列號不正確" };
    var title = sheet.getRange(rowIndex, 1).getValue();
    if (versionType === "tdm") {
      sheet.getRange(rowIndex, 10).setValue(newDate);
    } else {
      sheet.getRange(rowIndex, 7).setValue(newDate);
    }
    logRevisionHistory("拖曳改期", title, versionType === "tdm" ? "TDM" : "IG", "日期改為 " + newDate, "");
    return { success: true, message: "日期已更新" };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function addProject(formObject) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName("Projects") || ss.getSheetByName("專案");
    if (!sheet) {
      sheet = ss.insertSheet("Projects");
      sheet.appendRow(["活動 /節目名稱", "節目類別", "專案負責人", "剪接師", "專案 Drive 連結", "校園版進度", "普通版預計上架日", "需要 TDM版本", "TDM版進度", "TDM版預計交片日", "備註說明"]);
    }
    var title = formObject.title;
    var category = formObject.category;
    var leader = formObject.leader;
    var editor = formObject.editor;
    var driveUrl = formObject.driveUrl || "";
    var targetDate = formObject.targetDate;
    var needTdm = formObject.needTdm ? "是" : "否";
    var tdmDate = formObject.tdmDate || "";
    var notes = formObject.notes || "";
    var tdmStatus = (needTdm === "是") ? "待剪接" : "不需要";
    sheet.appendRow([title, category, leader, editor, driveUrl, "待剪接", targetDate, needTdm, tdmStatus, tdmDate, notes]);
    var details = "新增專案  (預計上架 :" + targetDate + (needTdm === "是" ? ", TDM交片 :" + tdmDate : "") + ")";
    logRevisionHistory("新增項目", title, "建立專案", details, notes);
    return { success: true, message: "專案建立成功 ！" };
  } catch (e) {
    return { success: false, message: "錯誤 ：" + e.toString() };
  }
}

function updateVersionProgress(rowIndex, versionType, status, notes, leader, editor, normalTargetDate, tdmTargetDate) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName("Projects") || ss.getSheetByName("專案");
    var title = sheet.getRange(rowIndex, 1).getValue();
    if (leader) sheet.getRange(rowIndex, 3).setValue(leader);
    if (editor) sheet.getRange(rowIndex, 4).setValue(editor);
    var verText = (versionType === 'tdm') ? "TDM 版本" : "普通版本";
    if (versionType === 'normal') {
      sheet.getRange(rowIndex, 6).setValue(status);
      if (normalTargetDate) sheet.getRange(rowIndex, 7).setValue(normalTargetDate);
    } else if (versionType === 'tdm') {
      sheet.getRange(rowIndex, 9).setValue(status);
      if (tdmTargetDate) sheet.getRange(rowIndex, 10).setValue(tdmTargetDate);
    }
    if (notes !== undefined && notes !== null) {
      sheet.getRange(rowIndex, 11).setValue(notes);
    }
    var details = "狀態 :" + status + " | 負責人 :" + leader + " | 剪接師 :" + editor;
    if (normalTargetDate) details += " | 普通版日期 :" + normalTargetDate;
    if (tdmTargetDate) details += " | TDM日期 :" + tdmTargetDate;
    logRevisionHistory("更新資料與進度", title, verText, details, notes);
    return { success: true, message: "資料已更新" };
  } catch (e) {
    return { success: false, message: "更新失敗 :" + e.toString() };
  }
}

function formatDate(dateVal) {
  if (!dateVal) return "";
  if (dateVal instanceof Date) {
    return Utilities.formatDate(dateVal, sheetTimeZone_(), "yyyy-MM-dd");
  }
  var str = String(dateVal).trim();
  if (str.match(/^\d{4}-\d{2}-\d{2}/)) {
    return str.substring(0, 10);
  }
  return str;
}

function getAdminSheet_() {
  var ss = getSpreadsheet();
  return ss.getSheetByName("Admin") || ss.getSheetByName("管理員") || ss.getSheetByName("Admins");
}

function getCurrentIdentity_() {
  var userEmail = Session.getActiveUser().getEmail() || "";
  var userName = "";
  var memberSheet = getSpreadsheet().getSheetByName("Member") || getSpreadsheet().getSheetByName("成員");
  if (memberSheet && userEmail) {
    var memberData = memberSheet.getDataRange().getValues();
    for (var i = 1; i < memberData.length; i++) {
      var account = String(memberData[i][1] || "").trim().toLowerCase();
      if (account && account === String(userEmail).trim().toLowerCase()) {
        userName = String(memberData[i][0] || "").trim();
        break;
      }
    }
  }
  return { email: String(userEmail || "").trim(), name: userName };
}

function findAdminRow_() {
  var identity = getCurrentIdentity_();
  var sheet = getAdminSheet_();
  if (!sheet || sheet.getLastRow() < 2) return null;
  var width = Math.max(sheet.getLastColumn(), 3);
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, width).getValues();
  var email = identity.email.toLowerCase();
  var name = identity.name;
  for (var i = 0; i < rows.length; i++) {
    var adminName = String(rows[i][0] || "").trim();
    var adminEmail = String(rows[i][1] || "").trim().toLowerCase();
    if ((email && adminEmail && adminEmail === email) || (name && adminName && adminName === name)) {
      return { name: adminName, email: adminEmail, password: String(rows[i][2] || "").trim() };
    }
  }
  return null;
}

function isCurrentUserAdmin_() {
  return !!findAdminRow_();
}

function requireAdminPassword_(password) {
  var row = findAdminRow_();
  if (!row) return "沒有後台權限";
  if (!row.password) return "Admin 分頁  C 欄未設定密碼";
  if (String(password || "").trim() !== row.password) return "密碼不正確";
  return "";
}

function verifyAdminPassword(password) {
  var message = requireAdminPassword_(password);
  return { success: !message, message: message || "密碼正確" };
}

function checkAdminAccess() {
  var identity = getCurrentIdentity_();
  return {
    success: true,
    isAdmin: isCurrentUserAdmin_(),
    email: identity.email,
    name: identity.name,
    sheetFound: !!getAdminSheet_()
  };
}

function adminUpdateProject(formObject) {
  try {
    if (!isCurrentUserAdmin_()) return { success: false, message: "沒有後台權限" };
    var sheet = getSpreadsheet().getSheetByName("Projects") || getSpreadsheet().getSheetByName("專案");
    if (!sheet) return { success: false, message: "找不到  Projects 工作表" };
    var rowIndex = Number(formObject.rowIndex);
    if (!rowIndex || rowIndex < 2) return { success: false, message: "列號不正確" };
    var oldTitle = sheet.getRange(rowIndex, 1).getValue();
    var needTdm = formObject.needTdm === "是" || formObject.needTdm === true ? "是" : "否";
    var tdmStatus = formObject.tdmStatus || (needTdm === "是" ? "待剪接" : "不需要");
    if (needTdm !== "是") tdmStatus = "不需要";
    sheet.getRange(rowIndex, 1, 1, 11).setValues([[
      formObject.title || "",
      formObject.category || "",
      formObject.leader || "",
      formObject.editor || "",
      formObject.driveUrl || "",
      formObject.normalStatus || "待剪接",
      formObject.normalTargetDate || "",
      needTdm,
      tdmStatus,
      needTdm === "是" ? (formObject.tdmTargetDate || "") : "",
      formObject.notes || ""
    ]]);
    logRevisionHistory("後台修改", formObject.title || oldTitle, "全部資料", "由後台更新全部欄位 （ 原名稱 :" + oldTitle + "）", formObject.notes || "");
    return { success: true, message: "後台已更新" };
  } catch (e) {
    return { success: false, message: "後台更新失敗 :" + e.toString() };
  }
}

function adminDeleteProject(rowIndex, password) {
  try {
    if (!isCurrentUserAdmin_()) return { success: false, message: "沒有後台權限" };
    var sheet = getSpreadsheet().getSheetByName("Projects") || getSpreadsheet().getSheetByName("專案");
    if (!sheet) return { success: false, message: "找不到  Projects 工作表" };
    rowIndex = Number(rowIndex);
    if (!rowIndex || rowIndex < 2) return { success: false, message: "列號不正確" };
    var title = sheet.getRange(rowIndex, 1).getValue();
    sheet.deleteRow(rowIndex);
    logRevisionHistory("後台刪除", title, "全部資料", "由後台刪除專案", "");
    return { success: true, message: "已刪除" };
  } catch (e) {
    return { success: false, message: "刪除失敗 :" + e.toString() };
  }
}

function adminSheetConfig_(key) {
  var map = {
    revision: { names: ["Revision History", "修訂紀錄"], label: "Revision History", readonly: true },
    login: { names: ["Login Record", "登入紀錄"], label: "Login Record", readonly: true },
    manager: { names: ["Project Manager", "專案負責人"], label: "Project Manager", readonly: false },
    member: { names: ["Member", "成員"], label: "Member", readonly: false },
    category: { names: ["Categories", "類別"], label: "Categories", readonly: false }
  };
  return map[key] || null;
}

function adminFindSheet_(key) {
  var config = adminSheetConfig_(key);
  if (!config) return null;
  var ss = getSpreadsheet();
  for (var i = 0; i < config.names.length; i++) {
    var sheet = ss.getSheetByName(config.names[i]);
    if (sheet) return sheet;
  }
  return null;
}

function adminCellToText_(value) {
  if (value instanceof Date) return Utilities.formatDate(value, sheetTimeZone_(), "yyyy-MM-dd HH:mm:ss");
  return value === null || value === undefined ? "" : String(value);
}

function adminListSheet(key, password) {
  try {
    if (!isCurrentUserAdmin_()) return { success: false, message: "沒有後台權限" };
    var sheet = adminFindSheet_(key);
    if (!sheet) return { success: false, message: "找不到工作表" };
    var lastRow = sheet.getLastRow();
    var lastCol = Math.max(sheet.getLastColumn(), 1);
    var headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
    var rows = [];
    if (lastRow >= 2) {
      var startRow = 2;
      var readCount = lastRow - 1;
      if (key === "login") {
        startRow = 2;
        readCount = Math.min(lastRow, 401) - 1;
      } else if (key === "revision") {
        startRow = Math.max(2, lastRow - 399);
        readCount = lastRow - startRow + 1;
      } else {
        readCount = Math.min(lastRow, 401) - 1;
      }
      var values = sheet.getRange(startRow, 1, readCount, lastCol).getDisplayValues();
      for (var i = 0; i < values.length; i++) {
        var texts = values[i].map(function(cell) { return cell == null ? "" : String(cell); });
        if (texts.join("").trim() === "") continue;
        rows.push({ rowIndex: startRow + i, values: texts });
      }
    }
    var config = adminSheetConfig_(key);
    return { success: true, label: config.label, headers: headers, rows: rows, truncated: lastRow > 401, readonly: !!config.readonly };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function adminSaveSheetRow(key, rowIndex, values, password) {
  try {
    if (!isCurrentUserAdmin_()) return { success: false, message: "沒有後台權限" };
    var config = adminSheetConfig_(key);
    if (config && config.readonly) {
      return { success: false, message: "此工作表為唯讀 ， 不能新增或修改" };
    }
    var sheet = adminFindSheet_(key);
    if (!sheet) return { success: false, message: "找不到工作表" };
    values = values || [];
    var width = Math.max(sheet.getLastColumn(), values.length, 1);
    var row = [];
    for (var i = 0; i < width; i++) row.push(values[i] === undefined || values[i] === null ? "" : values[i]);
    rowIndex = Number(rowIndex);
    if (!rowIndex || rowIndex < 2) {
      sheet.appendRow(row);
    } else {
      sheet.getRange(rowIndex, 1, 1, width).setValues([row]);
    }
    logRevisionHistory("後台改表", adminSheetConfig_(key).label, key, "更新第" + (rowIndex || "新") + " 列", "");
    return { success: true, message: "已儲存" };
  } catch (e) {
    return { success: false, message: "儲存失敗 :" + e.toString() };
  }
}

function adminDeleteSheetRow(key, rowIndex, password) {
  try {
    if (!isCurrentUserAdmin_()) return { success: false, message: "沒有後台權限" };
    var config = adminSheetConfig_(key);
    if (config && config.readonly) {
      return { success: false, message: "此工作表為唯讀 ， 不能刪除" };
    }
    var sheet = adminFindSheet_(key);
    if (!sheet) return { success: false, message: "找不到工作表" };
    rowIndex = Number(rowIndex);
    if (!rowIndex || rowIndex < 2) return { success: false, message: "列號不正確" };
    sheet.deleteRow(rowIndex);
    logRevisionHistory("後台改表", adminSheetConfig_(key).label, key, "刪除第" + rowIndex + " 列", "");
    return { success: true, message: "已刪除" };
  } catch (e) {
    return { success: false, message: "刪除失敗 :" + e.toString() };
  }
}

// ============================================================
// 拍攝活動 Activity（僅 Admin）
// 工作表：Activity
// 欄位：A項目 B名稱 C時段資料(JSON)
// 時段 JSON 結構：[{ date, leader, coordinator, leaveList:[{name,start,end}] }]
// ============================================================

function getActivitySheet_() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("Activity");
  if (!sheet) {
    sheet = ss.insertSheet("Activity");
    sheet.appendRow(["項目", "名稱", "時段資料"]);
  }
  return sheet;
}

function getActivities() {
  try {
    if (!isCurrentUserAdmin_()) return { success: false, message: "沒有權限" };
    var sheet = getActivitySheet_();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, data: [] };
    var values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    var result = [];
    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      if (!row[0] && !row[1]) continue;
      var slots = [];
      try {
        if (row[2]) slots = JSON.parse(row[2]);
      } catch (e) { slots = []; }
      result.push({
        rowIndex: i + 2,
        item: row[0] || "",
        name: row[1] || "",
        slots: slots
      });
    }
    return { success: true, data: result };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function saveActivity(formObject) {
  try {
    if (!isCurrentUserAdmin_()) return { success: false, message: "沒有權限" };
    var sheet = getActivitySheet_();
    var slotsJson = "[]";
    try {
      slotsJson = JSON.stringify(formObject.slots || []);
    } catch (e) { slotsJson = "[]"; }
    var rowData = [
      formObject.item || "",
      formObject.name || "",
      slotsJson
    ];
    var rowIndex = Number(formObject.rowIndex);
    if (rowIndex && rowIndex >= 2) {
      sheet.getRange(rowIndex, 1, 1, 3).setValues([rowData]);
      logRevisionHistory("拍攝活動修改", formObject.name || formObject.item, "Activity", "更新拍攝活動", "");
    } else {
      sheet.appendRow(rowData);
      logRevisionHistory("拍攝活動新增", formObject.name || formObject.item, "Activity", "新增拍攝活動", "");
    }
    SpreadsheetApp.flush();
    return { success: true, message: "已儲存" };
  } catch (e) {
    return { success: false, message: "儲存失敗：" + e.toString() };
  }
}

function deleteActivity(rowIndex) {
  try {
    if (!isCurrentUserAdmin_()) return { success: false, message: "沒有權限" };
    var sheet = getActivitySheet_();
    rowIndex = Number(rowIndex);
    if (!rowIndex || rowIndex < 2) return { success: false, message: "列號不正確" };
    var name = sheet.getRange(rowIndex, 2).getValue() || sheet.getRange(rowIndex, 1).getValue();
    sheet.deleteRow(rowIndex);
    logRevisionHistory("拍攝活動刪除", name, "Activity", "刪除拍攝活動", "");
    return { success: true, message: "已刪除" };
  } catch (e) {
    return { success: false, message: "刪除失敗：" + e.toString() };
  }
}


function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action;
    var result;
    if (action === "getInitialData") result = getInitialData();
    else if (action === "getProjects") result = getProjects();
    else if (action === "getActivities") result = getActivities();
    else if (action === "addProject") result = addProject(data.payload);
    else if (action === "saveActivity") result = saveActivity(data.payload);
    else if (action === "deleteActivity") result = deleteActivity(data.payload);
    else if (action === "updateVersionProgress") result = updateVersionProgress(data.rowIndex, data.versionType, data.status, data.notes, data.leader, data.editor, data.normalTargetDate, data.tdmTargetDate);
    else if (action === "updateEventDate") result = updateEventDate(data.rowIndex, data.versionType, data.newDate);
    else result = { success: false, message: "未知動作" };
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: err.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}
