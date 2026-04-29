// Code.gs - Backend (Google Apps Script)

// Code.gs

const SHEET_NAME = "Sheet1"; // Change to your sheet tab name at bottom
const SPREADSHEET_ID = "1D0T8-qsGEajU9E9yNJvu6Z3VDJ454AmyhvWpjmVc5h0"; // Paste your Sheet ID here

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function doGet() {
  try {
    return HtmlService.createTemplateFromFile("Index")
      .evaluate()
      .setTitle("Kanban Board")
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
  } catch (e) {
    return ContentService.createTextOutput("ERROR: " + e.message);
  }
}

function getTasks() {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  var tasks = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) {
      tasks.push({
        id: data[i][0],
        title: data[i][1],
        description: data[i][2],
        status: data[i][3],
        lastUpdated: data[i][4]
      });
    }
  }
  return tasks;
}

function addTask(title, description) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    var id = "task_" + new Date().getTime();
    var timestamp = new Date().toISOString();
    sheet.appendRow([id, title, description, "Backlog", timestamp]);
    return { success: true, id: id };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function updateTaskStatus(taskId, newStatus) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === taskId) {
        sheet.getRange(i + 1, 4).setValue(newStatus);
        sheet.getRange(i + 1, 5).setValue(new Date().toISOString());
        return { success: true };
      }
    }
    return { success: false, error: "Task not found" };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function deleteTask(taskId) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === taskId) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: "Task not found" };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function editTask(taskId, newTitle, newDesc) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === taskId) {
        sheet.getRange(i + 1, 2).setValue(newTitle);
        sheet.getRange(i + 1, 3).setValue(newDesc);
        sheet.getRange(i + 1, 5).setValue(new Date().toISOString());
        return { success: true };
      }
    }
    return { success: false, error: "Task not found" };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function getLastUpdated() {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  var latest = "";
  for (var i = 1; i < data.length; i++) {
    if (data[i][4] && data[i][4] > latest) {
      latest = data[i][4];
    }
  }
  return latest;
}
