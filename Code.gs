// Code.gs [ฉบับสมบูรณ์: Final Version + Pagination Fix]

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setTitle('ระบบนิเทศ ติดตาม และประเมินผล PISA69');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// -------------------------------------------------------------------------
// Helper Functions
// -------------------------------------------------------------------------
function getSS() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheetData(sheetName) {
  var sheet = getSS().getSheetByName(sheetName);
  var data = sheet.getDataRange().getValues();
  var result = [];
  var headers = data[0];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = data[i][j];
    }
    result.push(obj);
  }
  return result;
}

function findRowIndex(sheet, taskId, schoolId) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === String(taskId).trim() && String(data[i][2]).trim() === String(schoolId).trim()) {
      return i + 1;
    }
  }
  return -1;
}

// -------------------------------------------------------------------------
// 1. Authentication
// -------------------------------------------------------------------------
function hashPassword(password) {
  var rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password);
  var txtHash = '';
  for (i = 0; i < rawHash.length; i++) {
    var hashVal = rawHash[i];
    if (hashVal < 0) { hashVal += 256; }
    if (hashVal.toString(16).length == 1) { txtHash += '0'; }
    txtHash += hashVal.toString(16);
  }
  return txtHash;
}

function loginUser(username, password) {
  var sheet = getSS().getSheetByName('Users');
  var data = sheet.getDataRange().getValues();
  var inputHash = hashPassword(password);
   
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (String(row[1]) === String(username) && String(row[2]) === inputHash && row[7] === 'ACTIVE') {
      return {
        status: 'SUCCESS',
        username: row[1],
        role: row[3],
        ref_id: row[4],
        full_name: row[5]
      };
    }
  }
  return { status: 'FAILED', message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };
}

// -------------------------------------------------------------------------
// 2. School Logic
// -------------------------------------------------------------------------
function getTasksForSchool(schoolId) {
  try {
    var tasks = getSheetData('Tasks');
    var submissions = getSheetData('Submissions');
    var evaluations = getSheetData('Evaluations');
    
    var openTasks = tasks.filter(function(t) { return t.status === 'OPEN'; });

    var taskList = openTasks.map(function(task) {
      var mySub = submissions.find(function(s) {
        return String(s.task_id).trim() === String(task.task_id).trim() && String(s.school_id).trim() === String(schoolId).trim();
      });

      var myEval = null;
      if (mySub) {
        var subEvals = evaluations.filter(function(e) { return String(e.submission_id) === String(mySub.submission_id); });
        if (subEvals.length > 0) {
          myEval = subEvals[subEvals.length - 1];
        }
      }

      return {
        task_id: task.task_id,
        task_name: task.task_name,
        task_desc: task.description,
        task_type: task.task_type,
        deadline: task.close_date ? new Date(task.close_date).toLocaleDateString("th-TH") : "-",
        is_submitted: (mySub ? true : false),
        file_url: (mySub ? mySub.file_url : ''),
        eval_result: (myEval ? myEval.evaluation_result : null),
        eval_level: (myEval ? myEval.evaluation_level : null),
        eval_feedback: (myEval ? myEval.feedback : null)
      };
    });
    return taskList;
  } catch (e) {
    return [];
  }
}

// -------------------------------------------------------------------------
// 3. Supervisor & Admin Grading Logic
// -------------------------------------------------------------------------
function getSupervisorSchools(supervisorId) {
  var schools = getSheetData('Schools');
  var subms = getSheetData('Submissions');
  var mySchools = schools.filter(function(s) { return String(s.supervisor_id) === String(supervisorId); });

  return mySchools.map(function(s) {
    var count = subms.filter(function(sub) { return String(sub.school_id).trim() === String(s.school_id).trim(); }).length;
    return {
      school_id: s.school_id,
      school_name: s.school_name,
      cluster: s.cluster,
      submitted_count: count
    };
  });
}

function getSchoolWorksForGrading(schoolId) {
  var tasks = getSheetData('Tasks');
  var submissions = getSheetData('Submissions');
  var evaluations = getSheetData('Evaluations');

  var openTasks = tasks.filter(function(t) { return t.status === 'OPEN'; });

  return openTasks.map(function(task) {
    var sub = submissions.find(function(s) {
      return String(s.task_id).trim() === String(task.task_id).trim() && String(s.school_id).trim() === String(schoolId).trim();
    });
    
    var eval = null;
    if (sub) {
      var evals = evaluations.filter(function(e) { return String(e.submission_id) === String(sub.submission_id); });
      if (evals.length > 0) {
        eval = evals[evals.length - 1];
      }
    }

    return {
      task_id: task.task_id,
      task_name: task.task_name,
      submission: sub ? {
        id: sub.submission_id,
        file_url: sub.file_url,
        type: "FILE", 
        date: new Date(sub.submit_date).toLocaleDateString("th-TH")
      } : null,
      evaluation: eval ? {
        result: eval.evaluation_result,
        feedback: eval.feedback
      } : null
    };
  });
}

function saveEvaluation(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    var sheet = getSS().getSheetByName('Evaluations');
    var timestamp = new Date();
    
    sheet.appendRow([
      'EVAL-' + Utilities.getUuid().slice(0,8),
      data.submission_id,
      data.supervisor_id, 
      data.level,
      data.feedback,
      timestamp,
      data.result
    ]);
    
    return { status: 'SUCCESS', message: 'บันทึกผลการประเมินแล้ว' };
  } catch (e) {
    return { status: 'FAILED', message: 'Error: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// -------------------------------------------------------------------------
// 4. Admin Logic
// -------------------------------------------------------------------------
function getAllTasksAdmin() {
  try {
    var sheet = getSS().getSheetByName('Tasks');
    var data = sheet.getDataRange().getDisplayValues();
    var headers = data[0];
    var result = [];

    var colMap = {};
    headers.forEach(function(h, i) { colMap[String(h).trim()] = i; });

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      result.push({
        task_id: row[colMap['task_id']] || '',
        task_name: row[colMap['task_name']] || '',
        task_type: row[colMap['task_type']] || 'FILE',
        status: row[colMap['status']] || 'CLOSED',
        folder_id: row[colMap['folder_id']] || '',
        description: row[colMap['description']] || '',
        fiscal_year: row[colMap['fiscal_year']] || '',
        round: row[colMap['round']] || '',
        open_date: row[colMap['open_date']] || '-',
        close_date: row[colMap['close_date']] || '-'
      });
    }

    return result.reverse(); 

  } catch(e) {
    console.log("Error in getAllTasksAdmin: " + e.toString());
    return []; 
  }
}

function getAdminStats() {
  try {
    var schools = getSheetData('Schools');
    var tasks = getSheetData('Tasks');
    var submissions = getSheetData('Submissions');

    var totalSchools = schools.length;
    var activeTasks = tasks.filter(function(t) { return t.status === 'OPEN'; });
    var totalActiveTasks = activeTasks.length;

    var schoolProgress = schools.map(function(s) {
      var submittedCount = 0;
      if (totalActiveTasks > 0) {
        submittedCount = activeTasks.filter(function(t) {
          return submissions.some(function(sub) {
            return String(sub.task_id).trim() === String(t.task_id).trim() && String(sub.school_id).trim() === String(s.school_id).trim();
          });
        }).length;
      }

      var percent = (totalActiveTasks > 0) ? (submittedCount / totalActiveTasks) * 100 : 0;

      return {
        school_id: s.school_id,
        school_name: s.school_name,
        cluster: s.cluster,
        submitted: submittedCount,
        total: totalActiveTasks,
        percent: Math.round(percent)
      };
    });

    schoolProgress.sort(function(a, b) { return a.percent - b.percent; });

    return {
      totalSchools: totalSchools,
      totalActiveTasks: totalActiveTasks,
      schoolProgress: schoolProgress
    };

  } catch (e) {
    console.log(e);
    return null;
  }
}

function createNewTask(form) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    var ss = getSS();
    var sheet = ss.getSheetByName('Tasks');
    var lastRow = sheet.getLastRow();
    
    var newId = 'T001';
    if (lastRow > 1) {
      var lastId = sheet.getRange(lastRow, 1).getValue();
      var num = parseInt(lastId.replace('T', '')) + 1;
      newId = 'T' + String(num).padStart(3, '0');
    }

    var folderId = "";

    if (form.taskType !== 'LINK') {
      var parentFolderId = "1IrnW0rRVTHsrN6BY2xfCKzVdDT4U9hNL"; 
      var parentFolder;
      try { parentFolder = DriveApp.getFolderById(parentFolderId); } 
      catch (e) { parentFolder = DriveApp.getRootFolder(); }

      var folderName = newId + "_" + form.taskName;
      var newFolder = parentFolder.createFolder(folderName);
      newFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      folderId = newFolder.getId();
    }

    sheet.appendRow([
      newId, form.taskName, form.description, form.taskType, form.round, form.fiscalYear,
      "'" + form.openDate, "'" + form.closeDate, folderId, 'OPEN'
    ]);

    return { status: 'SUCCESS', message: 'สร้างงานใหม่เรียบร้อย! (ID: ' + newId + ')' };
  } catch (e) {
    return { status: 'FAILED', message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function toggleTaskStatus(taskId, currentStatus) {
  var sheet = getSS().getSheetByName('Tasks');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(taskId).trim()) {
      var newStatus = (currentStatus === 'OPEN') ? 'CLOSED' : 'OPEN';
      sheet.getRange(i + 1, 10).setValue(newStatus);
      return { status: 'SUCCESS', message: 'เปลี่ยนสถานะเป็น ' + newStatus };
    }
  }
  return { status: 'FAILED', message: 'ไม่พบงาน' };
}

function updateTask(data) {
  try {
    var sheet = getSS().getSheetByName('Tasks');
    var rows = sheet.getDataRange().getValues();
    var rowIndex = -1;

    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === String(data.taskId).trim()) {
        rowIndex = i + 1; 
        break;
      }
    }

    if (rowIndex === -1) {
      return { status: 'FAILED', message: 'ไม่พบรหัสภาระงานที่ระบุ' };
    }

    sheet.getRange(rowIndex, 2).setValue(data.taskName);
    sheet.getRange(rowIndex, 3).setValue(data.description);
    sheet.getRange(rowIndex, 4).setValue(data.taskType);
    sheet.getRange(rowIndex, 5).setValue(data.round);
    sheet.getRange(rowIndex, 6).setValue(data.fiscalYear);
    sheet.getRange(rowIndex, 7).setValue("'" + data.openDate);
    sheet.getRange(rowIndex, 8).setValue("'" + data.closeDate);

    return { status: 'SUCCESS', message: 'บันทึกการแก้ไขเรียบร้อยแล้ว' };
  } catch (e) {
    return { status: 'FAILED', message: 'Error: ' + e.toString() };
  }
}

function handleFileUpload(fileData) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    var contentType = fileData.mimeType || "application/pdf";
    var blob = Utilities.newBlob(Utilities.base64Decode(fileData.base64), contentType, fileData.fileName);
    
    var tasks = getSheetData('Tasks');
    var targetTask = tasks.find(function(t) { return String(t.task_id).trim() === String(fileData.taskId).trim(); });
    var rawFolderId = (targetTask && targetTask.folder_id) ? String(targetTask.folder_id).trim() : "";
    var folder;
    
    try {
      if (rawFolderId && rawFolderId.length > 10) folder = DriveApp.getFolderById(rawFolderId);
      else folder = DriveApp.getRootFolder();
    } catch (err) { folder = DriveApp.getRootFolder(); }

    var newFileName = fileData.schoolId + "_" + fileData.taskId + "_" + fileData.fileName;
    var file = folder.createFile(blob).setName(newFileName);
    
    var sheet = getSS().getSheetByName('Submissions');
    var timestamp = new Date();
    var existingRow = findRowIndex(sheet, fileData.taskId, fileData.schoolId);
    var submitInfo = JSON.stringify({ type: 'FILE', originalName: fileData.fileName, mode: (existingRow > 0 ? 'RESUBMIT' : 'FIRST') });

    if (existingRow > 0) {
      sheet.getRange(existingRow, 4).setValue(file.getUrl());
      sheet.getRange(existingRow, 5).setValue(file.getId());
      sheet.getRange(existingRow, 6).setValue(submitInfo);
      sheet.getRange(existingRow, 7).setValue(timestamp);
    } else {
      sheet.appendRow(['SUB-' + Utilities.getUuid().slice(0,8), fileData.taskId, fileData.schoolId, file.getUrl(), file.getId(), submitInfo, timestamp, 'ON_TIME']);
    }
    return { status: 'SUCCESS', message: 'ส่งไฟล์เรียบร้อยแล้ว' };
  } catch (e) {
    return { status: 'FAILED', message: 'Upload Error: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function handleLinkSubmission(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    var sheet = getSS().getSheetByName('Submissions');
    var timestamp = new Date();
    var existingRow = findRowIndex(sheet, data.taskId, data.schoolId);
    var submitInfo = JSON.stringify({ type: 'LINK', mode: (existingRow > 0 ? 'RESUBMIT' : 'FIRST') });

    if (existingRow > 0) {
      sheet.getRange(existingRow, 4).setValue(data.url);
      sheet.getRange(existingRow, 5).setValue('LINK_SUBMISSION');
      sheet.getRange(existingRow, 6).setValue(submitInfo);
      sheet.getRange(existingRow, 7).setValue(timestamp);
    } else {
      sheet.appendRow(['SUB-' + Utilities.getUuid().slice(0,8), data.taskId, data.schoolId, data.url, 'LINK_SUBMISSION', submitInfo, timestamp, 'ON_TIME']);
    }
    return { status: 'SUCCESS', message: 'บันทึกลิงก์เรียบร้อยแล้ว' };
  } catch (e) {
    return { status: 'FAILED', message: 'Link Error: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function deleteTask(taskId) {
  try {
    var ss = getSS();
    var sheet = ss.getSheetByName('Tasks');
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(taskId).trim()) {
        var folderId = data[i][8]; 
        if (folderId && folderId.toString().length > 5) {
          try {
            DriveApp.getFolderById(folderId).setTrashed(true); 
          } catch (err) {
            console.log("ไม่สามารถลบ Folder ได้: " + err);
          }
        }
        sheet.deleteRow(i + 1); 
        return { status: 'SUCCESS', message: 'ลบภาระงานและโฟลเดอร์เรียบร้อยแล้ว' };
      }
    }
    return { status: 'FAILED', message: 'ไม่พบรหัสภาระงานที่ต้องการลบ' };
  } catch (e) {
    return { status: 'FAILED', message: 'Error: ' + e.toString() };
  }
}// Code.gs [ฉบับสมบูรณ์: Final Version]

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setTitle('ระบบนิเทศ ติดตาม และประเมินผล PISA69');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// -------------------------------------------------------------------------
// Helper Functions
// -------------------------------------------------------------------------
function getSS() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheetData(sheetName) {
  var sheet = getSS().getSheetByName(sheetName);
  var data = sheet.getDataRange().getValues();
  var result = [];
  var headers = data[0];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = data[i][j];
    }
    result.push(obj);
  }
  return result;
}

function findRowIndex(sheet, taskId, schoolId) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === String(taskId).trim() && String(data[i][2]).trim() === String(schoolId).trim()) {
      return i + 1;
    }
  }
  return -1;
}

// -------------------------------------------------------------------------
// 1. Authentication
// -------------------------------------------------------------------------
function hashPassword(password) {
  var rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password);
  var txtHash = '';
  for (i = 0; i < rawHash.length; i++) {
    var hashVal = rawHash[i];
    if (hashVal < 0) { hashVal += 256; }
    if (hashVal.toString(16).length == 1) { txtHash += '0'; }
    txtHash += hashVal.toString(16);
  }
  return txtHash;
}

function loginUser(username, password) {
  var sheet = getSS().getSheetByName('Users');
  var data = sheet.getDataRange().getValues();
  var inputHash = hashPassword(password);
   
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (String(row[1]) === String(username) && String(row[2]) === inputHash && row[7] === 'ACTIVE') {
      return {
        status: 'SUCCESS',
        username: row[1],
        role: row[3],
        ref_id: row[4],
        full_name: row[5]
      };
    }
  }
  return { status: 'FAILED', message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };
}

// -------------------------------------------------------------------------
// 2. School Logic
// -------------------------------------------------------------------------
function getTasksForSchool(schoolId) {
  try {
    var tasks = getSheetData('Tasks');
    var submissions = getSheetData('Submissions');
    var evaluations = getSheetData('Evaluations');
    
    var openTasks = tasks.filter(function(t) { return t.status === 'OPEN'; });

    var taskList = openTasks.map(function(task) {
      var mySub = submissions.find(function(s) {
        return String(s.task_id).trim() === String(task.task_id).trim() && String(s.school_id).trim() === String(schoolId).trim();
      });

      var myEval = null;
      if (mySub) {
        var subEvals = evaluations.filter(function(e) { return String(e.submission_id) === String(mySub.submission_id); });
        if (subEvals.length > 0) {
          myEval = subEvals[subEvals.length - 1];
        }
      }

      return {
        task_id: task.task_id,
        task_name: task.task_name,
        task_desc: task.description,
        task_type: task.task_type,
        deadline: task.close_date ? new Date(task.close_date).toLocaleDateString("th-TH") : "-",
        is_submitted: (mySub ? true : false),
        file_url: (mySub ? mySub.file_url : ''),
        eval_result: (myEval ? myEval.evaluation_result : null),
        eval_level: (myEval ? myEval.evaluation_level : null),
        eval_feedback: (myEval ? myEval.feedback : null)
      };
    });
    return taskList;
  } catch (e) {
    return [];
  }
}

// -------------------------------------------------------------------------
// 3. Supervisor & Admin Grading Logic
// -------------------------------------------------------------------------
function getSupervisorSchools(supervisorId) {
  var schools = getSheetData('Schools');
  var subms = getSheetData('Submissions');
  var mySchools = schools.filter(function(s) { return String(s.supervisor_id) === String(supervisorId); });

  return mySchools.map(function(s) {
    var count = subms.filter(function(sub) { return String(sub.school_id).trim() === String(s.school_id).trim(); }).length;
    return {
      school_id: s.school_id,
      school_name: s.school_name,
      cluster: s.cluster,
      submitted_count: count
    };
  });
}

function getSchoolWorksForGrading(schoolId) {
  var tasks = getSheetData('Tasks');
  var submissions = getSheetData('Submissions');
  var evaluations = getSheetData('Evaluations');

  var openTasks = tasks.filter(function(t) { return t.status === 'OPEN'; });

  return openTasks.map(function(task) {
    var sub = submissions.find(function(s) {
      return String(s.task_id).trim() === String(task.task_id).trim() && String(s.school_id).trim() === String(schoolId).trim();
    });
    
    var eval = null;
    if (sub) {
      var evals = evaluations.filter(function(e) { return String(e.submission_id) === String(sub.submission_id); });
      if (evals.length > 0) {
        eval = evals[evals.length - 1];
      }
    }

    return {
      task_id: task.task_id,
      task_name: task.task_name,
      submission: sub ? {
        id: sub.submission_id,
        file_url: sub.file_url,
        type: "FILE", 
        date: new Date(sub.submit_date).toLocaleDateString("th-TH")
      } : null,
      evaluation: eval ? {
        result: eval.evaluation_result,
        feedback: eval.feedback
      } : null
    };
  });
}

function saveEvaluation(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    var sheet = getSS().getSheetByName('Evaluations');
    var timestamp = new Date();
    
    sheet.appendRow([
      'EVAL-' + Utilities.getUuid().slice(0,8),
      data.submission_id,
      data.supervisor_id, 
      data.level,
      data.feedback,
      timestamp,
      data.result
    ]);
    
    return { status: 'SUCCESS', message: 'บันทึกผลการประเมินแล้ว' };
  } catch (e) {
    return { status: 'FAILED', message: 'Error: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// -------------------------------------------------------------------------
// 4. Admin Logic
// -------------------------------------------------------------------------
function getAllTasksAdmin() {
  try {
    var sheet = getSS().getSheetByName('Tasks');
    var data = sheet.getDataRange().getDisplayValues();
    var headers = data[0];
    var result = [];

    var colMap = {};
    headers.forEach(function(h, i) { colMap[String(h).trim()] = i; });

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      result.push({
        task_id: row[colMap['task_id']] || '',
        task_name: row[colMap['task_name']] || '',
        task_type: row[colMap['task_type']] || 'FILE',
        status: row[colMap['status']] || 'CLOSED',
        folder_id: row[colMap['folder_id']] || '',
        description: row[colMap['description']] || '',
        fiscal_year: row[colMap['fiscal_year']] || '',
        round: row[colMap['round']] || '',
        open_date: row[colMap['open_date']] || '-',
        close_date: row[colMap['close_date']] || '-'
      });
    }

    return result.reverse(); 

  } catch(e) {
    console.log("Error in getAllTasksAdmin: " + e.toString());
    return []; 
  }
}

function getAdminStats() {
  try {
    var schools = getSheetData('Schools');
    var tasks = getSheetData('Tasks');
    var submissions = getSheetData('Submissions');

    var totalSchools = schools.length;
    var activeTasks = tasks.filter(function(t) { return t.status === 'OPEN'; });
    var totalActiveTasks = activeTasks.length;

    var schoolProgress = schools.map(function(s) {
      var submittedCount = 0;
      if (totalActiveTasks > 0) {
        submittedCount = activeTasks.filter(function(t) {
          return submissions.some(function(sub) {
            return String(sub.task_id).trim() === String(t.task_id).trim() && String(sub.school_id).trim() === String(s.school_id).trim();
          });
        }).length;
      }

      var percent = (totalActiveTasks > 0) ? (submittedCount / totalActiveTasks) * 100 : 0;

      return {
        school_id: s.school_id,
        school_name: s.school_name,
        cluster: s.cluster,
        submitted: submittedCount,
        total: totalActiveTasks,
        percent: Math.round(percent)
      };
    });

    schoolProgress.sort(function(a, b) { return a.percent - b.percent; });

    return {
      totalSchools: totalSchools,
      totalActiveTasks: totalActiveTasks,
      schoolProgress: schoolProgress
    };

  } catch (e) {
    console.log(e);
    return null;
  }
}

function createNewTask(form) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    var ss = getSS();
    var sheet = ss.getSheetByName('Tasks');
    var lastRow = sheet.getLastRow();
    
    var newId = 'T001';
    if (lastRow > 1) {
      var lastId = sheet.getRange(lastRow, 1).getValue();
      var num = parseInt(lastId.replace('T', '')) + 1;
      newId = 'T' + String(num).padStart(3, '0');
    }

    var folderId = "";

    if (form.taskType !== 'LINK') {
      var parentFolderId = "1IrnW0rRVTHsrN6BY2xfCKzVdDT4U9hNL"; 
      var parentFolder;
      try { parentFolder = DriveApp.getFolderById(parentFolderId); } 
      catch (e) { parentFolder = DriveApp.getRootFolder(); }

      var folderName = newId + "_" + form.taskName;
      var newFolder = parentFolder.createFolder(folderName);
      newFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      folderId = newFolder.getId();
    }

    sheet.appendRow([
      newId, form.taskName, form.description, form.taskType, form.round, form.fiscalYear,
      "'" + form.openDate, "'" + form.closeDate, folderId, 'OPEN'
    ]);

    return { status: 'SUCCESS', message: 'สร้างงานใหม่เรียบร้อย! (ID: ' + newId + ')' };
  } catch (e) {
    return { status: 'FAILED', message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function toggleTaskStatus(taskId, currentStatus) {
  var sheet = getSS().getSheetByName('Tasks');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(taskId).trim()) {
      var newStatus = (currentStatus === 'OPEN') ? 'CLOSED' : 'OPEN';
      sheet.getRange(i + 1, 10).setValue(newStatus);
      return { status: 'SUCCESS', message: 'เปลี่ยนสถานะเป็น ' + newStatus };
    }
  }
  return { status: 'FAILED', message: 'ไม่พบงาน' };
}

function updateTask(data) {
  try {
    var sheet = getSS().getSheetByName('Tasks');
    var rows = sheet.getDataRange().getValues();
    var rowIndex = -1;

    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === String(data.taskId).trim()) {
        rowIndex = i + 1; 
        break;
      }
    }

    if (rowIndex === -1) {
      return { status: 'FAILED', message: 'ไม่พบรหัสภาระงานที่ระบุ' };
    }

    sheet.getRange(rowIndex, 2).setValue(data.taskName);
    sheet.getRange(rowIndex, 3).setValue(data.description);
    sheet.getRange(rowIndex, 4).setValue(data.taskType);
    sheet.getRange(rowIndex, 5).setValue(data.round);
    sheet.getRange(rowIndex, 6).setValue(data.fiscalYear);
    sheet.getRange(rowIndex, 7).setValue("'" + data.openDate);
    sheet.getRange(rowIndex, 8).setValue("'" + data.closeDate);

    return { status: 'SUCCESS', message: 'บันทึกการแก้ไขเรียบร้อยแล้ว' };
  } catch (e) {
    return { status: 'FAILED', message: 'Error: ' + e.toString() };
  }
}

function handleFileUpload(fileData) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    var contentType = fileData.mimeType || "application/pdf";
    var blob = Utilities.newBlob(Utilities.base64Decode(fileData.base64), contentType, fileData.fileName);
    
    var tasks = getSheetData('Tasks');
    var targetTask = tasks.find(function(t) { return String(t.task_id).trim() === String(fileData.taskId).trim(); });
    var rawFolderId = (targetTask && targetTask.folder_id) ? String(targetTask.folder_id).trim() : "";
    var folder;
    
    try {
      if (rawFolderId && rawFolderId.length > 10) folder = DriveApp.getFolderById(rawFolderId);
      else folder = DriveApp.getRootFolder();
    } catch (err) { folder = DriveApp.getRootFolder(); }

    var newFileName = fileData.schoolId + "_" + fileData.taskId + "_" + fileData.fileName;
    var file = folder.createFile(blob).setName(newFileName);
    
    var sheet = getSS().getSheetByName('Submissions');
    var timestamp = new Date();
    var existingRow = findRowIndex(sheet, fileData.taskId, fileData.schoolId);
    var submitInfo = JSON.stringify({ type: 'FILE', originalName: fileData.fileName, mode: (existingRow > 0 ? 'RESUBMIT' : 'FIRST') });

    if (existingRow > 0) {
      sheet.getRange(existingRow, 4).setValue(file.getUrl());
      sheet.getRange(existingRow, 5).setValue(file.getId());
      sheet.getRange(existingRow, 6).setValue(submitInfo);
      sheet.getRange(existingRow, 7).setValue(timestamp);
    } else {
      sheet.appendRow(['SUB-' + Utilities.getUuid().slice(0,8), fileData.taskId, fileData.schoolId, file.getUrl(), file.getId(), submitInfo, timestamp, 'ON_TIME']);
    }
    return { status: 'SUCCESS', message: 'ส่งไฟล์เรียบร้อยแล้ว' };
  } catch (e) {
    return { status: 'FAILED', message: 'Upload Error: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function handleLinkSubmission(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    var sheet = getSS().getSheetByName('Submissions');
    var timestamp = new Date();
    var existingRow = findRowIndex(sheet, data.taskId, data.schoolId);
    var submitInfo = JSON.stringify({ type: 'LINK', mode: (existingRow > 0 ? 'RESUBMIT' : 'FIRST') });

    if (existingRow > 0) {
      sheet.getRange(existingRow, 4).setValue(data.url);
      sheet.getRange(existingRow, 5).setValue('LINK_SUBMISSION');
      sheet.getRange(existingRow, 6).setValue(submitInfo);
      sheet.getRange(existingRow, 7).setValue(timestamp);
    } else {
      sheet.appendRow(['SUB-' + Utilities.getUuid().slice(0,8), data.taskId, data.schoolId, data.url, 'LINK_SUBMISSION', submitInfo, timestamp, 'ON_TIME']);
    }
    return { status: 'SUCCESS', message: 'บันทึกลิงก์เรียบร้อยแล้ว' };
  } catch (e) {
    return { status: 'FAILED', message: 'Link Error: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function deleteTask(taskId) {
  try {
    var ss = getSS();
    var sheet = ss.getSheetByName('Tasks');
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(taskId).trim()) {
        var folderId = data[i][8]; 
        if (folderId && folderId.toString().length > 5) {
          try {
            DriveApp.getFolderById(folderId).setTrashed(true); 
          } catch (err) {
            console.log("ไม่สามารถลบ Folder ได้: " + err);
          }
        }
        sheet.deleteRow(i + 1); 
        return { status: 'SUCCESS', message: 'ลบภาระงานและโฟลเดอร์เรียบร้อยแล้ว' };
      }
    }
    return { status: 'FAILED', message: 'ไม่พบรหัสภาระงานที่ต้องการลบ' };
  } catch (e) {
    return { status: 'FAILED', message: 'Error: ' + e.toString() };
  }
}