/**
 * @OnlyCurrentDoc
 * This is the main server file for the Google Apps Script project.
 */

// =================================================================
// MAIN SERVER FUNCTIONS
// =================================================================

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  template.urlParams = e.parameter;
  return template.evaluate()
    .setTitle("Internal System Development")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function handleClientRequest(functionName, args) {
  try {
    const availableFunctions = {
      getInitialData,
      getProjectById,
      saveData,
      updateProjectStatus,
      updateProjectDetails,
      addCommentToChecklistItem,
      updateChecklistItemStatus,
      assignChecklistItem,
      acceptChecklistTask,
      editChecklistItem,
      deleteChecklistItem,
      getAdmins,
      addLog,
      updateProjectCoreInfo,
      updateProjectInitialInfo,
      deleteProject,
      updateProjectCoreInfoEx,
      addAdminUser,
      updateProjectCoreInfoEx, 
      addAdminUser,
      updateProjectDates,  // <-- เพิ่มบรรทัดนี้
      updateSubtaskDates
    };

    if (typeof availableFunctions[functionName] === 'function') {
      return availableFunctions[functionName](...args);
    } else {
      throw new Error(`Function "${functionName}" is not defined or accessible.`);
    }
  } catch (e) {
    Logger.log(`Error in handleClientRequest calling ${functionName}: ${e.toString()}\nStack: ${e.stack}`);
    return { status: 'error', message: e.message };
  }
}

// =================================================================
// INITIAL DATA & CORE LOGIC
// =================================================================

function getInitialData() {
  const userEmail = Session.getActiveUser().getEmail();
  return {
    projects: getProjectData(),
    admins: getAdmins(),
    currentUser: {
      email: userEmail,
      name: getUserNameByEmail(userEmail)
    },
    unreadProjectIds: [] // Removed notification logic
  };
}

function saveData(formData) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DATA_SHEET_NAME);
  const projectId = `PROJ-${new Date().getTime()}`;
  const fileInfo = handleFileUpload(formData.fileData, formData.mimeType, formData.fileName);

  const creatorEmail = Session.getActiveUser().getEmail();
  const creatorName = getUserNameByEmail(creatorEmail);

  const initialDetailsData = {
    checklist: [],
    logs: [{
      text: "สร้างโปรเจกต์ใหม่",
      display_text: "สร้างโปรเจกต์ใหม่",
      type: "project_creation",
      files: fileInfo.url ? [fileInfo] : [],
      timestamp: new Date().toISOString(),
      user: creatorEmail
    }]
  };

  const projectLinks = formData.projectLink ? JSON.stringify([{ url: formData.projectLink, label: formData.projectLink }]) : JSON.stringify([]);

  sheet.appendRow([
    new Date(),                             // TIMESTAMP
    formData.projectName,                   // PROJECT_NAME
    "New",                                  // STATUS
    formData.details,                       // DETAILS
    fileInfo.name,                          // FILE_NAME
    fileInfo.url,                           // FILE_URL
    formData.assignedTo,                    // ASSIGNED_TO
    formData.assignedToImage,               // ASSIGNED_TO_IMAGE
    projectId,                              // PROJECT_ID
    JSON.stringify(initialDetailsData),     // DETAILS_DATA
    formData.projectDueDate || null,        // PROJECT_DUE_DATE
    creatorName,                            // FOLLOWERS
    projectLinks,                           // PROJECT_LINKS
    ""                                      // REVIEWER (Legacy)
  ]);

  return { status: 'success', message: 'บันทึกข้อมูลโครงการสำเร็จ!' };
}

function updateProjectCoreInfoEx(projectId, newName, newDueDate, newAssignee, newAssigneeImage) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DATA_SHEET_NAME);
    const finder = sheet.getRange("I:I").createTextFinder(projectId);
    const foundCell = finder.findNext();
    if (!foundCell) throw new Error("ไม่พบโปรเจกต์");
    
    const rowIndex = foundCell.getRow();
    
    sheet.getRange(rowIndex, COL.PROJECT_NAME).setValue(newName);
    sheet.getRange(rowIndex, COL.PROJECT_DUE_DATE).setValue(newDueDate || null);
    sheet.getRange(rowIndex, COL.ASSIGNED_TO).setValue(newAssignee || "");
    sheet.getRange(rowIndex, COL.ASSIGNED_TO_IMAGE).setValue(newAssigneeImage || "");

    return { status: 'success' };
  } catch (e) {
    return { status: 'error', message: 'เกิดข้อผิดพลาด: ' + e.message };
  }
}

function addAdminUser(name, email, imageUrl) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(ADMIN_SHEET_NAME);
    // แถวที่เพิ่มคือ Name, ImageUrl, LineId (เว้นไว้), Email
    sheet.appendRow([name, imageUrl || "", "", email]);
    
    // ล้าง Cache เพื่อให้ดึงข้อมูลใหม่ทันที
    CacheService.getScriptCache().remove('admins_data');
    
    return { status: 'success' };
  } catch (e) {
    return { status: 'error', message: 'เกิดข้อผิดพลาดในการเพิ่มผู้ใช้: ' + e.message };
  }
}

// --- สำหรับอัปเดตวันที่ของโปรเจกต์หลักในตาราง Gantt ---
function updateProjectDates(projectId, startDateStr, endDateStr) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DATA_SHEET_NAME);
    const finder = sheet.getRange("I:I").createTextFinder(projectId); // คอลัมน์ I คือ PROJECT_ID
    const foundCell = finder.findNext();
    if (!foundCell) throw new Error("ไม่พบโปรเจกต์");
    
    const rowIndex = foundCell.getRow();
    
    // อัปเดต วันที่เริ่ม (คอลัมน์ 1) และ กำหนดส่ง (คอลัมน์ 11)
    sheet.getRange(rowIndex, 1).setValue(new Date(startDateStr));
    sheet.getRange(rowIndex, 11).setValue(new Date(endDateStr));

    return { status: 'success' };
  } catch (e) {
    return { status: 'error', message: 'เกิดข้อผิดพลาด: ' + e.message };
  }
}

// --- สำหรับอัปเดตวันที่ของงานย่อย (Sub-tasks) ในตาราง Gantt ---
function updateSubtaskDates(projectId, subtaskIndex, startDateStr, endDateStr) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DATA_SHEET_NAME);
    const finder = sheet.getRange("I:I").createTextFinder(projectId);
    const foundCell = finder.findNext();
    if (!foundCell) throw new Error("ไม่พบโปรเจกต์");
    
    const rowIndex = foundCell.getRow();
    const detailsDataCell = sheet.getRange(rowIndex, 10); // คอลัมน์ 10 คือ DETAILS_DATA
    const detailsData = JSON.parse(detailsDataCell.getValue() || '{}');
    
    if (detailsData.checklist && detailsData.checklist[subtaskIndex]) {
       detailsData.checklist[subtaskIndex].startDate = startDateStr;
       detailsData.checklist[subtaskIndex].dueDate = endDateStr;
       detailsDataCell.setValue(JSON.stringify(detailsData));
    }
    
    return { status: 'success' };
  } catch (e) {
    return { status: 'error', message: 'เกิดข้อผิดพลาด: ' + e.message };
  }
}
