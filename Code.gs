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
      deleteProject
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
