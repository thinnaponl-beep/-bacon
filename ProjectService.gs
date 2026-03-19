// ProjectService - Personal Use Version (Notifications Removed)

function getProjectData() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DATA_SHEET_NAME);
    if (!sheet || sheet.getLastRow() <= 1) return { status: 'success', data: [] };
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    return { status: 'success', data: values.map(parseProjectRow).sort((a, b) => new Date(b.rawTimestamp) - new Date(a.rawTimestamp)) };
  } catch (e) {
    Logger.log("Error in getProjectData: " + e.message);
    return { status: 'error', message: e.message, data: [] };
  }
}

function getProjectById(projectId) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DATA_SHEET_NAME);
  const rowIndex = findProjectRowIndex_(sheet, projectId);
  if (rowIndex === -1) return { status: 'error', message: 'ไม่พบโปรเจกต์' };

  const projectRow = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
  const projectData = parseProjectRow(projectRow);
  return { status: 'success', data: projectData };
}

function updateProjectStatus(projectId, newStatus) {
  const allowedStatuses = ["New", "In Progress", "On Hold", "Incomplete", "Completed"];
  if (!allowedStatuses.includes(newStatus)) {
    return { status: 'error', message: `ไม่สามารถเปลี่ยนเป็นสถานะ '${newStatus}' ได้` };
  }

  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DATA_SHEET_NAME);
    const rowIndex = findProjectRowIndex_(sheet, projectId);
    if (rowIndex === -1) throw new Error("ไม่พบโปรเจกต์");

    const projectRow = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    const assignee = (projectRow[COL.ASSIGNED_TO - 1] || "").trim();
    const currentUserEmail = Session.getActiveUser().getEmail();
    const currentUserName = getUserNameByEmail(currentUserEmail);

    const isAllowedToChangeStatus = assignee === currentUserName || assignee === 'Team' || assignee === 'bot' || assignee === '';

    if (!isAllowedToChangeStatus) {
      throw new Error("คุณไม่มีสิทธิ์อัปเดตสถานะของงานนี้");
    }

    const oldStatus = sheet.getRange(rowIndex, COL.STATUS).getValue();
    if (oldStatus === newStatus) return { status: 'success', message: 'สถานะไม่เปลี่ยนแปลง' };
    
    const statusCell = sheet.getRange(rowIndex, COL.STATUS);
    const detailsDataCell = sheet.getRange(rowIndex, COL.DETAILS_DATA);
    const detailsData = JSON.parse(detailsDataCell.getValue() || '{"logs":[]}');
    const logText = `[สถานะโครงการ] เปลี่ยนจาก '${oldStatus}' เป็น '${newStatus}'`;
    detailsData.logs.push({
      text: logText,
      display_text: logText,
      type: "status_change",
      timestamp: new Date().toISOString(),
      user: currentUserEmail,
      files: []
    });

    statusCell.setValue(newStatus);
    detailsDataCell.setValue(JSON.stringify(detailsData));

    return { status: 'success' };
  } catch (e) {
    Logger.log(e);
    return { status: 'error', message: 'เกิดข้อผิดพลาดในการอัปเดตสถานะ: ' + e.message };
  }
}

function updateProjectCoreInfo(projectId, newName, newDueDate) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DATA_SHEET_NAME);
    const rowIndex = findProjectRowIndex_(sheet, projectId);
    if (rowIndex === -1) throw new Error("ไม่พบโปรเจกต์");

    const editorEmail = Session.getActiveUser().getEmail();
    const editorName = getUserNameByEmail(editorEmail);
    
    const nameCell = sheet.getRange(rowIndex, COL.PROJECT_NAME);
    const dueDateCell = sheet.getRange(rowIndex, COL.PROJECT_DUE_DATE);
    const detailsDataCell = sheet.getRange(rowIndex, COL.DETAILS_DATA);
    
    const oldName = nameCell.getValue();
    const oldDueDateRaw = dueDateCell.getValue();
    const oldDueDate = oldDueDateRaw instanceof Date ? Utilities.formatDate(oldDueDateRaw, Session.getScriptTimeZone(), "yyyy-MM-dd") : "";

    const detailsData = JSON.parse(detailsDataCell.getValue() || '{"logs":[]}');
    const logMessages = [];

    if (oldName !== newName) {
      nameCell.setValue(newName);
      logMessages.push(`เปลี่ยนชื่องานจาก "${oldName}" เป็น "${newName}"`);
    }

    if (oldDueDate !== newDueDate) {
      dueDateCell.setValue(newDueDate || null);
      logMessages.push(`เปลี่ยนกำหนดส่งเป็น ${newDueDate || 'ไม่มี'}`);
    }

    if (logMessages.length > 0) {
      const logText = `[${editorName}] ${logMessages.join(' และ ')}`;
      detailsData.logs.push({
        text: logText,
        display_text: logText,
        type: "project_update",
        timestamp: new Date().toISOString(),
        user: editorEmail
      });
      detailsDataCell.setValue(JSON.stringify(detailsData));
    }

    return { status: 'success' };
  } catch (e) {
    Logger.log(e);
    return { status: 'error', message: 'เกิดข้อผิดพลาดในการแก้ไขข้อมูลงาน: ' + e.message };
  }
}

function updateProjectInitialInfo(projectId, newDetails, newLink) {
    try {
        const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DATA_SHEET_NAME);
        const rowIndex = findProjectRowIndex_(sheet, projectId);
        if (rowIndex === -1) throw new Error("ไม่พบโปรเจกต์");

        const editorEmail = Session.getActiveUser().getEmail();
        const editorName = getUserNameByEmail(editorEmail);

        const detailsCell = sheet.getRange(rowIndex, COL.DETAILS);
        const linksCell = sheet.getRange(rowIndex, COL.PROJECT_LINKS);
        const detailsDataCell = sheet.getRange(rowIndex, COL.DETAILS_DATA);

        const oldDetails = detailsCell.getValue();
        const oldLinksRaw = linksCell.getValue() || '[]';
        const oldLinks = JSON.parse(oldLinksRaw);
        const oldLinkUrl = oldLinks.length > 0 ? oldLinks[0].url : "";

        const detailsData = JSON.parse(detailsDataCell.getValue() || '{"logs":[]}');
        const logMessages = [];

        if (oldDetails !== newDetails) {
            detailsCell.setValue(newDetails);
            logMessages.push('แก้ไขรายละเอียดเริ่มต้น');
        }

        if (oldLinkUrl !== newLink) {
            const newLinksArray = newLink ? [{ url: newLink, label: newLink }] : [];
            linksCell.setValue(JSON.stringify(newLinksArray));
            logMessages.push('แก้ไขลิงก์เริ่มต้น');
        }

        if (logMessages.length > 0) {
            const logText = `[${editorName}] ${logMessages.join(' และ ')}`;
            detailsData.logs.push({
                text: logText,
                display_text: logText,
                type: "project_update",
                timestamp: new Date().toISOString(),
                user: editorEmail
            });
            detailsDataCell.setValue(JSON.stringify(detailsData));
        }

        return { status: 'success' };
    } catch (e) {
        Logger.log(e);
        return { status: 'error', message: 'เกิดข้อผิดพลาดในการอัปเดตข้อมูลเริ่มต้น: ' + e.message };
    }
}

function deleteProject(projectId) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DATA_SHEET_NAME);
    const rowIndex = findProjectRowIndex_(sheet, projectId);
    if (rowIndex === -1) throw new Error("ไม่พบโปรเจกต์ที่จะลบ");

    sheet.deleteRow(rowIndex);
    return { status: 'success' };
  } catch (e) {
    Logger.log(`Error deleting project ${projectId}: ${e.message}`);
    return { status: 'error', message: 'เกิดข้อผิดพลาดในการลบงาน: ' + e.message };
  }
}

function parseProjectRow(row) {
  let detailsData = { checklist: [], logs: [] };
  try {
    if (row[COL.DETAILS_DATA - 1]) {
      detailsData = JSON.parse(row[COL.DETAILS_DATA - 1]);
    }
  } catch (e) {
    Logger.log("Invalid JSON in DetailsData for project: " + row[COL.PROJECT_ID - 1]);
  }

  if (detailsData.checklist && Array.isArray(detailsData.checklist)) {
    detailsData.checklist = detailsData.checklist.map(item => ({
      ...item,
      history: item.history || [],
      comments: item.comments || []
    }));
  }

  detailsData.logs = (detailsData.logs || []).map(log => {
    if (!log.type) {
      log.type = log.text === "สร้างโปรเจกต์ใหม่" ? "project_creation" : "general";
    }
    if (!log.display_text) {
      log.display_text = log.text;
    }
    return log;
  });

  const timestampDate = new Date(row[COL.TIMESTAMP - 1]);
  const dueDateValue = row[COL.PROJECT_DUE_DATE - 1];
  const followersString = row[COL.FOLLOWERS - 1] || "";
  const projectLinksString = row[COL.PROJECT_LINKS - 1] || '[]';
  let projectLinks = [];
  try {
    projectLinks = JSON.parse(projectLinksString);
  } catch (e) {
    Logger.log("Invalid JSON in ProjectLinks for project: " + row[COL.PROJECT_ID - 1]);
  }

  const projectDueDate = (dueDateValue && dueDateValue instanceof Date)
    ? Utilities.formatDate(dueDateValue, Session.getScriptTimeZone(), "yyyy-MM-dd")
    : null;

  return {
    timestamp: Utilities.formatDate(timestampDate, "Asia/Bangkok", "d MMM yy"),
    rawTimestamp: timestampDate.toISOString(),
    projectName: (row[COL.PROJECT_NAME - 1] || "").trim(),
    status: (row[COL.STATUS - 1] || "").trim(),
    details: row[COL.DETAILS - 1],
    assignedTo: (row[COL.ASSIGNED_TO - 1] || "").trim(),
    assignedToImage: row[COL.ASSIGNED_TO_IMAGE - 1],
    projectId: row[COL.PROJECT_ID - 1],
    detailsData: detailsData,
    projectDueDate: projectDueDate,
    followers: followersString ? followersString.split(',').map(name => name.trim()) : [],
    projectLinks: projectLinks,
    reviewer: (row[COL.REVIEWER - 1] || "").trim() || null
  };
}

function addLog(projectId, logData) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DATA_SHEET_NAME);
    const rowIndex = findProjectRowIndex_(sheet, projectId);
    if (rowIndex === -1) throw new Error("ไม่พบโปรเจกต์");

    const detailsDataCell = sheet.getRange(rowIndex, COL.DETAILS_DATA);
    const detailsData = JSON.parse(detailsDataCell.getValue() || '{}');
    detailsData.logs = detailsData.logs || [];

    const uploadedFiles = (logData.files || []).map(fileInfo =>
      handleFileUpload(fileInfo.fileData, fileInfo.mimeType, fileInfo.fileName)
    );

    const updaterEmail = Session.getActiveUser().getEmail();

    const logEntry = {
      timestamp: new Date().toISOString(),
      user: updaterEmail,
      files: uploadedFiles,
      link: logData.link || null,
      type: logData.type || 'general',
      text: logData.commentText || logData.text,
      display_text: `[คอมเมนต์] ${logData.commentText || logData.text}`
    };

    detailsData.logs.push(logEntry);
    detailsDataCell.setValue(JSON.stringify(detailsData));
    
    return { status: 'success' };
  } catch (e) {
    Logger.log(`Error in addLog for project ${projectId}: ${e.message}`);
    return { status: 'error', message: 'เกิดข้อผิดพลาดในการเพิ่ม Log: ' + e.message };
  }
}
