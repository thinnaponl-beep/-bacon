// ChecklistService - Personal Use Version (Notifications Removed)

function updateProjectDetails(projectId, newDetailsData) {
  if (newDetailsData.checklist && newDetailsData.checklist.length > 0) {
      const newItem = newDetailsData.checklist[newDetailsData.checklist.length - 1];
      
      if (newItem && !newItem.history) { 
          newItem.history = [];
          newItem.comments = [];
          
          const creatorName = getUserNameByEmail(Session.getActiveUser().getEmail());
          let historyText = `สร้างขึ้นโดย ${creatorName}`;
          if (newItem.assignee) {
              historyText += ` และมอบหมายให้ ${newItem.assignee}`;
          }
          newItem.history.push({
              text: historyText,
              user: Session.getActiveUser().getEmail(),
              timestamp: new Date().toISOString()
          });
      }
  }
  return updateCell(projectId, COL.DETAILS_DATA, JSON.stringify(newDetailsData));
}

function addCommentToChecklistItem(projectId, commentData) {
  try {
    const { checklistIndex, commentText, files } = commentData;
    if (typeof checklistIndex === 'undefined') throw new Error("Checklist index is required.");
    if (!commentText && (!files || files.length === 0)) throw new Error("Comment text or file is required.");

    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DATA_SHEET_NAME);
    const rowIndex = findProjectRowIndex_(sheet, projectId);
    if (rowIndex === -1) throw new Error("ไม่พบโปรเจกต์");

    const detailsDataCell = sheet.getRange(rowIndex, COL.DETAILS_DATA);
    const detailsData = JSON.parse(detailsDataCell.getValue() || '{}');
    const checklist = detailsData.checklist || [];
    const item = checklist[checklistIndex];

    if (!item) throw new Error("ไม่พบขั้นตอนที่ต้องการคอมเมนต์");

    item.comments = item.comments || [];

    const uploadedFiles = (files || []).map(fileInfo =>
      handleFileUpload(fileInfo.fileData, fileInfo.mimeType, fileInfo.fileName)
    );

    const commenterEmail = Session.getActiveUser().getEmail();

    item.comments.push({
      text: commentText,
      user: commenterEmail,
      timestamp: new Date().toISOString(),
      files: uploadedFiles
    });

    detailsDataCell.setValue(JSON.stringify(detailsData));

    return { status: 'success', data: item };
  } catch (e) {
    Logger.log(e);
    return { status: 'error', message: 'เกิดข้อผิดพลาดในการเพิ่มคอมเมนต์: ' + e.message };
  }
}

function updateChecklistItemStatus(projectId, updateData) {
  try {
    const { checklistIndex, newStatus, logText, files, link } = updateData;
    if (typeof checklistIndex === 'undefined') throw new Error("Checklist index is required.");

    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DATA_SHEET_NAME);
    const rowIndex = findProjectRowIndex_(sheet, projectId);
    if (rowIndex === -1) throw new Error("ไม่พบโปรเจกต์");

    const projectRow = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    const projectAssignee = (projectRow[COL.ASSIGNED_TO - 1] || "").trim();

    const detailsDataCell = sheet.getRange(rowIndex, COL.DETAILS_DATA);
    const detailsData = JSON.parse(detailsDataCell.getValue() || '{}');
    const checklist = detailsData.checklist || [];
    const item = checklist[checklistIndex];
    if (!item) throw new Error("ไม่พบขั้นตอนที่ต้องการอัปเดต");
    
    const currentUserEmail = Session.getActiveUser().getEmail();
    const isProjectForTeam = projectAssignee === 'Team' || projectAssignee === 'bot';
    const isItemForTeam = item.assignee === 'Team' || item.assignee === 'bot';
    const isCurrentUserAssigned = item.assigneeEmail === currentUserEmail;

    if (!isCurrentUserAssigned && !isProjectForTeam && !isItemForTeam) {
        return { status: 'error', message: 'คุณไม่มีสิทธิ์อัปเดตงานย่อยนี้' };
    }
    
    const oldStatus = item.status;
    item.status = newStatus;
    
    item.history = item.history || [];

    const updaterEmail = Session.getActiveUser().getEmail();
    const updaterName = getUserNameByEmail(updaterEmail);
    
    const uploadedFiles = (files || []).map(fileInfo => 
      handleFileUpload(fileInfo.fileData, fileInfo.mimeType, fileInfo.fileName)
    );

    const historyEntry = {
      text: `[${updaterName}] ${logText || `เปลี่ยนสถานะจาก '${oldStatus}' เป็น '${newStatus}'`}`,
      user: updaterEmail,
      timestamp: new Date().toISOString(),
      files: uploadedFiles,
      link: link || null
    };
    item.history.push(historyEntry);

    detailsDataCell.setValue(JSON.stringify(detailsData));

    return { status: 'success' };
  } catch (e) {
    Logger.log(e);
    return { status: 'error', message: 'เกิดข้อผิดพลาดในการอัปเดตขั้นตอน: ' + e.message };
  }
}

function assignChecklistItem(projectId, checklistIndex, assigneeName) {
  try {
    if (typeof checklistIndex === 'undefined' || !assigneeName) {
      throw new Error("ข้อมูลไม่เพียงพอสำหรับการมอบหมายงาน");
    }

    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DATA_SHEET_NAME);
    const rowIndex = findProjectRowIndex_(sheet, projectId);
    if (rowIndex === -1) throw new Error("ไม่พบโปรเจกต์");

    const detailsDataCell = sheet.getRange(rowIndex, COL.DETAILS_DATA);
    const detailsData = JSON.parse(detailsDataCell.getValue() || '{}');
    const item = detailsData.checklist[checklistIndex];
    if (!item) throw new Error("ไม่พบขั้นตอนที่ต้องการมอบหมาย");

    const admin = getAdmins().data.find(a => a.name === assigneeName);
    if (!admin) {
      throw new Error(`ไม่พบผู้ใช้ชื่อ '${assigneeName}' ในระบบ Admin`);
    }
    
    if (!admin.email) {
      throw new Error(`ไม่สามารถมอบหมายงานได้: ผู้ใช้ '${assigneeName}' ไม่มีอีเมลลงทะเบียนในชีต Admin`);
    }

    const assignerName = getUserNameByEmail(Session.getActiveUser().getEmail());
    
    item.assignee = admin.name;
    item.assigneeImage = admin.imageUrl;
    item.assigneeEmail = admin.email;
    item.status = 'New';
    item.history = item.history || [];
    item.history.push({
      text: `[${assignerName}] มอบหมายงานให้ ${admin.name}`,
      user: Session.getActiveUser().getEmail(),
      timestamp: new Date().toISOString()
    });

    detailsDataCell.setValue(JSON.stringify(detailsData));

    return { status: 'success', data: item };
  } catch (e) {
    Logger.log(e);
    return { status: 'error', message: e.message };
  }
}

function acceptChecklistTask(projectId, checklistIndex) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DATA_SHEET_NAME);
    const rowIndex = findProjectRowIndex_(sheet, projectId);
    if (rowIndex === -1) throw new Error("ไม่พบโปรเจกต์");
    
    const projectRow = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    const projectAssignee = (projectRow[COL.ASSIGNED_TO - 1] || "").trim();

    const detailsDataCell = sheet.getRange(rowIndex, COL.DETAILS_DATA);
    const detailsData = JSON.parse(detailsDataCell.getValue() || '{}');

    if (!detailsData.checklist || !detailsData.checklist[checklistIndex]) {
      throw new Error("ไม่พบขั้นตอนที่ต้องการรับงาน");
    }

    const item = detailsData.checklist[checklistIndex];
    const currentUserEmail = Session.getActiveUser().getEmail();
    
    const isProjectForTeam = projectAssignee === 'Team' || projectAssignee === 'bot';
    const isItemForTeam = item.assignee === 'Team' || item.assignee === 'bot';
    const isCurrentUserAssigned = item.assigneeEmail === currentUserEmail;

    if (!isCurrentUserAssigned && !isProjectForTeam && !isItemForTeam) {
      return { status: 'error', message: 'คุณไม่มีสิทธิ์รับงานนี้' };
    }

    if (item.status !== 'New') {
      return { status: 'error', message: 'งานนี้ถูกรับไปแล้วหรือกำลังดำเนินการอยู่' };
    }

    if ((isProjectForTeam || isItemForTeam) && item.assigneeEmail !== currentUserEmail) {
        const admins = getAdmins().data;
        const user = admins.find(a => a.email === currentUserEmail);
        if (user) {
            item.assignee = user.name;
            item.assigneeImage = user.imageUrl;
            item.assigneeEmail = user.email;
        }
    }

    const result = updateChecklistItemStatus(projectId, {
      checklistIndex: checklistIndex,
      newStatus: 'In Progress',
      logText: 'รับงานและเริ่มดำเนินการ'
    });

    if (result.status === 'success') {
      const updatedDetailsData = JSON.parse(SpreadsheetApp.openById(SHEET_ID).getSheetByName(DATA_SHEET_NAME).getRange(rowIndex, COL.DETAILS_DATA).getValue());
      return { status: 'success', data: updatedDetailsData.checklist[checklistIndex] };
    }
    return result;

  } catch (e) {
    Logger.log(e);
    return { status: 'error', message: 'เกิดข้อผิดพลาดในการรับงาน: ' + e.message };
  }
}

function editChecklistItem(projectId, checklistIndex, newText, newDueDate) {
  try {
    if (typeof checklistIndex === 'undefined') throw new Error("Checklist index is required.");

    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DATA_SHEET_NAME);
    const rowIndex = findProjectRowIndex_(sheet, projectId);
    if (rowIndex === -1) throw new Error("ไม่พบโปรเจกต์");
    
    const projectRow = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    const projectAssignee = (projectRow[COL.ASSIGNED_TO - 1] || "").trim();

    const detailsDataCell = sheet.getRange(rowIndex, COL.DETAILS_DATA);
    const detailsData = JSON.parse(detailsDataCell.getValue() || '{}');
    const checklist = detailsData.checklist || [];
    const item = checklist[checklistIndex];
    if (!item) throw new Error("ไม่พบขั้นตอนที่ต้องการแก้ไข");

    const editorEmail = Session.getActiveUser().getEmail();
    const editorName = getUserNameByEmail(editorEmail);
    const isProjectForTeam = projectAssignee === 'Team' || projectAssignee === 'bot';
    const isItemForTeam = item.assignee === 'Team' || item.assignee === 'bot';
    const isCurrentUserAssigned = item.assigneeEmail === editorEmail;

    if (!isCurrentUserAssigned && !isProjectForTeam && !isItemForTeam) {
        return { status: 'error', message: 'คุณไม่มีสิทธิ์แก้ไขงานย่อยนี้' };
    }

    const oldText = item.text;
    const oldDueDate = item.dueDate || "";
    const textChanged = oldText !== newText;
    const dueDateChanged = oldDueDate !== newDueDate;

    if (!textChanged && !dueDateChanged) {
      return { status: 'success', message: 'ไม่มีการเปลี่ยนแปลง' };
    }

    let logMessages = [];
    if (textChanged) {
      logMessages.push(`แก้ไขข้อความจาก "${oldText}" เป็น "${newText}"`);
      item.text = newText;
    }
    if (dueDateChanged) {
      logMessages.push(`เปลี่ยนกำหนดส่งเป็น ${newDueDate || 'ไม่มี'}`);
      item.dueDate = newDueDate || "";
    }

    item.history = item.history || [];
    item.history.push({
      text: `[${editorName}] ${logMessages.join(' และ ')}`,
      user: editorEmail,
      timestamp: new Date().toISOString()
    });

    detailsDataCell.setValue(JSON.stringify(detailsData));

    return { status: 'success', data: item };
  } catch (e) {
    Logger.log(e);
    return { status: 'error', message: 'เกิดข้อผิดพลาดในการแก้ไขงานย่อย: ' + e.message };
  }
}

function deleteChecklistItem(projectId, checklistIndex) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DATA_SHEET_NAME);
    const rowIndex = findProjectRowIndex_(sheet, projectId);
    if (rowIndex === -1) throw new Error("ไม่พบโปรเจกต์");

    const detailsDataCell = sheet.getRange(rowIndex, COL.DETAILS_DATA);
    const detailsData = JSON.parse(detailsDataCell.getValue() || '{}');

    if (!detailsData.checklist || !detailsData.checklist[checklistIndex]) {
      throw new Error("ไม่พบขั้นตอนที่ต้องการลบ");
    }

    const itemToDelete = detailsData.checklist.splice(checklistIndex, 1)[0];
    const logText = `[งานย่อยถูกลบ] ${itemToDelete.text}`;
    
    detailsData.logs = detailsData.logs || [];
    detailsData.logs.push({
      text: logText,
      display_text: logText,
      type: "task_deleted",
      timestamp: new Date().toISOString(),
      user: Session.getActiveUser().getEmail(),
      files: []
    });

    detailsDataCell.setValue(JSON.stringify(detailsData));
    
    // There is no `item` variable to return here after splice.
    return { status: 'success' };

  } catch (e) {
    Logger.log(e);
    return { status: 'error', message: 'เกิดข้อผิดพลาดในการลบขั้นตอน: ' + e.message };
  }
}
