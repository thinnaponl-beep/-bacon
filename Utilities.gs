// =================================================================
// UTILITY & HELPER FUNCTIONS
// =================================================================

/**
 * Gets the deployed Web App URL.
 * @returns {string} The URL of the web app.
 */
function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

/**
 * Finds the row number for a given project ID.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to search in.
 * @param {string} projectId The project ID to find.
 * @returns {number} The row index, or -1 if not found.
 * @private
 */
function findProjectRowIndex_(sheet, projectId) {
  const finder = sheet.getRange("I:I").createTextFinder(projectId); // Assumes Project ID is in Column I
  const foundCell = finder.findNext();
  return foundCell ? foundCell.getRow() : -1;
}

/**
 * Handles file uploads to a specific Google Drive folder.
 * @param {string} fileData Base64 encoded file data.
 * @param {string} mimeType The MIME type of the file.
 * @param {string} fileName The name of the file.
 * @returns {{url: string, name: string}} An object with the file URL and name.
 */
function handleFileUpload(fileData, mimeType, fileName) {
  if (!fileData) return { url: "", name: "" };
  try {
    const decodedFile = Utilities.base64Decode(fileData, Utilities.Charset.UTF_8);
    const blob = Utilities.newBlob(decodedFile, mimeType, fileName);
    
    // Find or create the target folder
    const folders = DriveApp.getFoldersByName(FOLDER_NAME); // FOLDER_NAME from Settings.gs
    const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(FOLDER_NAME);
    
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); // Make file viewable
    
    return { url: file.getUrl(), name: file.getName() };
  } catch (e) {
    Logger.log(`File upload failed for ${fileName}: ${e.toString()}`);
    return { url: "", name: "" };
  }
}

/**
 * A generic function to update a single cell for a project.
 * @param {string} projectId The project ID.
 * @param {number} column The column number to update.
 * @param {any} value The new value for the cell.
 * @returns {{status: string, message?: string}} The result of the operation.
 */
function updateCell(projectId, column, value) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DATA_SHEET_NAME); // from Settings.gs
    const rowIndex = findProjectRowIndex_(sheet, projectId);
    if (rowIndex === -1) throw new Error("ไม่พบโปรเจกต์");
    
    sheet.getRange(rowIndex, column).setValue(value);
    return { status: 'success' };
  } catch (e) {
    Logger.log(`Error in updateCell for project ${projectId}: ${e.message}`);
    return { status: 'error', message: 'เกิดข้อผิดพลาดในการอัปเดต: ' + e.message };
  }
}

/**
 * Retrieves the project name for a given project ID.
 * @param {string} projectId The project ID.
 * @returns {string} The project name or a default string.
 */
function getProjectNameById(projectId) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DATA_SHEET_NAME);
  const rowIndex = findProjectRowIndex_(sheet, projectId);
  return rowIndex !== -1 ? sheet.getRange(rowIndex, COL.PROJECT_NAME).getValue() : "Unknown Project";
}

/**
 * Retrieves the list of followers for a given project ID.
 * @param {string} projectId The project ID.
 * @returns {string[]} An array of follower names.
 */
function getProjectFollowersById(projectId) {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DATA_SHEET_NAME);
    const rowIndex = findProjectRowIndex_(sheet, projectId);
    if (rowIndex === -1) return [];
    
    const followersString = sheet.getRange(rowIndex, COL.FOLLOWERS).getValue();
    return followersString ? followersString.split(',').map(name => name.trim()).filter(Boolean) : [];
}
