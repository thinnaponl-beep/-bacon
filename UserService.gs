// UserService

function getAdmins() {
  const cache = CacheService.getScriptCache();
  const CACHE_KEY = 'admins_data';
  const cachedAdmins = cache.get(CACHE_KEY);

  if (cachedAdmins) {
    return JSON.parse(cachedAdmins);
  }

  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(ADMIN_SHEET_NAME);
    if (!sheet || sheet.getLastRow() <= 1) {
        const result = { status: 'success', data: [] };
        cache.put(CACHE_KEY, JSON.stringify(result), 300); 
        return result;
    }
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
    const adminData = values.map(row => ({ 
      name: (row[0] || "").trim(), 
      imageUrl: row[1], 
      lineUserId: row[2], 
      email: (row[3] || "").trim().toLowerCase()
    }));

    const result = { status: 'success', data: adminData };
    cache.put(CACHE_KEY, JSON.stringify(result), 3600);
    return result;
  } catch (e) { 
    return { status: 'error', message: e.message, data: [] }; 
  }
}

function getUserNameByEmail(email) {
    if (!email) return 'System';
    const admins = getAdmins().data;
    const admin = admins.find(a => a.email === email.trim().toLowerCase());
    return admin ? admin.name : email.split('@')[0];
}

function getEmailByUserName(name) {
  if (!name) return null;
  const admins = getAdmins().data;
  const admin = admins.find(a => a.name === name.trim());
  return admin ? admin.email : null;
}
