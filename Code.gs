const APP_NAME = 'TaskFlow Database';
const USERS_SHEET = 'Users';
const TASKS_SHEET = 'Tasks';
const CACHE_DURATION = 300; // seconds

const USER_HEADERS = [
  'User ID', 'Name', 'Email', 'Password Hash', 'Salt',
  'Token', 'Avatar', 'Join Date', 'Last Login'
];
const TASK_HEADERS = [
  'Task ID', 'User Email', 'Title', 'Description',
  'Priority', 'Status', 'Due Date', 'Created Date', 'Updated Date'
];

/* ============================================================
 * WEB APP ENTRY POINT
 * ============================================================ */

function doGet(e) {
  setupDatabase();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('TaskFlow — Task Management')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * MANUAL ONE-CLICK SETUP
 * ----------------------------------------------------------------
 * The "TaskFlow Database" spreadsheet is created automatically the
 * first time anyone opens the deployed web app (doGet -> setupDatabase).
 *
 * If you want the sheet to be created right now — without deploying
 * or opening the web app first — just run this function directly:
 *   1. In the Apps Script editor, open Code.gs
 *   2. In the function dropdown (top toolbar) select "createDatabaseNow"
 *   3. Click "Run"
 *   4. Approve the permissions prompt the first time
 *   5. Check View > Logs (or Ctrl+Enter) to see the spreadsheet URL
 * ---------------------------------------------------------------- */
function createDatabaseNow() {
  const ss = setupDatabase();
  const url = ss.getUrl();
  Logger.log('✅ TaskFlow Database is ready: ' + url);
  return url;
}

/** Allows Index.html to pull in CSS.html / JS.html as partials. */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ============================================================
 * DATABASE SETUP  (auto-creates the Sheet on first run)
 * ============================================================ */

function setupDatabase() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty('SPREADSHEET_ID');
  let ss = null;

  if (ssId) {
    try {
      ss = SpreadsheetApp.openById(ssId);
    } catch (err) {
      ss = null; // stored ID no longer valid, fall through to recreate
    }
  }

  if (!ss) {
    ss = createSpreadsheet();
    props.setProperty('SPREADSHEET_ID', ss.getId());
  }

  createSheets(ss);
  return ss;
}

function createSpreadsheet() {
  return SpreadsheetApp.create(APP_NAME);
}

function createSheets(ss) {
  let usersSheet = ss.getSheetByName(USERS_SHEET);
  if (!usersSheet) {
    usersSheet = ss.insertSheet(USERS_SHEET);
    usersSheet.appendRow(USER_HEADERS);
    usersSheet.setFrozenRows(1);
    usersSheet.getRange(1, 1, 1, USER_HEADERS.length)
      .setFontWeight('bold').setBackground('#4361EE').setFontColor('#FFFFFF');
  }
  usersSheet.autoResizeColumns(1, USER_HEADERS.length);

  let tasksSheet = ss.getSheetByName(TASKS_SHEET);
  if (!tasksSheet) {
    tasksSheet = ss.insertSheet(TASKS_SHEET);
    tasksSheet.appendRow(TASK_HEADERS);
    tasksSheet.setFrozenRows(1);
    tasksSheet.getRange(1, 1, 1, TASK_HEADERS.length)
      .setFontWeight('bold').setBackground('#4361EE').setFontColor('#FFFFFF');
  }
  tasksSheet.autoResizeColumns(1, TASK_HEADERS.length);

  // Clean up the default "Sheet1" once our real sheets exist.
  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 2) {
    ss.deleteSheet(defaultSheet);
  }

  return { usersSheet, tasksSheet };
}

function getSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  const ssId = props.getProperty('SPREADSHEET_ID');
  if (!ssId) return setupDatabase();
  try {
    return SpreadsheetApp.openById(ssId);
  } catch (err) {
    return setupDatabase();
  }
}

/* ============================================================
 * AUTHENTICATION  (email + password, token-based sessions)
 * ============================================================ */

function signUp(name, email, password) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    name = (name || '').trim();
    email = (email || '').trim().toLowerCase();
    password = password || '';

    if (!name) throw new Error('Please enter your name.');
    if (!isValidEmail_(email)) throw new Error('Please enter a valid email address.');
    if (password.length < 6) throw new Error('Password must be at least 6 characters.');

    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(USERS_SHEET);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][2]).toLowerCase() === email) {
        throw new Error('An account with this email already exists. Please log in instead.');
      }
    }

    const now = new Date();
    const userId = 'USR-' + Utilities.getUuid().substring(0, 8).toUpperCase();
    const salt = Utilities.getUuid();
    const hash = hashPassword_(password, salt);
    const token = generateToken_();

    sheet.appendRow([userId, name, email, hash, salt, token, '', now, now]);
    sheet.autoResizeColumns(1, USER_HEADERS.length);

    return {
      success: true,
      token: token,
      user: {
        id: userId, name: name, email: email, avatar: '',
        joinDate: formatDate(now), lastLogin: formatDate(now)
      }
    };
  } catch (err) {
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

function logIn(email, password) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    email = (email || '').trim().toLowerCase();
    password = password || '';

    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(USERS_SHEET);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][2]).toLowerCase() === email) {
        const storedHash = data[i][3];
        const salt = data[i][4];
        const hash = hashPassword_(password, salt);
        if (hash !== storedHash) throw new Error('Incorrect password. Please try again.');

        const now = new Date();
        const token = generateToken_();
        const row = i + 1;
        sheet.getRange(row, 6).setValue(token);   // Token
        sheet.getRange(row, 9).setValue(now);      // Last Login
        sheet.autoResizeColumn(9);

        return {
          success: true,
          token: token,
          user: {
            id: data[i][0], name: data[i][1], email: data[i][2],
            avatar: data[i][6] || '',
            joinDate: formatDate(data[i][7]), lastLogin: formatDate(now)
          }
        };
      }
    }

    throw new Error('No account found with this email. Please sign up first.');
  } catch (err) {
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

function validateSession(token) {
  try {
    if (!token) throw new Error('No active session.');
    const found = findUserByToken_(token);
    const row = found.row;
    return {
      success: true,
      user: {
        id: row[0], name: row[1], email: row[2], avatar: row[6] || '',
        joinDate: formatDate(row[7]), lastLogin: formatDate(row[8])
      }
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function logOut(token) {
  try {
    const found = findUserByToken_(token);
    const ss = getSpreadsheet();
    ss.getSheetByName(USERS_SHEET).getRange(found.rowIndex, 6).setValue('');
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function updateProfileName(token, name) {
  try {
    name = (name || '').trim();
    if (!name) throw new Error('Name cannot be empty.');
    const found = findUserByToken_(token);
    const ss = getSpreadsheet();
    ss.getSheetByName(USERS_SHEET).getRange(found.rowIndex, 2).setValue(name);
    return { success: true, name: name };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function updateAvatar(token, dataUrl) {
  try {
    const found = findUserByToken_(token);
    if (dataUrl && dataUrl.length > 45000) {
      throw new Error('Image is too large. Please choose a smaller picture.');
    }
    const ss = getSpreadsheet();
    ss.getSheetByName(USERS_SHEET).getRange(found.rowIndex, 7).setValue(dataUrl || '');
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function getFullProfile(token) {
  try {
    const found = findUserByToken_(token);
    const email = found.row[2];
    const ss = getSpreadsheet();
    const data = ss.getSheetByName(TASKS_SHEET).getDataRange().getValues();

    let total = 0, completed = 0;
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === email) {
        total++;
        if (data[i][5] === 'Completed') completed++;
      }
    }

    return {
      success: true,
      profile: {
        id: found.row[0], name: found.row[1], email: found.row[2],
        avatar: found.row[6] || '',
        joinDate: formatDate(found.row[7]), lastLogin: formatDate(found.row[8]),
        totalTasks: total, completedTasks: completed
      }
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/* ---- helpers ---- */

function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function hashPassword_(password, salt) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password + '::' + salt);
  return bytes.map(b => ((b < 0 ? b + 256 : b).toString(16)).padStart(2, '0')).join('');
}

function generateToken_() {
  return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
}

/** Looks up a user row by session token. Throws if not found/expired. */
function findUserByToken_(token) {
  if (!token) throw new Error('Not signed in. Please log in again.');
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(USERS_SHEET);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][5] === token) {
      return { row: data[i], rowIndex: i + 1 };
    }
  }
  throw new Error('Your session has expired. Please log in again.');
}

/* ============================================================
 * TASKS — CRUD  (every query is filtered by the logged-in user's email)
 * ============================================================ */

function getTasks(token) {
  try {
    const email = requireEmail_(token);
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(TASKS_SHEET);
    const data = sheet.getDataRange().getValues();

    const tasks = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[1] === email) {
        var dueDateRawISO = '';
        if (row[6]) {
          try {
            var dd = (typeof row[6] === 'number') ? new Date((row[6] - 25569) * 86400 * 1000) : new Date(row[6]);
            if (!isNaN(dd.getTime())) dueDateRawISO = dd.toISOString();
          } catch(ex) {}
        }
        var createdRawISO = '';
        if (row[7]) {
          try {
            var cd = (typeof row[7] === 'number') ? new Date((row[7] - 25569) * 86400 * 1000) : new Date(row[7]);
            if (!isNaN(cd.getTime())) createdRawISO = cd.toISOString();
          } catch(ex) {}
        }
        tasks.push({
          id: row[0],
          userEmail: row[1],
          title: row[2],
          description: row[3],
          priority: row[4],
          status: row[5],
          dueDate: row[6] ? formatDateOnly(row[6]) : '',
          dueDateRaw: dueDateRawISO,
          createdDate: row[7] ? formatDate(row[7]) : '',
          createdDateRaw: createdRawISO,
          updatedDate: row[8] ? formatDate(row[8]) : ''
        });
      }
    }

    tasks.sort((a, b) => new Date(b.createdDateRaw) - new Date(a.createdDateRaw));
    return { success: true, tasks };
  } catch (err) {
    return { success: false, error: err.message, tasks: [] };
  }
}

function saveTask(token, taskData) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const email = requireEmail_(token);
    validateTaskInput_(taskData);

    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(TASKS_SHEET);
    const now = new Date();
    const taskId = 'TSK-' + Utilities.getUuid().substring(0, 8).toUpperCase();

    // Convert DD/MM/YYYY to a proper date string for storage
    var dueDateStr = '';
    if (taskData.dueDate) {
      dueDateStr = parseDueDateInput_(taskData.dueDate);
    }

    sheet.appendRow([
      taskId,
      email,
      taskData.title.trim(),
      (taskData.description || '').trim(),
      taskData.priority,
      taskData.status,
      dueDateStr,
      now,
      now
    ]);
    sheet.autoResizeColumns(1, TASK_HEADERS.length);

    invalidateDashboardCache_(email);
    return { success: true, taskId };
  } catch (err) {
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

function updateTask(token, taskData) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const email = requireEmail_(token);
    if (!taskData.id) throw new Error('Task ID is required.');
    validateTaskInput_(taskData);

    const rowIndex = findTaskRow_(taskData.id, email);
    if (rowIndex === -1) throw new Error('Task not found or you do not have access to it.');

    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(TASKS_SHEET);
    const now = new Date();

    var dueDateStr = '';
    if (taskData.dueDate) {
      dueDateStr = parseDueDateInput_(taskData.dueDate);
    }

    sheet.getRange(rowIndex, 3, 1, 5).setValues([[
      taskData.title.trim(),
      (taskData.description || '').trim(),
      taskData.priority,
      taskData.status,
      dueDateStr
    ]]);
    sheet.getRange(rowIndex, 9).setValue(now); // Updated Date
    sheet.autoResizeColumn(9);

    invalidateDashboardCache_(email);
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteTask(token, taskId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const email = requireEmail_(token);
    const rowIndex = findTaskRow_(taskId, email);
    if (rowIndex === -1) throw new Error('Task not found or you do not have access to it.');

    const ss = getSpreadsheet();
    ss.getSheetByName(TASKS_SHEET).deleteRow(rowIndex);

    invalidateDashboardCache_(email);
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

/** Quick status toggle used by the "Mark Complete" action. */
function markTaskComplete(token, taskId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const email = requireEmail_(token);
    const rowIndex = findTaskRow_(taskId, email);
    if (rowIndex === -1) throw new Error('Task not found or you do not have access to it.');

    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(TASKS_SHEET);
    sheet.getRange(rowIndex, 6).setValue('Completed');
    sheet.getRange(rowIndex, 9).setValue(new Date());
    sheet.autoResizeColumn(9);

    invalidateDashboardCache_(email);
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

/* ============================================================
 * DASHBOARD & REPORTS
 * ============================================================ */

function getDashboard(token) {
  try {
    const email = requireEmail_(token);
    const cache = CacheService.getScriptCache();
    const cacheKey = 'dashboard_' + email;
    const cached = cache.get(cacheKey);
    if (cached) return { success: true, stats: JSON.parse(cached) };

    const ss = getSpreadsheet();
    const data = ss.getSheetByName(TASKS_SHEET).getDataRange().getValues();

    let total = 0, pending = 0, completed = 0, overdue = 0, inProgress = 0;
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[1] !== email) continue;
      total++;

      const status = row[5];
      if (status === 'Completed') completed++;
      else if (status === 'In Progress') inProgress++;
      else pending++;

      if (row[6] && status !== 'Completed') {
        const due = new Date(row[6]);
        due.setHours(0, 0, 0, 0);
        if (due < today) overdue++;
      }
    }

    const completionPercentage = total > 0 ? Math.round((completed / total) * 100) : 0;
    const stats = { total, pending, completed, overdue, inProgress, completionPercentage };

    cache.put(cacheKey, JSON.stringify(stats), CACHE_DURATION);
    return { success: true, stats };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function getReports(token) {
  try {
    const email = requireEmail_(token);
    const ss = getSpreadsheet();
    const data = ss.getSheetByName(TASKS_SHEET).getDataRange().getValues();

    const now = new Date();
    const weekAgo = new Date(now); weekAgo.setDate(now.getDate() - 7);
    const monthAgo = new Date(now); monthAgo.setDate(now.getDate() - 30);

    let total = 0, completed = 0;
    let weekCreated = 0, weekCompleted = 0;
    let monthCreated = 0, monthCompleted = 0;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[1] !== email) continue;

      total++;
      const status = row[5];
      const created = row[7] ? new Date(row[7]) : null;
      const updated = row[8] ? new Date(row[8]) : null;

      if (status === 'Completed') completed++;
      if (created) {
        if (created >= weekAgo) weekCreated++;
        if (created >= monthAgo) monthCreated++;
      }
      if (status === 'Completed' && updated) {
        if (updated >= weekAgo) weekCompleted++;
        if (updated >= monthAgo) monthCompleted++;
      }
    }

    const completionPercentage = total > 0 ? Math.round((completed / total) * 100) : 0;

    return {
      success: true,
      reports: {
        completionPercentage,
        totalTasks: total,
        completedTasks: completed,
        weekly: { created: weekCreated, completed: weekCompleted },
        monthly: { created: monthCreated, completed: monthCompleted }
      }
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/* ============================================================
 * PRIVATE HELPERS
 * ============================================================ */

function requireEmail_(token) {
  const found = findUserByToken_(token);
  return found.row[2];
}

function findTaskRow_(taskId, email) {
  const ss = getSpreadsheet();
  const data = ss.getSheetByName(TASKS_SHEET).getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === taskId && data[i][1] === email) return i + 1;
  }
  return -1;
}

function validateTaskInput_(taskData) {
  if (!taskData || !taskData.title || taskData.title.trim() === '') {
    throw new Error('Task title is required.');
  }
  if (taskData.title.trim().length > 200) {
    throw new Error('Task title is too long (max 200 characters).');
  }
  if (!['Low', 'Medium', 'High'].includes(taskData.priority)) {
    throw new Error('Invalid priority value.');
  }
  if (!['Pending', 'In Progress', 'Completed'].includes(taskData.status)) {
    throw new Error('Invalid status value.');
  }
  if (taskData.dueDate) {
    // Accept DD/MM/YYYY or YYYY-MM-DD
    var parsed = parseDueDateInput_(taskData.dueDate);
    if (!parsed) {
      throw new Error('Invalid due date. Use DD/MM/YYYY format.');
    }
  }
}

/**
 * Parse a due date from DD/MM/YYYY or YYYY-MM-DD input into a formatted string.
 * Returns formatted date string or empty string on failure.
 */
function parseDueDateInput_(dateStr) {
  if (!dateStr) return '';
  dateStr = String(dateStr).trim();

  var d;
  // Try DD/MM/YYYY
  var parts = dateStr.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (parts) {
    d = new Date(parseInt(parts[3]), parseInt(parts[2]) - 1, parseInt(parts[1]));
  } else {
    // Try YYYY-MM-DD (ISO)
    d = new Date(dateStr);
  }

  if (!d || isNaN(d.getTime())) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone() || 'GMT', 'dd MMM yyyy');
}

function invalidateDashboardCache_(email) {
  CacheService.getScriptCache().remove('dashboard_' + email);
}

function formatDate(date) {
  if (!date) return '';
  try {
    var d;
    // Handle Google Sheets date serial numbers
    if (typeof date === 'number') {
      // Google Sheets epoch starts at Dec 30 1899
      d = new Date((date - 25569) * 86400 * 1000);
    } else if (date instanceof Date) {
      d = date;
    } else {
      d = new Date(date);
    }
    if (isNaN(d.getTime())) return String(date);
    return Utilities.formatDate(d, Session.getScriptTimeZone() || 'GMT', 'dd MMM yyyy, hh:mm a');
  } catch (e) {
    return String(date);
  }
}

function formatDateOnly(date) {
  if (!date) return '';
  try {
    var d;
    if (typeof date === 'number') {
      d = new Date((date - 25569) * 86400 * 1000);
    } else if (date instanceof Date) {
      d = date;
    } else {
      d = new Date(date);
    }
    if (isNaN(d.getTime())) return String(date);
    return Utilities.formatDate(d, Session.getScriptTimeZone() || 'GMT', 'dd MMM yyyy');
  } catch (e) {
    return String(date);
  }
}
