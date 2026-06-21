// ============================================================
//  kinn Assignment Tracker — Google Apps Script Backend
//  Sheet: Tasks  |  columns: id,title,description,owner,
//         department,client,startDate,dueDate,progress,
//         status,priority,type,notes,importance,urgency
// ============================================================

const TASKS_SHEET  = 'Tasks';
const CONFIG_SHEET = 'Config';

/* ------ Web App entry point ------ */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('kinn Assignment Tracker')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ------ Date helper ------ */
function fmtDate(val) {
  if (!val) return '';
  const d = new Date(val);
  if (isNaN(d.getTime())) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/* ------ Map Sheet row → UI record ------ */
function rowToTask(headers, row) {
  const r = {};
  headers.forEach((h, i) => {
    const v = row[i];
    if (h === 'startdate') r.start = fmtDate(v);
    else if (h === 'duedate') r.due = fmtDate(v);
    else if (h === 'title') r.task = (v === '' ? '' : String(v));
    else if (h === 'client') r.project = (v === '' ? '' : String(v));
    else if (h === 'department') r.category = (v === '' ? '' : String(v));
    else if (h === 'description') r.description = (v === '' ? '' : String(v));
    else r[h] = (v === '' ? '' : String(v));
  });
  // Normalise status
  const st = String(r.status || '').toLowerCase().replace(/\s/g, '');
  if (st === 'inprogress') r.status = 'In progress';
  else if (st === 'notstarted' || st === '') r.status = 'Not started';
  else if (st === 'completed' || st === 'done') r.status = 'Completed';
  else if (st === 'review') r.status = 'Review';
  else if (st === 'revising') r.status = 'Revising';
  // Normalise priority → importance
  if (!r.importance) {
    const p = String(r.priority || '').toLowerCase();
    r.importance = (p === 'high') ? 'high' : 'low';
  }
  if (!r.urgency) {
    const tp = String(r.type || '').toLowerCase();
    r.urgency = (tp === 'urgent' || tp === 'high') ? 'high' : 'low';
  }
  return r;
}

/* ------ Map UI record → Sheet row ------ */
function taskToRow(headers, task) {
  return headers.map(h => {
    if (h === 'startdate') return task.start || '';
    if (h === 'duedate')   return task.due   || '';
    if (h === 'title')     return task.task  || '';
    if (h === 'client')    return task.project   || '';
    if (h === 'department') return task.category || '';
    if (h === 'description') return task.description || task.notes || '';
    return task[h] !== undefined ? String(task[h]) : '';
  });
}

/* ------ Load all data ------ */
function getSheetData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Tasks
  let tasks = [];
  const ts = ss.getSheetByName(TASKS_SHEET);
  if (ts && ts.getLastRow() > 1) {
    const data = ts.getDataRange().getValues();
    const headers = data[0].map(x => String(x).toLowerCase().trim());
    tasks = data.slice(1)
      .map(row => rowToTask(headers, row))
      .filter(t => t.id);
  }

  // Config
  const cfg = loadConfig(ss, tasks);

  return { tasks, config: cfg };
}

function loadConfig(ss, tasks) {
  let projects = [], owners = [], categories = [], statuses = null, kpiWeights = null;

  const cs = ss.getSheetByName(CONFIG_SHEET);
  if (cs && cs.getLastRow() > 1) {
    const d = cs.getDataRange().getValues();
    d.slice(1).forEach(r => {
      const type = String(r[0]).toLowerCase().trim();
      const val  = String(r[1]).trim();
      const json = String(r[2] || '').trim();
      if (type === 'project'  && val) projects.push(val);
      if (type === 'owner'    && val) owners.push(val);
      if (type === 'category' && val) categories.push(val);
      if (type === 'statuses' && json) try { statuses = JSON.parse(json); } catch(e) {}
      if (type === 'kpiweights' && json) try { kpiWeights = JSON.parse(json); } catch(e) {}
    });
  }

  // Fall back: auto-extract unique values from Tasks
  if (!projects.length)   projects   = [...new Set(tasks.map(t => t.project).filter(Boolean))];
  if (!owners.length)     owners     = [...new Set(tasks.map(t => t.owner).filter(Boolean))];
  if (!categories.length) categories = [...new Set(tasks.map(t => t.category).filter(Boolean))];

  return {
    projects,
    owners,
    categories,
    statuses: statuses || [
      { name: 'Not started', pct: 0,   color: '#95a5a6' },
      { name: 'In progress', pct: 40,  color: '#f39c12' },
      { name: 'Review',      pct: 60,  color: '#3498db' },
      { name: 'Revising',    pct: 80,  color: '#9b59b6' },
      { name: 'Completed',   pct: 100, color: '#27ae60' }
    ],
    kpiWeights: kpiWeights || { completion: 30, onTime: 20, progress: 50 }
  };
}

/* ------ Save (create or update) one task ------ */
function saveTask(task) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(TASKS_SHEET);

  const HEADERS = ['id','title','description','owner','department','client',
                   'startDate','dueDate','progress','status','priority','type',
                   'notes','importance','urgency'];

  if (!sheet) {
    sheet = ss.insertSheet(TASKS_SHEET);
    sheet.appendRow(HEADERS);
  }

  // Ensure all required headers exist
  const existingH = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  HEADERS.forEach(h => {
    if (!existingH.includes(h)) { existingH.push(h); sheet.getRange(1, existingH.length).setValue(h); }
  });

  const rawHeaders = existingH.map(h => h.toLowerCase().trim());

  if (!task.id) task.id = String(Date.now());

  const row = taskToRow(rawHeaders, task);

  // Find existing row by id
  if (sheet.getLastRow() > 1) {
    const idCol = rawHeaders.indexOf('id') + 1;
    const ids   = sheet.getRange(2, idCol, sheet.getLastRow() - 1, 1).getValues().flat().map(String);
    const idx   = ids.indexOf(String(task.id));
    if (idx >= 0) {
      sheet.getRange(idx + 2, 1, 1, row.length).setValues([row]);
      return { success: true, id: task.id };
    }
  }

  sheet.appendRow(row);
  return { success: true, id: task.id };
}

/* ------ Delete one task by id ------ */
function deleteTask(id) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TASKS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return { success: false };

  const rawH  = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).toLowerCase().trim());
  const idCol = rawH.indexOf('id') + 1;
  if (idCol < 1) return { success: false };

  const ids = sheet.getRange(2, idCol, sheet.getLastRow() - 1, 1).getValues().flat().map(String);
  const row = ids.indexOf(String(id));
  if (row >= 0) { sheet.deleteRow(row + 2); return { success: true }; }
  return { success: false };
}

/* ------ Save config (projects, owners, categories, statuses, weights) ------ */
function saveConfig(config) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG_SHEET);
  if (!sheet) sheet = ss.insertSheet(CONFIG_SHEET);

  sheet.clearContents();
  sheet.appendRow(['type', 'value', 'json']);
  config.projects.forEach(v   => sheet.appendRow(['project',    v, '']));
  config.owners.forEach(v     => sheet.appendRow(['owner',      v, '']));
  config.categories.forEach(v => sheet.appendRow(['category',   v, '']));
  sheet.appendRow(['statuses',   '', JSON.stringify(config.statuses)]);
  sheet.appendRow(['kpiweights', '', JSON.stringify(config.kpiWeights)]);
  return { success: true };
}
