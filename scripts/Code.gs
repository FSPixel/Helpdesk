const KB = {
  APP_NAME: 'IT Support Knowledge Base',
  SHEET_ARTICLES: 'Articles',
  SHEET_CATEGORIES: 'Categories',
  SHEET_LOGS: 'Logs',
  SHEET_USERS: 'Users',
  SHEET_SETTINGS: 'Settings',
  FOLDER_ID: 'PASTE_DRIVE_FOLDER_ID_HERE'
};

function onOpen(e) {
  ensureSetup_();
  SpreadsheetApp.getUi()
    .createMenu('KB Admin')
    .addItem('Open Dashboard', 'openDashboard')
    .addItem('Seed Sample Data', 'seedSampleData')
    .addItem('Refresh Counters', 'refreshCounters')
    .addSeparator()
    .addItem('Manage Users', 'manageUsers')
    .addItem('View Audit Log', 'viewAuditLog')
    .addItem('Export Current Sheet PDF', 'exportActiveSheetPdf')
    .addToUi();

  logAction_('OPEN_SPREADSHEET', 'Spreadsheet opened');
}

function onInstall(e) {
  onOpen(e);
}

function doGet() {
  ensureSetup_();
  const t = HtmlService.createTemplateFromFile('UI');
  return t.evaluate()
    .setTitle(KB.APP_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function openDashboard() {
  requireRole_(['admin', 'editor', 'viewer', 'auditor']);
  const html = HtmlService.createHtmlOutputFromFile('UI')
    .setWidth(1450)
    .setHeight(950);
  SpreadsheetApp.getUi().showModalDialog(html, KB.APP_NAME);
  logAction_('OPEN_DASHBOARD', 'Dashboard opened');
}

function manageUsers() {
  requireRole_(['admin']);
  const html = HtmlService.createHtmlOutputFromFile('Users')
    .setWidth(900)
    .setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, 'Manage Users');
}

function viewAuditLog() {
  requireRole_(['admin', 'auditor']);
  const html = HtmlService.createHtmlOutputFromFile('Audit')
    .setWidth(1100)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Audit Log');
}

function exportActiveSheetPdf() {
  requireRole_(['admin', 'auditor']);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const fileName = `${KB.APP_NAME} - ${sheet.getName()} - ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmm')}.pdf`;

  const html = HtmlService.createHtmlOutput(
    `<html><body><p>Use article PDF export or all-articles PDF export from the dashboard.</p></body></html>`
  );
  const blob = html.getBlob().getAs(MimeType.PDF);
  const file = getOutputFolder_().createFile(blob).setName(fileName);

  logAction_('EXPORT_ACTIVE_SHEET_PDF', `Active sheet exported: ${file.getName()}`, {
    sheetName: sheet.getName(),
    fileId: file.getId(),
    fileUrl: file.getUrl()
  });

  SpreadsheetApp.getUi().alert(`PDF created:\n${file.getUrl()}`);
}

function ensureSetup_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Script harus dibound ke Google Sheets.');

  initSheet_(ss, KB.SHEET_ARTICLES, ['ID','CreatedAt','UpdatedAt','Title','Category','Tags','Content','AttachmentName','AttachmentUrl','CreatedBy','Status','CreatedByEmail','UpdatedByEmail','Visibility','Version']);
  initSheet_(ss, KB.SHEET_CATEGORIES, ['Category','Icon','Count']);
  initSheet_(ss, KB.SHEET_LOGS, ['Timestamp','Action','Detail','Actor','Email','Role','Meta']);
  initSheet_(ss, KB.SHEET_USERS, ['Email','Name','Role','Active']);
  initSheet_(ss, KB.SHEET_SETTINGS, ['Key','Value']);

  const users = ss.getSheetByName(KB.SHEET_USERS);
  if (users.getLastRow() === 1) {
    const email = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || 'admin@example.com';
    users.appendRow([email, 'Admin', 'admin', 'TRUE']);
  }

  const settings = ss.getSheetByName(KB.SHEET_SETTINGS);
  if (settings.getLastRow() === 1) {
    settings.appendRow(['APP_NAME', KB.APP_NAME]);
    settings.appendRow(['COMPANY', 'IT Support']);
  }

  const cat = ss.getSheetByName(KB.SHEET_CATEGORIES);
  if (cat.getLastRow() === 1) {
    cat.getRange(2,1,6,3).setValues([
      ['Hardware Issues','🖥️',0],
      ['Software & Apps','💻',0],
      ['Network & Internet','📶',0],
      ['Account & Security','🛡️',0],
      ['Onboarding','👤',0],
      ['Printer Support','🖨️',0]
    ]);
  }
}

function initSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    sh.setFrozenRows(1);
  }
}

function getAppData(query) {
  ensureSetup_();
  refreshCounters();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const art = ss.getSheetByName(KB.SHEET_ARTICLES);
  const cat = ss.getSheetByName(KB.SHEET_CATEGORIES);

  const categories = cat.getDataRange().getValues().slice(1).map(r => ({
    name: r[0], icon: r[1], count: Number(r[2]) || 0
  }));

  const user = getCurrentUser_();
  const rows = art.getDataRange().getValues().slice(1).map((r, idx) => ({
    row: idx + 2,
    id: r[0],
    createdAt: r[1],
    updatedAt: r[2],
    title: r[3],
    category: r[4],
    tags: r[5],
    content: r[6],
    attachmentName: r[7],
    attachmentUrl: r[8],
    createdBy: r[9],
    status: r[10],
    createdByEmail: r[11],
    updatedByEmail: r[12],
    visibility: r[13] || 'public',
    version: r[14] || 1
  })).filter(r => r.title && canViewArticle_(r, user.email, user.role));

  const q = (query || '').trim().toLowerCase();
  const filtered = q
    ? rows.filter(r => [r.title, r.category, r.tags, r.content].join(' ').toLowerCase().includes(q))
    : rows;

  const recent = filtered
    .sort((a, b) => new Date(b.updatedAt || b.createdAt) - new Date(a.updatedAt || a.createdAt))
    .slice(0, 12);

  return {
    appName: KB.APP_NAME,
    total: rows.length,
    categories,
    recent,
    role: user.role,
    email: user.email,
    name: user.name
  };
}

function getArticleById(id) {
  ensureSetup_();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KB.SHEET_ARTICLES);
  const values = sh.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === id) {
      const article = {
        row: i + 1,
        id: values[i][0],
        createdAt: values[i][1],
        updatedAt: values[i][2],
        title: values[i][3],
        category: values[i][4],
        tags: values[i][5],
        content: values[i][6],
        attachmentName: values[i][7],
        attachmentUrl: values[i][8],
        createdBy: values[i][9],
        status: values[i][10],
        createdByEmail: values[i][11],
        updatedByEmail: values[i][12],
        visibility: values[i][13] || 'public',
        version: values[i][14] || 1
      };

      const user = getCurrentUser_();
      if (!canViewArticle_(article, user.email, user.role)) throw new Error('Access denied');
      return article;
    }
  }
  return null;
}

function saveArticle(payload) {
  ensureSetup_();

  const role = getCurrentRole_();
  if (!['admin', 'editor'].includes(role)) throw new Error('No permission to save articles');

  validateArticle_(payload);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(KB.SHEET_ARTICLES);
  const now = new Date();
  const user = getCurrentUser_();

  let attachmentName = payload.existingAttachmentName || '';
  let attachmentUrl = payload.existingAttachmentUrl || '';

  if (payload.attachmentBase64 && payload.attachmentName) {
    const folder = getOutputFolder_();
    const bytes = Utilities.base64Decode(payload.attachmentBase64);
    const blob = Utilities.newBlob(bytes, payload.attachmentMime || MimeType.PLAIN_TEXT, payload.attachmentName);
    const file = folder.createFile(blob);
    attachmentName = file.getName();
    attachmentUrl = file.getUrl();
  }

  if (payload.id) {
    const article = getArticleById(payload.id);
    if (!article) throw new Error('Article not found');
    if (!canEditArticle_(article, user.email, role)) throw new Error('Access denied');

    const version = Number(article.version || 1) + 1;
    sh.getRange(article.row, 2, 1, 14).setValues([[
      article.createdAt,
      now,
      payload.title.trim(),
      payload.category.trim(),
      normalizeCsv_(payload.tags),
      payload.content.trim(),
      attachmentName,
      attachmentUrl,
      payload.createdBy.trim(),
      payload.status || 'Published',
      article.createdByEmail || user.email,
      user.email,
      payload.visibility || article.visibility || 'public',
      version
    ]]);

    logAction_('UPDATE_ARTICLE', payload.title, { articleId: payload.id, version: version });
    refreshCounters();
    return { ok: true, id: payload.id };
  }

  const id = Utilities.getUuid();
  sh.appendRow([
    id,
    now,
    now,
    payload.title.trim(),
    payload.category.trim(),
    normalizeCsv_(payload.tags),
    payload.content.trim(),
    attachmentName,
    attachmentUrl,
    payload.createdBy.trim(),
    'Published',
    user.email,
    user.email,
    payload.visibility || 'public',
    1
  ]);

  logAction_('CREATE_ARTICLE', payload.title, { articleId: id });
  refreshCounters();
  return { ok: true, id: id };
}

function deleteArticle(id) {
  ensureSetup_();

  const role = getCurrentRole_();
  if (role !== 'admin') throw new Error('Only admin can delete articles');

  const article = getArticleById(id);
  if (!article) throw new Error('Article not found');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheetByName(KB.SHEET_ARTICLES).deleteRow(article.row);

  logAction_('DELETE_ARTICLE', article.title, { articleId: id });
  refreshCounters();
  return { ok: true };
}

function refreshCounters() {
  ensureSetup_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const art = ss.getSheetByName(KB.SHEET_ARTICLES);
  const cat = ss.getSheetByName(KB.SHEET_CATEGORIES);

  const rows = art.getDataRange().getValues().slice(1);
  const counts = {};
  rows.forEach(r => {
    const category = r[4];
    if (category) counts[category] = (counts[category] || 0) + 1;
  });

  const values = cat.getDataRange().getValues();
  const updates = [];
  for (let i = 1; i < values.length; i++) {
    updates.push([counts[values[i][0]] || 0]);
  }

  if (updates.length) {
    cat.getRange(2, 3, updates.length, 1).setValues(updates);
  }

  return { ok: true };
}

function seedSampleData() {
  ensureSetup_();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KB.SHEET_ARTICLES);
  if (sh.getLastRow() > 1) return { ok: true, message: 'Already seeded' };

  const now = new Date();
  const user = getCurrentUser_();
  const rows = [
    [Utilities.getUuid(), now, now, 'How to Reset VPN Password', 'Account & Security', 'vpn,password,portal', 'Reset password via portal then relogin to the VPN client.', '', '', 'Jane Doe', 'Published', user.email, user.email, 'public', 1],
    [Utilities.getUuid(), now, now, 'How to Fix Wi-Fi Dropouts', 'Network & Internet', 'wifi,network,router', 'Check router placement, cables, DNS, and restart the access point.', '', '', 'Jane Doe', 'Published', user.email, user.email, 'public', 1],
    [Utilities.getUuid(), now, now, 'Printer Offline Troubleshooting', 'Printer Support', 'printer,offline,spooler', 'Restart spooler service and reinstall the driver if needed.', '', '', 'Jane Doe', 'Published', user.email, user.email, 'public', 1]
  ];

  sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  refreshCounters();
  logAction_('SEED_SAMPLE_DATA', 'Sample data inserted', { count: rows.length });
  return { ok: true };
}

function validateArticle_(payload) {
  if (!payload) throw new Error('Payload is required');
  if (!payload.title || !String(payload.title).trim()) throw new Error('Title is required');
  if (!payload.category || !String(payload.category).trim()) throw new Error('Category is required');
  if (!payload.content || !String(payload.content).trim()) throw new Error('Content is required');
  if (!payload.createdBy || !String(payload.createdBy).trim()) throw new Error('Created By is required');
}

function normalizeCsv_(value) {
  return String(value || '')
    .split(',')
    .map(s => s.trim())
    .filter(Boolean)
    .join(', ');
}

function getCurrentEmail_() {
  return Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || 'unknown';
}

function getCurrentUser_() {
  const email = getCurrentEmail_();
  const usersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KB.SHEET_USERS);
  const users = usersSheet.getDataRange().getValues().slice(1);
  const found = users.find(r => String(r[0]).toLowerCase() === String(email).toLowerCase() && String(r[3]).toUpperCase() === 'TRUE');
  return found ? { email: email, name: found[1] || email, role: found[2] || 'viewer' } : { email: email, name: email, role: 'viewer' };
}