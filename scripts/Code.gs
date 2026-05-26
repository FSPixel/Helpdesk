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
  const html = HtmlService.createHtmlOutputFromFile('UI').setWidth(1450).setHeight(950);
  SpreadsheetApp.getUi().showModalDialog(html, KB.APP_NAME);
  logAction_('OPEN_DASHBOARD', 'Dashboard opened');
}

function manageUsers() {
  requireRole_(['admin']);
  const html = HtmlService.createHtmlOutputFromFile('Users').setWidth(900).setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, 'Manage Users');
}

function viewAuditLog() {
  requireRole_(['admin', 'auditor']);
  const html = HtmlService.createHtmlOutputFromFile('Audit').setWidth(1100).setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Audit Log');
}

function exportActiveSheetPdf() {
  requireRole_(['admin', 'auditor']);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const fileName = `${KB.APP_NAME} - ${sheet.getName()} - ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmm')}.pdf`;

  const html = HtmlService.createHtmlOutput('<html><body><p>Use article PDF export or all-articles PDF export from the dashboard.</p></body></html>');
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

  seedDefaultCategories_();
  seedDefaultUsers_();
  seedDefaultSettings_();
}

function initSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    sh.setFrozenRows(1);
  }
}

function seedDefaultCategories_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KB.SHEET_CATEGORIES);
  if (!sh) throw new Error('Categories sheet not found');

  const values = sh.getDataRange().getValues();
  if (values.length > 1 && String(values[1][0] || '').trim() !== '') return;

  if (sh.getLastRow() > 1) {
    const hasAny = values.slice(1).some(r => String(r[0] || '').trim() !== '');
    if (hasAny) return;
  }

  sh.getRange(2, 1, 8, 3).setValues([
    ['Hardware Issues','🖥️',0],
    ['Software & Apps','💻',0],
    ['Network & Internet','📶',0],
    ['Account & Security','🛡️',0],
    ['Onboarding','👤',0],
    ['Printer Support','🖨️',0],
    ['General Support','📁',0],
    ['Other','🧩',0]
  ]);
}

function seedDefaultUsers_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KB.SHEET_USERS);
  if (sh.getLastRow() > 1) return;

  const email = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || 'admin@example.com';
  sh.appendRow([email, 'Admin', 'admin', 'TRUE']);
}

function seedDefaultSettings_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KB.SHEET_SETTINGS);
  if (sh.getLastRow() > 1) return;
  sh.appendRow(['APP_NAME', KB.APP_NAME]);
  sh.appendRow(['COMPANY', 'IT Support']);
}

function getCategories() {
  ensureSetup_();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KB.SHEET_CATEGORIES);
  const values = sh.getDataRange().getValues();

  if (values.length <= 1) {
    seedDefaultCategories_();
    return getCategories();
  }

  const cats = values.slice(1)
    .filter(r => String(r[0] || '').trim() !== '')
    .map(r => ({
      name: String(r[0]).trim(),
      icon: String(r[1] || '📁').trim(),
      count: Number(r[2]) || 0
    }));

  if (!cats.length) {
    seedDefaultCategories_();
    return getCategories();
  }

  return cats;
}

function addCategory(categoryName, icon) {
  requireRole_(['admin', 'editor']);
  ensureSetup_();

  const name = String(categoryName || '').trim();
  if (!name) throw new Error('Category name is required');

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KB.SHEET_CATEGORIES);
  const values = sh.getDataRange().getValues().slice(1);
  const exists = values.some(r => String(r[0]).toLowerCase() === name.toLowerCase());
  if (exists) return { ok: true, message: 'Category already exists' };

  sh.appendRow([name, icon || '📁', 0]);
  logAction_('ADD_CATEGORY', name, { icon: icon || '📁' });
  return { ok: true, message: 'Category added' };
}

function getAppData(query) {
  ensureSetup_();
  refreshCounters();

  let categories = getCategories();
  if (!categories.length) {
    seedDefaultCategories_();
    categories = getCategories();
  }

  const user = getCurrentUser_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const art = ss.getSheetByName(KB.SHEET_ARTICLES);

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

  const q = String(query || '').trim().toLowerCase();
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

    logAction_('UPDATE_ARTICLE', payload.title, { articleId: payload.id, version });
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
  return { ok: true, id };
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
  return String(value || '').split(',').map(s => s.trim()).filter(Boolean).join(', ');
}

function getCurrentEmail_() {
  return Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || 'unknown';
}

function getCurrentUser_() {
  const email = getCurrentEmail_();
  const usersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KB.SHEET_USERS);
  const users = usersSheet.getDataRange().getValues().slice(1);
  const found = users.find(r => String(r[0]).toLowerCase() === String(email).toLowerCase() && String(r[3]).toUpperCase() === 'TRUE');
  return found ? { email, name: found[1] || email, role: found[2] || 'viewer' } : { email, name: email, role: 'viewer' };
}

/* ====== HELPERS MISSING FROM EARLIER VERSION ====== */

function getCurrentRole_() {
  return getCurrentUser_().role || 'viewer';
}

function requireRole_(allowedRoles) {
  const role = getCurrentRole_();
  if (!allowedRoles.includes(role)) {
    throw new Error('Access denied');
  }
}

function normalizeVisibility_(visibility) {
  const v = String(visibility || 'public').toLowerCase();
  return v === 'private' ? 'private' : 'public';
}

function canViewArticle_(article, email, role) {
  if (!article) return false;
  const visibility = normalizeVisibility_(article.visibility);
  if (visibility === 'public') return true;
  if (role === 'admin' || role === 'auditor') return true;
  if (String(article.createdByEmail || '').toLowerCase() === String(email || '').toLowerCase()) return true;
  if (String(article.updatedByEmail || '').toLowerCase() === String(email || '').toLowerCase()) return true;
  return false;
}

function canEditArticle_(article, email, role) {
  if (!article) return false;
  if (role === 'admin') return true;
  if (role === 'editor' && String(article.createdByEmail || '').toLowerCase() === String(email || '').toLowerCase()) return true;
  return false;
}

function logAction_(action, detail, meta) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(KB.SHEET_LOGS);
  const user = getCurrentUser_();
  sh.appendRow([
    new Date(),
    action,
    detail || '',
    user.name || '',
    user.email || '',
    user.role || 'viewer',
    meta ? JSON.stringify(meta) : ''
  ]);
}

function getOutputFolder_() {
  if (KB.FOLDER_ID && KB.FOLDER_ID !== 'PASTE_DRIVE_FOLDER_ID_HERE') {
    return DriveApp.getFolderById(KB.FOLDER_ID);
  }
  return DriveApp.getRootFolder();
}

function getUsersData() {
  requireRole_(['admin']);
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KB.SHEET_USERS);
  const users = sh.getDataRange().getValues().slice(1).map(r => ({
    email: r[0],
    name: r[1],
    role: r[2],
    active: r[3]
  }));
  return { users };
}

function saveUsers(rows) {
  requireRole_(['admin']);
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KB.SHEET_USERS);
  const data = [['Email','Name','Role','Active']].concat(rows || []);
  sh.clearContents();
  sh.getRange(1, 1, data.length, data[0].length).setValues(data);
  sh.setFrozenRows(1);
  logAction_('SAVE_USERS', 'Users updated', { count: (rows || []).length });
  return { ok: true };
}

function getAuditLogs(limit) {
  requireRole_(['admin', 'auditor']);
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KB.SHEET_LOGS);
  const values = sh.getDataRange().getValues().slice(1);
  const n = Math.max(1, Number(limit) || 300);

  return values.slice(-n).reverse().map(r => ({
    ts: r[0],
    action: r[1],
    detail: r[2],
    actor: r[3],
    email: r[4],
    role: r[5],
    meta: r[6]
  }));
}

function exportArticlePdf(id) {
  requireRole_(['admin', 'editor', 'viewer', 'auditor']);
  const article = getArticleById(id);
  if (!article) throw new Error('Article not found');

  const html = HtmlService.createHtmlOutput(`
    <html><body style="font-family:Arial,sans-serif;padding:24px">
      <h1>${escapeHtml_(article.title)}</h1>
      <p><b>Category:</b> ${escapeHtml_(article.category || '')}</p>
      <p><b>Tags:</b> ${escapeHtml_(article.tags || '')}</p>
      <p>${escapeHtml_(article.content || '').replace(/\n/g, '<br>')}</p>
    </body></html>
  `);

  const file = getOutputFolder_().createFile(html.getBlob().getAs(MimeType.PDF)).setName(`${article.title}.pdf`);
  logAction_('EXPORT_ARTICLE_PDF', article.title, { articleId: id, fileId: file.getId(), fileUrl: file.getUrl() });
  return { ok: true, url: file.getUrl() };
}

function exportAllArticlesPdf(query) {
  requireRole_(['admin', 'editor', 'viewer', 'auditor']);
  const data = getAppData(query);
  const html = HtmlService.createHtmlOutput(`
    <html><body style="font-family:Arial,sans-serif;padding:24px">
      <h1>${escapeHtml_(KB.APP_NAME)}</h1>
      ${(data.recent || []).map(a => `
        <div style="margin-bottom:18px;border-bottom:1px solid #ddd;padding-bottom:12px">
          <h2>${escapeHtml_(a.title)}</h2>
          <p><b>Category:</b> ${escapeHtml_(a.category || '')}</p>
          <p>${escapeHtml_(a.content || '').replace(/\n/g, '<br>')}</p>
        </div>
      `).join('')}
    </body></html>
  `);

  const file = getOutputFolder_().createFile(html.getBlob().getAs(MimeType.PDF)).setName(`${KB.APP_NAME} - All Articles.pdf`);
  logAction_('EXPORT_ALL_ARTICLES_PDF', 'All articles exported', { fileId: file.getId(), fileUrl: file.getUrl() });
  return { ok: true, url: file.getUrl() };
}

function escapeHtml_(s) {
  return String(s || '').replace(/[&<>"']/g, m => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;'
  }[m]));
}