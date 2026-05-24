function logAction_(action, detail, meta) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(KB.SHEET_LOGS);
  if (!sh) throw new Error('Logs sheet not found');

  const user = getCurrentUser_();
  sh.appendRow([
    new Date(),
    action || '',
    detail || '',
    user.name || '',
    user.email || '',
    user.role || 'viewer',
    meta ? JSON.stringify(meta) : ''
  ]);
}

function getAuditLogs(limit) {
  requireRole_(['admin', 'auditor']);

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KB.SHEET_LOGS);
  if (!sh) throw new Error('Logs sheet not found');

  const values = sh.getDataRange().getValues().slice(1).reverse();
  const maxRows = Number(limit || 300);

  const rows = values.slice(0, maxRows).map(r => ({
    ts: r[0],
    action: r[1],
    detail: r[2],
    actor: r[3],
    email: r[4],
    role: r[5],
    meta: r[6]
  }));

  logAction_('VIEW_AUDIT_LOG', `Loaded ${rows.length} rows`, { limit: maxRows });
  return rows;
}

function saveUsers(rows) {
  requireRole_(['admin']);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(KB.SHEET_USERS);
  if (!sh) throw new Error('Users sheet not found');

  if (!Array.isArray(rows)) throw new Error('Rows must be an array');

  const cleaned = rows
    .map(r => [
      String(r[0] || '').trim(),
      String(r[1] || '').trim(),
      String(r[2] || 'viewer').trim().toLowerCase(),
      String(r[3] || 'TRUE').trim().toUpperCase()
    ])
    .filter(r => r[0]);

  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 4).clearContent();
  }

  if (cleaned.length) {
    sh.getRange(2, 1, cleaned.length, 4).setValues(cleaned);
  }

  logAction_('UPDATE_USERS', `Saved ${cleaned.length} users`, { count: cleaned.length });
  return { ok: true, count: cleaned.length };
}

function exportArticlePdf(id) {
  ensureSetup_();

  const role = getCurrentRole_();
  if (!['admin', 'editor', 'auditor'].includes(role)) {
    throw new Error('No permission to export PDF');
  }

  const article = getArticleById(id);
  if (!article) throw new Error('Article not found');

  const html = buildArticlePdfHtml_(article);
  const blob = HtmlService.createHtmlOutput(html).getBlob().getAs(MimeType.PDF);

  const fileName = sanitizeFileName_(
    `${article.title} - ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmm')}.pdf`
  );

  const file = getOutputFolder_().createFile(blob).setName(fileName);

  logAction_('EXPORT_ARTICLE_PDF', article.title, {
    articleId: id,
    fileId: file.getId(),
    fileUrl: file.getUrl()
  });

  return { ok: true, url: file.getUrl(), name: file.getName() };
}

function exportAllArticlesPdf(filterCategory) {
  ensureSetup_();

  const role = getCurrentRole_();
  if (!['admin', 'auditor'].includes(role)) {
    throw new Error('No permission to export all articles');
  }

  const query = String(filterCategory || '').trim();
  const data = getAppData(query);
  const html = buildAllArticlesPdfHtml_(data.recent || [], query || 'All');

  const blob = HtmlService.createHtmlOutput(html).getBlob().getAs(MimeType.PDF);

  const fileName = sanitizeFileName_(
    `KB-All-${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmm')}.pdf`
  );

  const file = getOutputFolder_().createFile(blob).setName(fileName);

  logAction_('EXPORT_ALL_PDF', 'Exported article list as PDF', {
    filterCategory: query,
    fileId: file.getId(),
    fileUrl: file.getUrl()
  });

  return { ok: true, url: file.getUrl(), name: file.getName() };
}

function buildArticlePdfHtml_(a) {
  return `
  <html>
    <head>
      <meta charset="UTF-8">
      <style>
        body{font-family:Arial,sans-serif;padding:24px;color:#111827}
        h1{margin:0 0 8px;font-size:22px}
        .meta{color:#6b7280;font-size:12px;margin-bottom:16px}
        .box{border:1px solid #e5e7eb;border-radius:10px;padding:16px;margin-top:16px}
        pre{white-space:pre-wrap;font-family:inherit;line-height:1.5;margin:0}
        .label{font-weight:700;margin-bottom:6px}
        a{color:#2563eb;text-decoration:none}
      </style>
    </head>
    <body>
      <h1>${escapeHtml_(a.title || '')}</h1>
      <div class="meta">
        Category: ${escapeHtml_(a.category || '')} |
        Tags: ${escapeHtml_(a.tags || '')} |
        By: ${escapeHtml_(a.createdBy || '')} |
        Visibility: ${escapeHtml_(a.visibility || 'public')}
      </div>

      <div class="box">
        <div class="label">Instructions</div>
        <pre>${escapeHtml_(a.content || '')}</pre>
      </div>

      ${a.attachmentUrl ? `
        <div class="box">
          <div class="label">Attachment</div>
          <a href="${a.attachmentUrl}" target="_blank">${escapeHtml_(a.attachmentName || 'Open file')}</a>
        </div>
      ` : ''}

      <div class="box">
        <div class="label">Metadata</div>
        <pre>Created At: ${escapeHtml_(String(a.createdAt || ''))}
Updated At: ${escapeHtml_(String(a.updatedAt || ''))}
Version: ${escapeHtml_(String(a.version || 1))}
Created By Email: ${escapeHtml_(String(a.createdByEmail || ''))}
Updated By Email: ${escapeHtml_(String(a.updatedByEmail || ''))}</pre>
      </div>
    </body>
  </html>`;
}

function buildAllArticlesPdfHtml_(articles, label) {
  const rows = (articles || []).map(a => `
    <tr>
      <td>${escapeHtml_(a.title || '')}</td>
      <td>${escapeHtml_(a.category || '')}</td>
      <td>${escapeHtml_(a.tags || '')}</td>
      <td>${escapeHtml_(a.createdBy || '')}</td>
      <td>${escapeHtml_(a.visibility || 'public')}</td>
    </tr>
  `).join('');

  return `
  <html>
    <head>
      <meta charset="UTF-8">
      <style>
        body{font-family:Arial,sans-serif;padding:24px;color:#111827}
        h1{margin:0 0 8px;font-size:22px}
        .meta{color:#6b7280;font-size:12px;margin-bottom:16px}
        table{width:100%;border-collapse:collapse}
        th,td{border:1px solid #e5e7eb;padding:8px;text-align:left;font-size:12px;vertical-align:top}
        th{background:#eff6ff}
      </style>
    </head>
    <body>
      <h1>Knowledge Base Export</h1>
      <div class="meta">Scope: ${escapeHtml_(label || 'All')}</div>
      <table>
        <thead>
          <tr>
            <th>Title</th>
            <th>Category</th>
            <th>Tags</th>
            <th>Created By</th>
            <th>Visibility</th>
          </tr>
        </thead>
        <tbody>
          ${rows || '<tr><td colspan="5">No data</td></tr>'}
        </tbody>
      </table>
    </body>
  </html>`;
}

function getOutputFolder_() {
  const folderId = KB.FOLDER_ID;
  if (folderId && folderId !== 'PASTE_DRIVE_FOLDER_ID_HERE') {
    return DriveApp.getFolderById(folderId);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const file = DriveApp.getFileById(ss.getId());
  const parents = file.getParents();
  const parentFolder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();

  const folderName = 'IT Support Knowledge Base PDFs';
  const subFolders = parentFolder.getFoldersByName(folderName);
  if (subFolders.hasNext()) return subFolders.next();

  return parentFolder.createFolder(folderName);
}

function sanitizeFileName_(name) {
  return String(name || '').replace(/[\\\/:*?"<>|]/g, '-');
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