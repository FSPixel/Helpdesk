function exportArticlePdf(id) {
  ensureSetup_();

  const role = getCurrentRole_();
  if (!['admin', 'editor', 'auditor'].includes(role)) {
    throw new Error('No permission to export PDF');
  }

  const article = getArticleById(id);
  if (!article) throw new Error('Article not found');

  const html = buildArticlePdfHtml_(article);
  const blob = HtmlService.createHtmlOutput(html)
    .getBlob()
    .getAs(MimeType.PDF);

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

  const blob = HtmlService.createHtmlOutput(html)
    .getBlob()
    .getAs(MimeType.PDF);

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
  const ssId = ss.getId();
  const parentFolder = DriveApp.getFileById(ssId).getParents().next();
  const folderName = 'IT Support Knowledge Base PDFs';

  const subFolders = parentFolder.getFolders();
  while (subFolders.hasNext()) {
    const folder = subFolders.next();
    if (folder.getName() === folderName) return folder;
  }

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