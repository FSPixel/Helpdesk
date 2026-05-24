function getCurrentRole_() {
  return getCurrentUser_().role || 'viewer';
}

function requireRole_(roles) {
  const role = getCurrentRole_();
  if (!roles.includes(role)) {
    throw new Error(`Access denied for role: ${role}`);
  }
}

function canViewArticle_(article, email, role) {
  if (role === 'admin' || role === 'editor' || role === 'auditor') return true;
  if (article.visibility === 'public') return true;
  return String(article.createdByEmail || '').toLowerCase() === String(email || '').toLowerCase();
}

function canEditArticle_(article, email, role) {
  if (role === 'admin') return true;
  if (role === 'editor' && String(article.createdByEmail || '').toLowerCase() === String(email || '').toLowerCase()) return true;
  return false;
}