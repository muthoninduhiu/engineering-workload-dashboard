function ensureTaskRelationSheet_(ss) {
  let sheet = ss.getSheetByName(CONFIG.TASK_RELATION_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.TASK_RELATION_SHEET);
    sheet.getRange(1, 1, 1, 5).setValues([['Task ID', 'Related Task ID', 'Relationship', 'Updated At', 'Updated By']]).setFontWeight('bold');
  }
  return sheet;
}

function getTaskRelations(taskId) {
  requireViewer_();
  const id = clean_(taskId);
  if (!id) return [];
  const sheet = ensureTaskRelationSheet_(getDashboardSpreadsheet_());
  if (sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues()
    .filter(row => clean_(row[0]).toLowerCase() === id.toLowerCase() || clean_(row[1]).toLowerCase() === id.toLowerCase())
    .map(row => ({taskId: clean_(row[0]), relatedTaskId: clean_(row[1]), relationship: clean_(row[2]), updatedAt: row[3] instanceof Date ? row[3].toISOString() : clean_(row[3])}));
}

function saveTaskRelation(taskId, relatedTaskId, relationship) {
  const actor = requireEditor_();
  const id = clean_(taskId), related = clean_(relatedTaskId);
  const allowed = ['Parent', 'Subtask', 'Related'];
  if (!id || !related || id.toLowerCase() === related.toLowerCase()) throw new Error('Choose a different task to link.');
  if (allowed.indexOf(relationship) === -1) throw new Error('Choose Parent, Subtask, or Related.');
  const sheet = ensureTaskRelationSheet_(getDashboardSpreadsheet_());
  const rows = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getDisplayValues() : [];
  const exists = rows.some(row => clean_(row[0]).toLowerCase() === id.toLowerCase() && clean_(row[1]).toLowerCase() === related.toLowerCase() && clean_(row[2]) === relationship);
  if (!exists) sheet.appendRow([id, related, relationship, new Date(), actor.email || '']);
  return getTaskRelations(id);
}
