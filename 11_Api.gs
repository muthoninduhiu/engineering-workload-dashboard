function apiCall(request) {
  request = request || {};
  const action = clean_(request.action);
  const args = Array.isArray(request.args) ? request.args : [];
  const access = readAppSession_(request.sessionToken);
  REQUEST_AUTH_CONTEXT_ = access;
  try {
    const allowed = {
      getBootstrapInfo:getBootstrapInfo,
      getStartupDashboardData:getStartupDashboardData,
      getDashboardData:getDashboardData,
      refreshDashboardData:refreshDashboardData,
      getInitialDashboardData:getInitialDashboardData,
      getWeekData:getWeekData,
      getWeekDataBatch:getWeekDataBatch,
      getTaskById:getTaskById,
      getActivityLog:getActivityLog,
      getTaskHistory:getTaskHistory,
      getAccessAdminData:getAccessAdminData,
      upsertAccessUser:upsertAccessUser,
      deactivateAccessUser:deactivateAccessUser,
      requestAccess:requestAccess,
      getDirectAuditTriggerStatus:getDirectAuditTriggerStatus,
      installDirectEditAuditTriggers:installDirectEditAuditTriggers,
      applySourceDataValidation:applySourceDataValidation,
      updateMasterTaskField:updateMasterTaskField,
      updateTaskById:updateTaskById,
      addTaskIdToWeek:addTaskIdToWeek,
      addTaskToWeek:addTaskToWeek,
      addTasksToWeek:addTasksToWeek,
      updateWeeklyTaskField:updateWeeklyTaskField,
      updateWeeklyTaskFields:updateWeeklyTaskFields,
      removeWeeklyTask:removeWeeklyTask,
      createMasterTask:createMasterTask,
      deleteTask:deleteTask,
      setOwnerCapacity:setOwnerCapacity,
      resetOwnerCapacity:resetOwnerCapacity
    };
    if (!allowed[action]) throw new Error('Unsupported app action: ' + action);
    return allowed[action].apply(null, args);
  } finally {
    REQUEST_AUTH_CONTEXT_ = null;
  }
}
