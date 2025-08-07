// --------------------- LogManager ------------------------ 

/**
 * 日志管理工具类
 */
class LogManager {
  /**
   * 初始化日志表
   * @private
   * @returns {Sheet} 日志表对象
   */
  static _initLogSheet() {
    const ss = SpreadsheetApp.getActive();
    let sheet = ss.getSheetByName(LOG_CONSTANTS.SHEET_NAME);
    
    if (!sheet) {
      sheet = ss.insertSheet(LOG_CONSTANTS.SHEET_NAME);
      sheet.getRange(1, 1, 1, LOG_CONSTANTS.HEADERS.length)
        .setValues([LOG_CONSTANTS.HEADERS])
        .setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    
    return sheet;
  }

  /**
   * 添加日志记录
   * @param {string} type 操作类型（使用 LOG_CONSTANTS.TYPES 中的值）
   * @param {string} sheetName 操作的表名
   * @param {string} action 操作内容
   * @param {string} [details=''] 详细信息（可选）
   */
  static addLog(type, sheetName, action, details = '') {
    const sheet = this._initLogSheet();
    const user = Session.getActiveUser().getEmail();
    const timestamp = new Date().toLocaleString("zh-CN");
    
    const logRow = [
      timestamp,
      type,
      user,
      sheetName,
      action,
      details
    ];
    
    // 在第二行插入新日志（保持表头在第一行）
    sheet.insertRowAfter(1);
    sheet.getRange(2, 1, 1, logRow.length).setValues([logRow]);

    // 检查是否需要清理日志
    this._checkAndCleanupLogs(sheet);
  }

  /**
   * 检查并清理日志
   * @private
   * @param {Sheet} sheet 日志表对象
   */
  static _checkAndCleanupLogs(sheet) {
    const currentRows = sheet.getLastRow();
    const threshold = LOG_CONSTANTS.RETENTION.MAX_ROWS * LOG_CONSTANTS.RETENTION.CLEANUP_THRESHOLD;
    
    // 如果当前行数超过阈值，触发清理
    if (currentRows > threshold) {
      const targetRows = Math.floor(LOG_CONSTANTS.RETENTION.MAX_ROWS * LOG_CONSTANTS.RETENTION.CLEANUP_TARGET);
      const data = sheet.getDataRange().getValues();
      
      // 确保保留表头
      if (data.length <= 1) return;
      
      // 计算最小保留日期
      const minDate = new Date();
      minDate.setDate(minDate.getDate() - LOG_CONSTANTS.RETENTION.MIN_DAYS);
      
      // 从后往前查找需要保留的最后一行
      let deleteFromRow = data.length;
      let foundDeleteRow = false;
      
      for (let i = data.length - 1; i > 1; i--) {
        const logDate = new Date(data[i][0]);
        
        // 如果找到了一行需要删除的数据（在最小保留日期之前，且超出目标行数）
        if (logDate < minDate && i > targetRows) {
          deleteFromRow = i;
          foundDeleteRow = true;
          break;
        }
      }
      
      // 如果需要删除行
      if (deleteFromRow < data.length) {
        const rowsToDelete = data.length - deleteFromRow;
        sheet.deleteRows(deleteFromRow + 1, rowsToDelete);
        
        // 记录清理操作（插入到第二行）
        const newLog = [
          new Date().toLocaleString("zh-CN"),
          "系统维护",
          "系统",
          LOG_CONSTANTS.SHEET_NAME,
          "日志清理",
          `清理了 ${rowsToDelete} 条历史日志记录`
        ];
        sheet.insertRowAfter(1);
        sheet.getRange(2, 1, 1, newLog.length).setValues([newLog]);
      }
    }
  }

  /**
   * 手动触发日志清理
   * @param {number} [days=30] 保留天数
   */
  static manualCleanup(days = LOG_CONSTANTS.RETENTION.MIN_DAYS) {
    const sheet = this._initLogSheet();
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) return; // 只有表头或空表，直接返回
    
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - days);
    
    // 从后往前查找需要删除的行
    let deleteFromRow = data.length;
    for (let i = data.length - 1; i > 1; i--) {
      const logDate = new Date(data[i][0]);
      if (logDate < cutoffDate) {
        deleteFromRow = i;
        break;
      }
    }
    
    if (deleteFromRow < data.length) {
      const rowsToDelete = data.length - deleteFromRow;
      sheet.deleteRows(deleteFromRow + 1, rowsToDelete);
      
      // 记录清理操作（插入到第二行）
      const newLog = [
        new Date().toLocaleString("zh-CN"),
        "系统维护",
        "系统",
        LOG_CONSTANTS.SHEET_NAME,
        "手动日志清理",
        `清理了 ${rowsToDelete} 条${days}天前的历史日志记录`
      ];
      sheet.insertRowAfter(1);
      sheet.getRange(2, 1, 1, newLog.length).setValues([newLog]);
    }
  }
}

// --------------------- LogManager ------------------------