/**
 * 执行合并操作
 * @param {Object} config 合并配置
 * @param {Sheet} targetSheet 目标表格，如果不指定则使用config中的targetSheet
 * @returns {Object} 合并结果
 */
function mergeSheets(config, targetSheet) {
  if (!config || !config.sourceSheet || (!config.targetSheet && !targetSheet)) {
    return {
      success: false,
      message: "配置参数无效"
    };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName(config.sourceSheet);
  var targetSheetName = targetSheet ? targetSheet.getName() : config.targetSheet;
  var targetSheet = targetSheet ? targetSheet : ss.getSheetByName(targetSheetName);
  
  if (!sourceSheet || !targetSheet) {
    return {
      success: false,
      message: "未找到指定的表格，请检查表格名称"
    };
  }

  try {
    // 获取三个表的数据：源表、目标表和基准表
    var sourceRange = sourceSheet.getDataRange();
    var targetRange = targetSheet.getDataRange();
    
    var sourceData = sourceRange.getValues();
    var targetData = targetRange.getValues();
    
    // 获取源表的注释（用于获取基准值）
    var sourceNotes = sourceRange.getNotes();
    var sourceBaseData = sourceNotes.map(row => 
      row.map(note => {
        var baseValue = NoteManager.getSystemNote(note, NOTE_CONSTANTS.TYPES.BASE_VALUE);
        return baseValue;
      })
    );
    
    // 获取表头
    var sourceHeaders = sourceData[0];
    var targetHeaders = targetData[0];
    
    // 找到ID列
    var sourceIdColIndex = -1;
    var targetIdColIndex = -1;
    
    sourceHeaders.forEach((header, index) => {
      if (header.toString().endsWith(MERGE_CONSTANTS.ID_SUFFIX)) {
        sourceIdColIndex = index;
      }
    });
    
    targetHeaders.forEach((header, index) => {
      if (header.toString().endsWith(MERGE_CONSTANTS.ID_SUFFIX)) {
        targetIdColIndex = index;
      }
    });
    
    if (sourceIdColIndex === -1 || targetIdColIndex === -1) {
      return {
        success: false,
        message: `未找到ID列（以${MERGE_CONSTANTS.ID_SUFFIX}结尾的列）`
      };
    }

    // 创建表头映射
    var headerMap = {};
    sourceHeaders.forEach((header, index) => {
      headerMap[header] = {sourceIndex: index, targetIndex: -1};
    });
    
    // 检查新增列
    var newColumns = [];
    sourceHeaders.forEach((header, index) => {
      if (!targetHeaders.includes(header) && !header.toString().endsWith(MERGE_CONSTANTS.ID_SUFFIX)) {
        newColumns.push({
          header: header,
          sourceIndex: index
        });
      }
    });

    // 如果有新增列，在目标表和预览表中添加这些列
    if (newColumns.length > 0) {
      // 在目标表最后添加新列
      targetHeaders = targetHeaders.concat(newColumns.map(col => col.header));
      targetSheet.getRange(1, targetHeaders.length - newColumns.length + 1, 1, newColumns.length)
        .setValues([newColumns.map(col => col.header)])
        .setBackground(MERGE_CONSTANTS.COLORS.NEW);
      
      // 更新headerMap
      newColumns.forEach((col, idx) => {
        headerMap[col.header].targetIndex = targetHeaders.length - newColumns.length + idx;
      });
      
      // 为新列添加空值
      var emptyColumns = Array(newColumns.length).fill('');
      for (var i = 1; i < targetData.length; i++) {
        targetSheet.getRange(i + 1, targetHeaders.length - newColumns.length + 1, 1, newColumns.length)
          .setValues([emptyColumns]);
      }
      
      // 更新targetData以包含新列
      targetData = targetSheet.getDataRange().getValues();
    }

    targetHeaders.forEach((header, index) => {
      if (headerMap[header]) {
        headerMap[header].targetIndex = index;
      }
    });

    // 将目标表数据转换为以ID为键的Map
    var targetDataMap = new Map();
    for (var i = 1; i < targetData.length; i++) {
      var id = targetData[i][targetIdColIndex];
      if (id) {
        targetDataMap.set(id.toString(), {
          rowIndex: i,
          data: targetData[i],
        });
      }
    }

    // 记录需要处理的变更
    var changes = {
      newRows: [],
      updates: [], // 新增：记录可以直接更新的行
      conflicts: []
    };

    // 处理源表数据
    for (var i = 1; i < sourceData.length; i++) {
      var sourceRow = sourceData[i];
      var id = sourceRow[sourceIdColIndex];
      
      if (!id) continue; // 跳过空ID行
      
      id = id.toString();
      var targetRow = targetDataMap.get(id);
      
      if (!targetRow) {
        // 新行，直接添加到新行列表
        changes.newRows.push(sourceRow);
      } else {
        // 检查修改情况
        var hasConflict = false;
        var conflictColumns = [];
        var updateColumns = [];
        
        for (var header in headerMap) {
          var sourceIndex = headerMap[header].sourceIndex;
          var targetIndex = headerMap[header].targetIndex;
          
          if (targetIndex === -1) continue; // 跳过目标表中不存在的列
          
          // 目标表中的当前值
          var currentValue = targetRow.data[targetIndex];
          // 当前表的当前值
          var sourceValue = sourceRow[sourceIndex];
          // 当前表中标记的 base
          var sourceBaseValue = sourceBaseData[i][sourceIndex];
          
          // 标准化值
          function normalizeValue(value) {
            if (value === null || value === undefined) return '';
            
            // 如果是数字或者可以转换为数字
            const num = Number(value);
            if (!isNaN(num)) {
              // 对于整数，返回整数字符串
              if (Number.isInteger(num)) {
                return String(num);
              }
              // 对于小数，统一格式化（去除末尾的0）
              return String(parseFloat(num.toFixed(10)));
            }
            
            // 非数字类型，转换为字符串
            return String(value);
          }
          
          // 检查源表和目标表是否都进行了修改
          var sourceModified = sourceBaseValue && normalizeValue(sourceValue) !== normalizeValue(sourceBaseValue);
          if (sourceModified && normalizeValue(sourceBaseValue) !== normalizeValue(currentValue) && normalizeValue(sourceValue) !== normalizeValue(currentValue) ) {
            hasConflict = true;
            conflictColumns.push({
              header: header,
              sourceValue: sourceValue,
              targetValue: currentValue,
              sourceBaseValue: sourceBaseValue,
            });
          } else if (sourceModified) {
            updateColumns.push({
              header: header,
              sourceValue: sourceValue,
              baseValue: sourceValue // 更新基准值为新的源值
            });
          }
        }
        
        if (hasConflict) {
          changes.conflicts.push({
            id: id,
            sourceRowIndex: i,
            targetRowIndex: targetRow.rowIndex,
            columns: conflictColumns
          });
        } else if (updateColumns.length > 0) {
          changes.updates.push({
            id: id,
            sourceRowIndex: i,
            targetRowIndex: targetRow.rowIndex,
            columns: updateColumns
          });
        }
      }
    }

    // 处理变更
    // 1. 添加新行（按ID顺序插入）
    if (changes.newRows.length > 0) {
      const insertResult = batchInsertRowsInOrder(
        targetSheet,
        changes.newRows,
        targetIdColIndex,
        {
          newRowColor: MERGE_CONSTANTS.COLORS.NEW,
          addBaseNotes: true
        }
      );
      
      if (!insertResult.success) {
        throw new Error(`无法插入新行: ${insertResult.message}`);
      }
    }

    // 2. 处理可以直接更新的行
    changes.updates.forEach(update => {
      update.columns.forEach(col => {
        var targetIndex = headerMap[col.header].targetIndex;
        var range = targetSheet.getRange(update.targetRowIndex + 1, targetIndex + 1);
        range.setValue(col.sourceValue);
        range.setBackground(MERGE_CONSTANTS.COLORS.UPDATED);
        
        // 更新基准值注释
        const currentNote = range.getNote();
        const newNote = NoteManager.addSystemNote(
          NoteManager.removeSystemNote(currentNote, NOTE_CONSTANTS.TYPES.BASE_VALUE),
          NOTE_CONSTANTS.TYPES.BASE_VALUE,
          col.sourceValue.toString()
        );
        range.setNote(newNote);
      });
    });

    // 3. 标记冲突
    changes.conflicts.forEach(conflict => {
      conflict.columns.forEach(col => {
        var targetIndex = headerMap[col.header].targetIndex;
        var range = targetSheet.getRange(conflict.targetRowIndex + 1, targetIndex + 1);
        range.setBackground(MERGE_CONSTANTS.COLORS.CONFLICT);
        
        const conflictInfo = `${config.targetSheet}: ${col.targetValue}\n${config.sourceSheet}: ${col.sourceValue}`;
        const currentNote = range.getNote();
        const newNote = NoteManager.addSystemNote(
          NoteManager.removeSystemNote(currentNote, NOTE_CONSTANTS.TYPES.CONFLICT),
          NOTE_CONSTANTS.TYPES.CONFLICT,
          conflictInfo
        );
        range.setNote(newNote);
      });
    });

    // 在合并成功后记录日志
    LogManager.addLog(
      LOG_CONSTANTS.TYPES.MERGE,
      config.sourceSheet,
      "生成合并预览成功",
      `目标表格：${targetSheetName}\n` +
      `新增行数：${changes.newRows.length}\n` +
      `更新行数：${changes.updates.length}\n` +
      `冲突行数：${changes.conflicts.length}`
    );

    return {
      success: true,
      message: `合并完成\n新增行数: ${changes.newRows.length}\n更新行数: ${changes.updates.length}\n冲突行数: ${changes.conflicts.length}`,
      changes: changes
    };
  } catch (error) {
    return {
      success: false,
      message: "合并过程中出错: " + error.toString()
    };
  }
}

/**
 * 从注释中提取基准值
 */
function extractBaseValue(note) {
  return NoteManager.getSystemNote(note, NOTE_CONSTANTS.TYPES.BASE_VALUE);
}

/**
 * 确认合并预览表到目标表
 * @param {string} sourceSheetName 源表格名称
 * @param {string} targetSheetName 目标表格名称
 * @param {string} previewSheetName 预览表格名称
 * @returns {Object} 合并结果
 */
function confirmMergeFromPreview(sourceSheetName, targetSheetName, previewSheetName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sourceSheet = ss.getSheetByName(sourceSheetName);
    var targetSheet = ss.getSheetByName(targetSheetName);
    var previewSheet = ss.getSheetByName(previewSheetName);
    
    if (!sourceSheet || !targetSheet || !previewSheet) {
      return {
        success: false,
        message: "未找到指定的表格，请检查表格名称"
      };
    }

    // 获取预览表的数据和注释
    var previewRange = previewSheet.getDataRange();
    var previewData = previewRange.getValues();
    var previewNotes = previewRange.getNotes();
    var previewBackgrounds = previewRange.getBackgrounds();
    
    // 检查是否存在未解决的冲突
    var hasUnresolvedConflicts = false;
    for (var i = 0; i < previewNotes.length; i++) {
      for (var j = 0; j < previewNotes[i].length; j++) {
        var note = previewNotes[i][j];
        var conflictInfo = NoteManager.getSystemNote(note, NOTE_CONSTANTS.TYPES.CONFLICT);
        if (conflictInfo) {
          hasUnresolvedConflicts = true;
          break;
        }
      }
      if (hasUnresolvedConflicts) break;
    }
    
    if (hasUnresolvedConflicts) {
      return {
        success: false,
        message: "存在未解决的冲突，请先解决所有冲突后再确认合并"
      };
    }

    // 找到ID列以识别新行
    var idColIndex = -1;
    var headers = previewData[0];
    headers.forEach((header, index) => {
      if (header.toString().endsWith(MERGE_CONSTANTS.ID_SUFFIX)) {
        idColIndex = index;
      }
    });
    
    if (idColIndex === -1) {
      return {
        success: false,
        message: `未找到ID列（以${MERGE_CONSTANTS.ID_SUFFIX}结尾的列）`
      };
    }

    // 获取变更行信息
    var rowsWithChanges = [];
    var newRowData = [];
    
    // 记录所有有变化的行
    for (var i = 0; i < previewData.length; i++) {
      var hasChanges = false;
      for (var j = 0; j < previewData[i].length; j++) {
        var background = previewBackgrounds[i][j];
        if (background === MERGE_CONSTANTS.COLORS.NEW || 
            background === MERGE_CONSTANTS.COLORS.UPDATED || 
            background === MERGE_CONSTANTS.COLORS.RESOLVED) {
          hasChanges = true;
        }
      }
      
      if (hasChanges) {
        if (i > 0 && previewBackgrounds[i][0] === MERGE_CONSTANTS.COLORS.NEW) {
          // 新行，需要在目标表中插入
          newRowData.push({
            rowIndex: i,
            id: previewData[i][idColIndex],
            data: previewData[i],
            notes: previewNotes[i]
          });
        } else {
          // 更新的现有行
          rowsWithChanges.push(i);
        }
      }
    }
    
    // 对目标表的修改
    
    // 1. 先处理现有行的更新（不会改变行数）
    for (var i = 0; i < rowsWithChanges.length; i++) {
      var rowIndex = rowsWithChanges[i];
      for (var j = 0; j < previewData[rowIndex].length; j++) {
        var background = previewBackgrounds[rowIndex][j];
        // 检查单元格是否有变化
        if (background === MERGE_CONSTANTS.COLORS.UPDATED || 
            background === MERGE_CONSTANTS.COLORS.RESOLVED) {
          var targetCell = targetSheet.getRange(rowIndex + 1, j + 1);
          targetCell.setValue(previewData[rowIndex][j]);
          targetCell.setNote(previewNotes[rowIndex][j]);
          targetCell.setBackground(MERGE_CONSTANTS.COLORS.MERGED);
        }
      }
    }
    
    // 2. 获取目标表现有数据
    var targetData = targetSheet.getDataRange().getValues();
    
    // 3. 处理新行插入
    if (newRowData.length > 0) {
      // 将newRowData格式转换为batchInsertRowsInOrder所需格式
      const formattedNewRows = newRowData.map(row => row.data);
      
      const insertResult = batchInsertRowsInOrder(
        targetSheet,
        formattedNewRows,
        idColIndex,
        {
          newRowColor: MERGE_CONSTANTS.COLORS.MERGED,
          addBaseNotes: true
        }
      );
      
      if (!insertResult.success) {
        throw new Error(`确认合并时无法插入新行: ${insertResult.message}`);
      }
    }

    // 标记源表为已合并
    var headerRow = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
    var statusColIndex = -1;
    
    // 在源表格名称后添加"(已合并)"标记
    const newSourceSheetName = sourceSheetName.endsWith('(已合并)') 
      ? sourceSheetName 
      : `${sourceSheetName}(已合并)`;
    sourceSheet.setName(newSourceSheetName);
    
    // 删除预览表
    ss.deleteSheet(previewSheet);
    
    // 清除预览状态
    const cache = CacheService.getScriptCache();
    cache.remove('merge_preview_state');
    
    // 激活目标页签
    targetSheet.activate();
    
    // 记录确认合并成功的日志
    LogManager.addLog(
      LOG_CONSTANTS.TYPES.MERGE,
      sourceSheetName,
      "确认合并成功",
      `目标表格：${targetSheetName}`
    );
    
    return {
      success: true,
      message: "合并完成！只更新了变更的单元格。源表已标记为已合并状态。",
      newRowCount: newRowData.length
    };
    
  } catch (error) {
    console.error('确认合并失败:', error);
    return {
      success: false,
      message: "确认合并过程中出错: " + error.toString()
    };
  }
}

/**
 * 获取预览状态
 * @param {string} sourceSheet 源表格名称
 * @param {string} targetSheet 目标表格名称
 * @returns {Object|null} 预览状态对象，如果没有则返回null
 */
function getPreviewState(sourceSheet, targetSheet) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const previewSheetName = `${sourceSheet} -> ${targetSheet} 合并预览`;
    const previewSheet = ss.getSheetByName(previewSheetName);
    
    if (!previewSheet) {
      return null;
    }
    
    return {
      previewSheetName: previewSheetName,
      sourceSheet: sourceSheet,
      targetSheet: targetSheet
    };
  } catch (e) {
    console.error('获取预览状态失败:', e);
    return null;
  }
}

/**
 * 显示确认对话框
 */
function showConfirmDialog() {
  try {
    // 获取当前活动页签
    const activeSheet = SpreadsheetApp.getActiveSheet();
    const sheetName = activeSheet.getName();
    console.log('当前页签名称:', sheetName);
    
    // 从页签名称中解析源表和目标表
    // 预览页签的命名格式为: "sourceSheet -> targetSheet 合并预览"
    const match = sheetName.match(/^(.*?)\s*->\s*(.*?)\s*合并预览$/);
    console.log('页签名称匹配结果:', match);
    
    if (!match) {
      showAlert('请在合并预览页签中使用此功能');
      return;
    }
    
    const [_, sourceSheet, targetSheet] = match;
    console.log('解析出的源表和目标表:', { sourceSheet, targetSheet });
    
    // 检查预览状态
    const previewState = getPreviewState(sourceSheet.trim(), targetSheet.trim());
    console.log('获取到的预览状态:', previewState);
    
    if (!previewState) {
      showAlert('未找到有效的预览状态，请重新执行合并预览');
      return;
    }
    
    console.log('准备显示对话框');
    // 显示合并对话框
    const ui = SpreadsheetApp.getUi();
    const html = HtmlService.createHtmlOutputFromFile('MergeDialog')
      .setWidth(600)
      .setHeight(600)
      .setTitle('合并表格');
    
    ui.showModalDialog(html, '合并表格');
    console.log('对话框已显示');
  } catch (error) {
    console.error('显示确认对话框失败:', error);
    showAlert('显示确认对话框失败: ' + error.toString());
  }
}

/**
 * 预览合并结果
 * @param {Object} config 合并配置
 * @returns {Object} 预览结果
 */
function previewMerge(config) {
  if (!config || !config.sourceSheet || !config.targetSheet) {
    return {
      success: false,
      message: "配置参数无效"
    };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 检查是否已存在预览状态
  const existingPreview = getPreviewState(config.sourceSheet, config.targetSheet);
  if (existingPreview) {
    return {
      success: false,
      message: `已存在 "${config.sourceSheet} -> ${config.targetSheet}" 的预览，请先完成或取消现有预览`
    };
  }

  Logger.log('解析预合并参数: %s', config);
  var sourceSheet = ss.getSheetByName(config.sourceSheet);
  var targetSheet = ss.getSheetByName(config.targetSheet);
  
  if (!sourceSheet || !targetSheet) {
    return {
      success: false,
      message: "未找到指定的表格，请检查表格名称"
    };
  }

  try {
    // 创建预览表格
    var previewSheet = createPreviewSheet(config.targetSheet, config.sourceSheet);
    
    // 复制目标表格的数据到预览表格，但不包括背景色和系统注释
    var targetRange = targetSheet.getDataRange();
    var targetData = targetRange.getValues();
    
    // 只复制数据，不复制格式和注释
    previewSheet.getRange(1, 1, targetData.length, targetData[0].length)
      .setValues(targetData);

    // 执行合并预览
    var result = mergeSheets(config, previewSheet);
    
    if (result.success) {
      return {
        success: true,
        previewSheetName: previewSheet.getName(),
        sourceSheet: config.sourceSheet,
        targetSheet: config.targetSheet,
        changes: result.changes,
        message: `预览已生成，请在"${previewSheet.getName()}"表格中查看\n${result.message}`
      };
    } else {
      // 如果预览失败，删除预览表格
      deletePreviewSheet(previewSheet.getName());
      return result;
    }
  } catch (error) {
    console.error('预览失败:', error);
    return {
      success: false,
      message: "预览生成失败：" + error.toString()
    };
  }
}

/**
 * 删除预览表格
 * @param {string} previewSheetName 预览表格名称
 * @returns {boolean} 是否成功删除
 */
function deletePreviewSheet(previewSheetName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(previewSheetName);
    if (sheet) {
      ss.deleteSheet(sheet);
      return true;
    }
    return false;
  } catch (error) {
    console.error('删除预览表格失败:', error);
    return false;
  }
}

/**
 * 解决合并冲突
 * @param {Object} config 解决配置
 * @returns {Object} 操作结果
 */
function resolveConflict(config) {
  if (!config || !config.row || !config.header || config.value === undefined) {
    return {
      success: false,
      message: "参数无效"
    };
  }

  try {
    var sheet = SpreadsheetApp.getActiveSheet();
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // 找到对应的列
    var colIndex = headers.findIndex(h => h === config.header);
    if (colIndex === -1) {
      return {
        success: false,
        message: "未找到指定列: " + config.header
      };
    }

    // 更新单元格值
    var cell = sheet.getRange(config.row + 1, colIndex + 1);
    cell.setValue(config.value);
    cell.setBackground(MERGE_CONSTANTS.COLORS.RESOLVED);
    
    // 更新注释：清除冲突信息，更新基准值
    const currentNote = cell.getNote();
    let newNote = NoteManager.removeSystemNote(currentNote, NOTE_CONSTANTS.TYPES.CONFLICT);
    newNote = NoteManager.addSystemNote(
      newNote,
      NOTE_CONSTANTS.TYPES.BASE_VALUE,
      config.value.toString()
    );
    cell.setNote(newNote);

    return {
      success: true,
      message: "已更新单元格值"
    };
  } catch (error) {
    return {
      success: false,
      message: "更新失败: " + error.toString()
    };
  }
}

/**
 * 显示提示信息
 * @param {string} message 提示信息
 */
function showAlert(message) {
  SpreadsheetApp.getUi().alert(message);
}

/**
 * 显示对话框
 * @param {string} [dialogType='merge'] 对话框类型
 */
function showDialog(dialogType = 'merge') {
  // 创建新的对话框实例
  var html = HtmlService.createHtmlOutputFromFile('MergeDialog')
    .setWidth(600)
    .setHeight(600)
    .setTitle('合并表格');
  
  // 使用showModalDialog而不是showDialog以确保对话框总是在前面
  SpreadsheetApp.getUi().showModalDialog(html, '合并表格');
}

/**
 * 创建预览表格
 * @param {string} targetSheetName 目标表格名称
 * @param {string} sourceSheetName 源表格名称
 * @returns {Sheet} 预览表格
 */
function createPreviewSheet(targetSheetName, sourceSheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var previewName = `${sourceSheetName} -> ${targetSheetName} 合并预览`;
  var existingSheet = ss.getSheetByName(previewName);
  if (existingSheet) {
    ss.deleteSheet(existingSheet);
  }
  return ss.insertSheet(previewName);
}

/**
 * 显示合并对话框
 */
function showMergeDialog() {
  // 检查当前表格是否已经标记为已合并
  var currentSheet = SpreadsheetApp.getActiveSheet();
  var currentSheetName = currentSheet.getName();
  
  if (currentSheetName.endsWith('(已合并)')) {
    SpreadsheetApp.getUi().alert(
      '无法合并',
      '当前表格已经标记为已合并状态，不能重复发起合并。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  var html = HtmlService.createHtmlOutputFromFile('MergeDialog')
    .setWidth(500)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, '表格合并工具');
}

// 用于比较ID的辅助函数
function compareIds(idA, idB) {
  // 尝试将ID解析为数字进行比较
  const numA = parseFloat(idA);
  const numB = parseFloat(idB);
  
  // 如果都是有效数字，按数字大小排序
  if (!isNaN(numA) && !isNaN(numB)) {
    return numA - numB;
  }
  
  // 否则按字符串排序
  return idA.localeCompare(idB);
}

/**
 * 优化的ID排序和行插入函数 - 在预览和确认合并时共用
 * @param {Sheet} sheet 目标表格
 * @param {Array} newRows 要插入的新行
 * @param {number} idColIndex ID列索引
 * @param {Object} options 可选配置项
 * @returns {Object} 处理结果
 */
function batchInsertRowsInOrder(sheet, newRows, idColIndex, options = {}) {
  if (newRows.length === 0) return { success: true, message: "没有新行需要插入" };
  
  try {
    const defaults = {
      preserveFormatting: true,    // 是否保留现有格式
      newRowColor: MERGE_CONSTANTS.COLORS.NEW,  // 新行的背景色
      addBaseNotes: true,          // 是否添加基准值注释
      batchSize: 1000              // 批量处理的最大行数
    };
    
    const config = { ...defaults, ...options };
    
    // 获取当前表所有数据、注释和格式
    const currentData = sheet.getDataRange().getValues();
    const currentNotes = config.addBaseNotes ? sheet.getDataRange().getNotes() : null;
    const currentBackgrounds = config.preserveFormatting ? sheet.getDataRange().getBackgrounds() : null;
    
    const headerRow = currentData[0];
    const dataRows = currentData.slice(1); // 不包括表头
    
    // 计算所需的列数(取最大值)
    const maxColumns = Math.max(
      headerRow.length,
      ...newRows.map(row => row.length)
    );
    
    // 提取所有已有ID
    const existingIds = new Map();
    dataRows.forEach((row, idx) => {
      const id = row[idColIndex]?.toString();
      if (id) {
        existingIds.set(id, {
          data: row,
          index: idx + 1, // 实际行号(从1开始)，不包括表头
          notes: config.addBaseNotes ? currentNotes[idx + 1] : null,
          background: config.preserveFormatting ? currentBackgrounds[idx + 1] : null
        });
      }
    });
    
    // 准备新行数据和ID
    const newRowsMap = new Map();
    newRows.forEach((row, idx) => {
      const id = row[idColIndex]?.toString();
      if (id) {
        // 确保行长度一致
        const paddedRow = [...row];
        while (paddedRow.length < maxColumns) {
          paddedRow.push('');
        }
        
        newRowsMap.set(id, {
          data: paddedRow,
          originalIndex: idx
        });
      }
    });
    
    // 合并所有唯一ID并排序
    const allIds = [...new Set([...existingIds.keys(), ...newRowsMap.keys()])];
    allIds.sort(compareIds);
    
    // 创建完整的排序数据集
    const sortedData = [headerRow]; // 先放入表头
    const sortedNotes = config.addBaseNotes ? [currentNotes[0]] : null; 
    const sortedBackgrounds = config.preserveFormatting ? [currentBackgrounds[0]] : null;
    const newRowIndices = [];
    
    allIds.forEach((id, idx) => {
      const rowIndex = idx + 1; // 实际行索引（从0开始，不包括表头）
      
      if (newRowsMap.has(id)) {
        // 这是一个新行
        const newRow = newRowsMap.get(id);
        sortedData.push(newRow.data);
        newRowIndices.push(rowIndex + 1); // +1 是因为包括表头
        
        if (config.addBaseNotes) {
          // 为新行创建基准值注释
          const rowNotes = newRow.data.map(value => 
            NoteManager.addSystemNote('', NOTE_CONSTANTS.TYPES.BASE_VALUE, 
              value !== null && value !== undefined ? value.toString() : '')
          );
          sortedNotes.push(rowNotes);
        }
        
        if (config.preserveFormatting) {
          // 为新行准备背景色
          const rowBackground = Array(maxColumns).fill(config.newRowColor);
          sortedBackgrounds.push(rowBackground);
        }
      } else if (existingIds.has(id)) {
        // 这是现有行
        const existingRow = existingIds.get(id);
        
        // 确保行长度一致
        const paddedRow = [...existingRow.data];
        while (paddedRow.length < maxColumns) {
          paddedRow.push('');
        }
        sortedData.push(paddedRow);
        
        if (config.addBaseNotes) {
          sortedNotes.push(existingRow.notes);
        }
        
        if (config.preserveFormatting) {
          sortedBackgrounds.push(existingRow.background);
        }
      }
    });
    
    // 清除并批量写入数据
    if (sortedData.length > config.batchSize) {
      // 对于大数据集，分批处理
      const batchCount = Math.ceil(sortedData.length / config.batchSize);
      
      // 先调整表格大小以适应所有数据
      sheet.clear();
      if (sheet.getMaxRows() < sortedData.length) {
        sheet.insertRows(1, sortedData.length - sheet.getMaxRows());
      }
      if (sheet.getMaxColumns() < maxColumns) {
        sheet.insertColumns(1, maxColumns - sheet.getMaxColumns());
      }
      
      // 分批写入数据
      for (let i = 0; i < batchCount; i++) {
        const startRow = i * config.batchSize + 1;
        const batchRowCount = Math.min(config.batchSize, sortedData.length - i * config.batchSize);
        
        if (batchRowCount <= 0) break;
        
        const batchData = sortedData.slice(startRow - 1, startRow - 1 + batchRowCount);
        sheet.getRange(startRow, 1, batchRowCount, maxColumns).setValues(batchData);
        
        if (config.addBaseNotes) {
          const batchNotes = sortedNotes.slice(startRow - 1, startRow - 1 + batchRowCount);
          sheet.getRange(startRow, 1, batchRowCount, maxColumns).setNotes(batchNotes);
        }
        
        if (config.preserveFormatting) {
          const batchBackgrounds = sortedBackgrounds.slice(startRow - 1, startRow - 1 + batchRowCount);
          sheet.getRange(startRow, 1, batchRowCount, maxColumns).setBackgrounds(batchBackgrounds);
        }
      }
    } else {
      // 对于小数据集，一次性处理
      sheet.clear();
      if (sheet.getMaxRows() < sortedData.length) {
        sheet.insertRows(1, sortedData.length - sheet.getMaxRows());
      }
      if (sheet.getMaxColumns() < maxColumns) {
        sheet.insertColumns(1, maxColumns - sheet.getMaxColumns());
      }
      
      sheet.getRange(1, 1, sortedData.length, maxColumns).setValues(sortedData);
      
      if (config.addBaseNotes) {
        sheet.getRange(1, 1, sortedData.length, maxColumns).setNotes(sortedNotes);
      }
      
      if (config.preserveFormatting) {
        sheet.getRange(1, 1, sortedData.length, maxColumns).setBackgrounds(sortedBackgrounds);
      }
    }
    
    return {
      success: true,
      message: `成功按ID顺序插入了 ${newRowIndices.length} 行`,
      newRowCount: newRowIndices.length,
      sortedData: sortedData,
      newRowIndices: newRowIndices
    };
  } catch (error) {
    console.error('批量插入行失败:', error);
    return {
      success: false,
      message: `批量插入行失败: ${error.toString()}`
    };
  }
}
