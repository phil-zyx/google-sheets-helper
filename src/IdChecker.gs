/**
 * 极简高效的ID冲突检查
 */
function checkSingleIdConflictImproved({ value, sheet: sheetName, row, column, columnName }) {
  const startTime = Date.now();
  console.log(`🔍 [ID检查] ${sheetName} 第${row}行 - "${value}"`);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentSheet = ss.getSheetByName(sheetName);
  
  // 1. 快速检查当前表格
  const currentConflict = fastCheckCurrentSheet(currentSheet, value, row, columnName);
  if (currentConflict) {
    console.log(`🚨 [当前表格冲突] 耗时: ${Date.now() - startTime}ms`);
    return [currentConflict];
  }
  
  // 2. 批量检查其他表格
  const crossConflict = fastCheckOtherSheets(ss, currentSheet, sheetName, value, row, columnName);
  
  console.log(`✅ [检查完成] 耗时: ${Date.now() - startTime}ms`);
  return crossConflict ? [crossConflict] : [];
}

/**
 * 快速检查当前表格
 */
function fastCheckCurrentSheet(sheet, value, currentRow, columnName) {
  // 批量获取表头和数据
  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1 || lastCol === 0) return null;
  
  // 一次性读取表头
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const columnIndex = headers.findIndex(header => header && header.toString() === columnName);
  
  if (columnIndex === -1) return null;
  
  // 一次性读取整列ID数据
  const idColumn = sheet.getRange(2, columnIndex + 1, lastRow - 1, 1).getValues().flat();
  
  // 快速遍历查找冲突
  for (let i = 0; i < idColumn.length; i++) {
    const actualRow = i + 2;
    const id = idColumn[i];
    
    if (actualRow !== currentRow && 
        id && 
        id.toString().trim() && 
        id.toString() === value.toString()) {
      
      return {
        sheet: sheet.getName(),
        row: actualRow,
        column: columnIndex + 1
      };
    }
  }
  
  return null;
}

/**
 * 快速检查其他表格
 */
function fastCheckOtherSheets(ss, currentSheet, currentSheetName, value, currentRow, columnName) {
  const sheets = ss.getSheets();
  let currentRowData = null;
  let currentHeaders = null;
  
  for (const sheet of sheets) {
    if (sheet.getName() === currentSheetName) continue;
    
    const conflict = fastCheckSingleOtherSheet(sheet, value, columnName);
    if (!conflict) continue;
    
    // 延迟加载当前行数据
    if (!currentRowData) {
      currentRowData = currentSheet.getRange(currentRow, 1, 1, currentSheet.getLastColumn()).getValues()[0];
      currentHeaders = currentSheet.getRange(1, 1, 1, currentSheet.getLastColumn()).getValues()[0];
    }
    
    // 检查数据是否一致
    const conflictRowData = sheet.getRange(conflict.row, 1, 1, sheet.getLastColumn()).getValues()[0];
    const conflictHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    if (!fastCompareRows(currentRowData, conflictRowData, currentHeaders, conflictHeaders)) {
      return conflict;
    }
  }
  
  return null;
}

/**
 * 快速检查单个其他表格
 */
function fastCheckSingleOtherSheet(sheet, value, columnName) {
  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1 || lastCol === 0) return null;
  
  // 一次性读取表头
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const columnIndex = headers.findIndex(header => header && header.toString() === columnName);
  
  if (columnIndex === -1) return null;
  
  // 一次性读取整列ID数据
  const idColumn = sheet.getRange(2, columnIndex + 1, lastRow - 1, 1).getValues().flat();
  
  // 快速查找第一个匹配的ID
  const foundIndex = idColumn.findIndex(id => 
    id && id.toString().trim() && id.toString() === value.toString()
  );
  
  if (foundIndex !== -1) {
    return {
      sheet: sheet.getName(),
      row: foundIndex + 2,
      column: columnIndex + 1
    };
  }
  
  return null;
}

/**
 * 快速行比较（简化版）
 */
function fastCompareRows(row1, row2, headers1, headers2) {
  // 快速长度检查
  if (!row1 || !row2 || !headers1 || !headers2) return false;
  
  // 创建简单的值映射进行比较
  const map1 = {};
  const map2 = {};
  
  headers1.forEach((header, i) => {
    if (header && header.toString().trim()) {
      map1[header.toString()] = row1[i] ? row1[i].toString().trim() : '';
    }
  });
  
  headers2.forEach((header, i) => {
    if (header && header.toString().trim()) {
      map2[header.toString()] = row2[i] ? row2[i].toString().trim() : '';
    }
  });
  
  // 比较所有相同的键
  const commonKeys = Object.keys(map1).filter(key => key in map2);
  
  if (commonKeys.length === 0) return false;
  
  return commonKeys.every(key => map1[key] === map2[key]);
}

/**
 * 检查ID冲突并标记
 */
function checkIdConflicts(editedCell) {
  console.log(`🚀 [ID冲突检查] 开始处理编辑单元格`);
  
  if (!editedCell) {
    console.log(`❌ [参数错误] editedCell 为空`);
    return;
  }
  
  const { sheet, range } = editedCell;
  const value = range.getValue();
  
  console.log(`📍 [编辑信息] 表: "${sheet.getName()}", 位置: 第${range.getRow()}行第${range.getColumn()}列, 值: "${value}"`);
  
  if (!value) {
    console.log(`🧹 [清除标记] 值为空，清除冲突标记`);
    range.setBackground(null);
    NoteManager.removeMarkFromCell(range, NOTE_CONSTANTS.TYPES.CONFLICT);
    return;
  }
  
  // 先移除历史标记
  console.log(`🧹 [清除历史标记] 移除之前的冲突标记`);
  NoteManager.removeMarkFromCell(range, NOTE_CONSTANTS.TYPES.CONFLICT);

  const headerValue = sheet.getRange(1, range.getColumn()).getValue();
  console.log(`📋 [列标题] "${headerValue}"`);
  
  try {
    const conflicts = checkSingleIdConflictImproved({
      value,
      sheet: sheet.getName(),
      row: range.getRow(),
      column: range.getColumn(),
      columnName: headerValue
    });
    
    if (conflicts.length > 0) {
      console.log(`🚨 [发现冲突] 找到 ${conflicts.length} 个冲突位置:`, conflicts);
      
      const conflictLocations = conflicts.map(loc => `${loc.sheet} 第${loc.row}行`).join('\n');
      const userNote = `在以下位置重复:\n${conflictLocations}`;
      
      console.log(`🎨 [设置标记] 设置红色背景和冲突注释`);
      range.setBackground(ID_CHECKER_CONFIG.COLORS.CONFLICT);
      range.setNote(NoteManager.addSystemNote(
        null,
        NOTE_CONSTANTS.TYPES.CONFLICT,  // 标记类型
        userNote
      ));
      
      console.log(`📢 [用户提示] 显示冲突警告`);
      SpreadsheetApp.getActiveSpreadsheet().toast('发现ID冲突，已用红色标记。', '警告', 3);
    } else {
      console.log(`✅ [无冲突] 未发现任何冲突`);
    }
  } catch (error) {
    console.error(`💥 [检查异常] ID冲突检查过程中出现错误:`, error);
    console.error(`🔍 [错误详情] 堆栈信息:`, error.stack);
    
    // 可选：显示错误提示给用户
    SpreadsheetApp.getActiveSpreadsheet().toast(`ID冲突检查出现错误: ${error.message}`, '错误', 5);
  }
  
  console.log(`🏁 [检查结束] ID冲突检查处理完成`);
}

/**
 * 手动验证并清理所有冲突标记
 * 这是一个兜底功能，用于清理可能过期的冲突标记
 */
function validateAndClearConflictMarks() {
  console.log(`🚀 [手动清理] 开始验证所有冲突标记`);
  
  const startTime = new Date().getTime();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let totalSheets = 0;
  let totalCells = 0;
  let conflictCells = 0;
  let clearedConflicts = 0;
  let validConflicts = 0;
  
  const results = {
    sheets: [],
    summary: {}
  };
  
  try {
    // 遍历所有表格
    const sheets = ss.getSheets();
    totalSheets = sheets.length;
    
    for (const sheet of sheets) {
      const sheetName = sheet.getName();
      console.log(`📑 [检查表格] ${sheetName}`);
      
      const sheetResult = {
        name: sheetName,
        totalCells: 0,
        conflictCells: 0,
        clearedConflicts: 0,
        validConflicts: 0,
        errors: []
      };
      
      try {
        const lastRow = sheet.getLastRow();
        const lastColumn = sheet.getLastColumn();
        
        if (lastRow <= 1 || lastColumn === 0) {
          console.log(`⏭️ [跳过表格] ${sheetName} - 无数据`);
          sheetResult.totalCells = 0;
          results.sheets.push(sheetResult);
          continue;
        }
        
        // 批量获取数据以提高性能
        const range = sheet.getRange(1, 1, lastRow, lastColumn);
        const backgrounds = range.getBackgrounds();
        const values = range.getValues();
        
        sheetResult.totalCells = lastRow * lastColumn;
        totalCells += sheetResult.totalCells;
        
        console.log(`📊 [表格范围] ${sheetName} - ${lastRow}行 × ${lastColumn}列 = ${sheetResult.totalCells}个单元格`);
        
        // 查找所有冲突标记的单元格
        for (let row = 0; row < backgrounds.length; row++) {
          for (let col = 0; col < backgrounds[row].length; col++) {
            
            // 检查是否是冲突标记
            if (backgrounds[row][col] === ID_CHECKER_CONFIG.COLORS.CONFLICT) {
              sheetResult.conflictCells++;
              conflictCells++;
              
              const actualRow = row + 1;
              const actualCol = col + 1;
              const cellValue = values[row][col];
              
              console.log(`🎯 [发现冲突标记] ${sheetName} 第${actualRow}行第${actualCol}列 - 值: "${cellValue}"`);
              
              try {
                // 重新检查这个单元格的冲突状态
                const cellRange = sheet.getRange(actualRow, actualCol);
                const stillHasConflict = recheckCellConflictStatus(cellRange, sheetName);
                
                if (!stillHasConflict) {
                  console.log(`✅ [清除过期冲突] ${sheetName} 第${actualRow}行第${actualCol}列`);
                  
                  // 清除冲突标记
                  cellRange.setBackground(null);
                  NoteManager.removeMarkFromCell(cellRange, NOTE_CONSTANTS.TYPES.CONFLICT);
                  
                  sheetResult.clearedConflicts++;
                  clearedConflicts++;
                } else {
                  console.log(`🔄 [保持有效冲突] ${sheetName} 第${actualRow}行第${actualCol}列`);
                  sheetResult.validConflicts++;
                  validConflicts++;
                }
              } catch (cellError) {
                console.error(`❌ [单元格处理错误] ${sheetName} 第${actualRow}行第${actualCol}列:`, cellError);
                sheetResult.errors.push(`第${actualRow}行第${actualCol}列: ${cellError.message}`);
              }
            }
          }
        }
        
        console.log(`✅ [表格完成] ${sheetName} - 冲突标记: ${sheetResult.conflictCells}, 清理: ${sheetResult.clearedConflicts}, 保留: ${sheetResult.validConflicts}`);
        
      } catch (sheetError) {
        console.error(`❌ [表格处理错误] ${sheetName}:`, sheetError);
        sheetResult.errors.push(`表格处理失败: ${sheetError.message}`);
      }
      
      results.sheets.push(sheetResult);
    }
    
    const endTime = new Date().getTime();
    const duration = endTime - startTime;
    
    // 汇总结果
    results.summary = {
      totalSheets,
      totalCells,
      conflictCells,
      clearedConflicts,
      validConflicts,
      duration,
      success: true
    };
    
    console.log(`🏁 [清理完成] 总耗时: ${duration}ms`);
    console.log(`📊 [最终统计] 表格: ${totalSheets}, 单元格: ${totalCells}, 冲突标记: ${conflictCells}, 清理: ${clearedConflicts}, 保留: ${validConflicts}`);
    
    // 显示结果给用户
    showCleanupResults(results);
    
    return results;
    
  } catch (error) {
    console.error(`💥 [清理异常] 手动清理过程中出现错误:`, error);
    
    const errorResult = {
      summary: {
        success: false,
        error: error.message,
        duration: new Date().getTime() - startTime
      }
    };
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `冲突标记清理失败: ${error.message}`, 
      '清理错误', 
      10
    );
    
    return errorResult;
  }
}

/**
 * 重新检查单个单元格的冲突状态
 * 这个函数被手动清理功能调用
 */
function recheckCellConflictStatus(cellRange, sheetName) {
  const value = cellRange.getValue();
  if (!value || !value.toString().trim()) {
    console.log(`📝 [重检查] 单元格值为空，无冲突`);
    return false;
  }
  
  // 检查是否是ID列
  const headerValue = cellRange.getSheet().getRange(1, cellRange.getColumn()).getValue();
  if (!headerValue || !headerValue.toString().endsWith(ID_CHECKER_CONFIG.ID_COLUMN_SUFFIX)) {
    console.log(`📝 [重检查] 不是ID列，无冲突`);
    return false;
  }
  
  console.log(`🔍 [重检查] 开始检查单元格冲突状态 - 值: "${value}", 列标题: "${headerValue}"`);
  
  // 使用现有的冲突检查逻辑
  const conflicts = checkSingleIdConflictImproved({
    value: value,
    sheet: sheetName,
    row: cellRange.getRow(),
    column: cellRange.getColumn(),
    columnName: headerValue
  });
  
  const hasConflict = conflicts.length > 0;
  console.log(`📝 [重检查结果] ${hasConflict ? '仍有冲突' : '无冲突'} - 冲突数量: ${conflicts.length}`);
  
  return hasConflict;
}

/**
 * 显示清理结果给用户
 */
function showCleanupResults(results) {
  const { summary, sheets } = results;
  
  if (!summary.success) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `清理失败: ${summary.error}`,
      '清理错误',
      10
    );
    return;
  }
  
  // 构建详细报告
  let message = `🧹 冲突标记清理完成！\n\n`;
  message += `📊 总体统计:\n`;
  message += `• 检查表格: ${summary.totalSheets} 个\n`;
  message += `• 检查单元格: ${summary.totalCells.toLocaleString()} 个\n`;
  message += `• 发现冲突标记: ${summary.conflictCells} 个\n`;
  message += `• 清理过期冲突: ${summary.clearedConflicts} 个\n`;
  message += `• 保留有效冲突: ${summary.validConflicts} 个\n`;
  message += `• 处理耗时: ${summary.duration}ms\n\n`;
  
  // 添加表格详情（只显示有操作的表格）
  const activeSheets = sheets.filter(sheet => sheet.conflictCells > 0);
  if (activeSheets.length > 0) {
    message += `📋 表格详情:\n`;
    activeSheets.forEach(sheet => {
      message += `• ${sheet.name}: 冲突${sheet.conflictCells}, 清理${sheet.clearedConflicts}, 保留${sheet.validConflicts}\n`;
    });
  }
  
  // 显示结果
  if (summary.clearedConflicts > 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `成功清理了 ${summary.clearedConflicts} 个过期的冲突标记！\n保留 ${summary.validConflicts} 个有效冲突。`,
      '清理完成',
      8
    );
  } else if (summary.conflictCells === 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      '未发现任何冲突标记，表格状态良好！',
      '清理完成',
      5
    );
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `所有 ${summary.validConflicts} 个冲突标记都是有效的，无需清理。`,
      '清理完成',
      5
    );
  }
  
  console.log(`📋 [清理报告]\n${message}`);
}
