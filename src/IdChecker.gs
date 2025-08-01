/**
 * 改进的ID冲突检查逻辑
 * - 当前表格内：相同ID直接认为冲突
 * - 跨表格：只有当ID相同但整行数据不同时才认为是真正的冲突
 */
function checkSingleIdConflictImproved({ value, sheet: sheetName, row, column, columnName }) {
  console.log(`🔍 [ID冲突检查] 开始检查 - 值: "${value}", 表: "${sheetName}", 行: ${row}, 列: ${column}, 列名: "${columnName}"`);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentSheet = ss.getSheetByName(sheetName);
  
  // 获取当前行的完整数据（只有跨表格检查时才需要）
  let currentRowData = null;
  
  let totalCheckedSheets = 0;
  let totalCheckedRows = 0;
  let foundSameIds = 0;
  let identicalRowsSkipped = 0;
  
  // 遍历所有表格查找相同ID
  for (const sheet of ss.getSheets()) {
    totalCheckedSheets++;
    const isCurrentSheet = sheet.getName() === sheetName;
    console.log(`📑 [检查表格] ${sheet.getName()}${isCurrentSheet ? ' (当前表格)' : ' (其他表格)'}`);
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const columnIndex = headers.findIndex(header => header.toString() === columnName);
    
    if (columnIndex === -1) {
      console.log(`❌ [跳过表格] ${sheet.getName()} - 未找到列 "${columnName}"`);
      continue;
    }
    
    if (sheet.getLastRow() <= 1) {
      console.log(`❌ [跳过表格] ${sheet.getName()} - 无数据行`);
      continue;
    }
    
    console.log(`✅ [检查表格] ${sheet.getName()} - 找到列 "${columnName}" 在第 ${columnIndex + 1} 列`);
    
    // 获取列数据并检查冲突
    const data = sheet.getRange(2, columnIndex + 1, sheet.getLastRow() - 1, 1).getValues();
    console.log(`📊 [数据范围] ${sheet.getName()} - 从第2行到第${sheet.getLastRow()}行，共 ${data.length} 行数据`);
    
    for (let i = 0; i < data.length; i++) {
      totalCheckedRows++;
      const [id] = data[i];
      
      if (!id?.toString().trim()) {
        console.log(`⏭️  [跳过空值] ${sheet.getName()} 第${i + 2}行 - ID为空`);
        continue;
      }
      
      const isCurrentCell = isCurrentSheet && 
                          i + 2 === row && 
                          columnIndex + 1 === column;
      
      if (isCurrentCell) {
        console.log(`🎯 [当前单元格] ${sheet.getName()} 第${i + 2}行 - 跳过自身`);
        continue;
      }
      
      if (id.toString() === value.toString()) {
        foundSameIds++;
        console.log(`🎯 [找到相同ID] ${sheet.getName()} 第${i + 2}行 - ID: "${id}"`);
        
        if (isCurrentSheet) {
          // 🔥 当前表格内：相同ID直接认为冲突
          console.log(`🚨 [当前表格冲突] ${sheet.getName()} 第${i + 2}行 - 同一表格内不允许重复ID`);
          console.log(`📊 [检查统计] 总共检查了 ${totalCheckedSheets} 个表格, ${totalCheckedRows} 行数据, 发现 ${foundSameIds} 个相同ID`);
          
          return [{
            sheet: sheet.getName(),
            row: i + 2,
            column: columnIndex + 1
          }];
        } else {
          // 🔄 跨表格：检查整行数据是否一致
          console.log(`🔄 [跨表格检查] 需要比较整行数据一致性`);
          
          // 延迟获取当前行数据，只有在需要时才获取
          if (currentRowData === null) {
            currentRowData = currentSheet.getRange(row, 1, 1, currentSheet.getLastColumn()).getValues()[0];
            console.log(`📋 [当前行数据] ${JSON.stringify(currentRowData)}`);
          }
          
          const conflictRowData = sheet.getRange(i + 2, 1, 1, sheet.getLastColumn()).getValues()[0];
          console.log(`📋 [冲突行数据] ${JSON.stringify(conflictRowData)}`);
          
          const currentHeaders = currentSheet.getRange(1, 1, 1, currentSheet.getLastColumn()).getValues()[0];
          const areIdentical = areRowsIdentical(currentRowData, conflictRowData, headers, currentHeaders);
          
          console.log(`🔄 [跨表格数据比较] 两行数据是否一致: ${areIdentical}`);
          
          if (!areIdentical) {
            console.log(`🚨 [跨表格冲突] ${sheet.getName()} 第${i + 2}行 - ID相同但数据不一致`);
            console.log(`📊 [检查统计] 总共检查了 ${totalCheckedSheets} 个表格, ${totalCheckedRows} 行数据, 发现 ${foundSameIds} 个相同ID, 跳过 ${identicalRowsSkipped} 个一致行`);
            
            return [{
              sheet: sheet.getName(),
              row: i + 2,
              column: columnIndex + 1
            }];
          } else {
            identicalRowsSkipped++;
            console.log(`✅ [跳过一致行] ${sheet.getName()} 第${i + 2}行 - 跨表格ID相同但整行数据一致，不视为冲突`);
          }
        }
      }
    }
  }
  
  console.log(`✅ [检查完成] 无冲突 - 统计: 检查了 ${totalCheckedSheets} 个表格, ${totalCheckedRows} 行数据, 发现 ${foundSameIds} 个相同ID, 跳过 ${identicalRowsSkipped} 个一致行`);
  return [];
}

/**
 * 比较两行数据是否一致（考虑表头对应关系）
 */
function areRowsIdentical(row1Data, row2Data, headers1, headers2) {
  console.log(`🔄 [开始行比较] 表头1: ${JSON.stringify(headers1)}`);
  console.log(`🔄 [开始行比较] 表头2: ${JSON.stringify(headers2)}`);
  
  // 创建表头映射
  const headerMap = {};
  let mappedColumns = 0;
  
  headers1.forEach((header, index) => {
    const mappedIndex = headers2.findIndex(h => h.toString() === header.toString());
    if (mappedIndex !== -1) {
      headerMap[index] = mappedIndex;
      mappedColumns++;
      console.log(`🗺️  [表头映射] "${header}" -> 列 ${index + 1} 映射到列 ${mappedIndex + 1}`);
    } else {
      console.log(`❌ [表头未映射] "${header}" 在目标表中未找到`);
    }
  });
  
  console.log(`📊 [映射统计] 成功映射 ${mappedColumns} 个列`);
  
  if (mappedColumns === 0) {
    console.log(`⚠️  [映射警告] 没有找到任何可比较的列，默认认为不一致`);
    return false;
  }
  
  // 比较对应列的数据
  let comparedColumns = 0;
  let differentColumns = 0;
  
  for (const [index1, index2] of Object.entries(headerMap)) {
    comparedColumns++;
    const value1 = normalizeValue(row1Data[parseInt(index1)]);
    const value2 = normalizeValue(row2Data[parseInt(index2)]);
    
    console.log(`🔍 [列比较] 列 ${parseInt(index1) + 1}("${headers1[parseInt(index1)]}") - 值1: "${value1}" vs 值2: "${value2}"`);
    
    if (value1 !== value2) {
      differentColumns++;
      console.log(`❌ [数据不一致] 列 ${parseInt(index1) + 1} - "${value1}" ≠ "${value2}"`);
      return false;
    } else {
      console.log(`✅ [数据一致] 列 ${parseInt(index1) + 1} - "${value1}" = "${value2}"`);
    }
  }
  
  console.log(`✅ [行比较完成] 比较了 ${comparedColumns} 列，${differentColumns} 列不同 - 结果: 完全一致`);
  return true;
}

/**
 * 标准化值用于比较
 */
function normalizeValue(value) {
  if (value === null || value === undefined) return '';
  const normalized = value.toString().trim();
  // console.log(`🔧 [标准化] "${value}" -> "${normalized}"`);  // 这个日志太多，可按需开启
  return normalized;
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
