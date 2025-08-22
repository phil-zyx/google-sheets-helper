/**
 * Google Sheets ID冲突检查器 - 优化版本
 * 提供高效的ID冲突检测、验证和清理功能
 * 
 * 主要功能：
 * 1. 单次ID冲突检查
 * 2. 批量冲突标记验证和清理
 * 3. 智能冲突检测（ID + 数据一致性验证）
 * 
 * 性能优化特性：
 * - 批量数据读取
 * - 表头缓存机制
 * - 避免重复检查
 * - 智能行数据比较
 */

// ============================================================================
// 批量检查配置常量
// ============================================================================

const BATCH_CHECKER_CONFIG = {
  // 智能选择阈值：表格数量超过此值时使用批量方式
  SHEET_COUNT_THRESHOLD: 3,
  
  // 批量验证阈值：冲突单元格数量超过此值时使用批量验证
  CONFLICT_CELL_THRESHOLD: 5,
  
  // 超级批量验证阈值：冲突单元格数量超过此值时使用超级批量验证
  SUPER_BATCH_THRESHOLD: 20,
  
  // 是否启用智能选择（默认true）
  ENABLE_SMART_SELECTION: true,
  
  // 是否启用批量验证（默认true）
  ENABLE_BATCH_VALIDATION: true,
  
  // 是否启用超级批量验证（默认true）
  ENABLE_SUPER_BATCH_VALIDATION: true,
  
  // 内存管理配置
  MEMORY_MANAGEMENT: {
    // 是否启用内存监控（默认true）
    ENABLED: true,
    
    // 单次读取的最大行数（避免一次性读取过多数据）
    MAX_ROWS_PER_BATCH: 1000,
    
    // 单次读取的最大表格数（避免同时处理过多表格）
    MAX_SHEETS_PER_BATCH: 10,
    
    // 内存使用警告阈值（MB）
    MEMORY_WARNING_THRESHOLD: 50,
    
    // 是否启用分批处理（默认true）
    ENABLE_BATCH_PROCESSING: true,
    
    // 分批处理时的延迟时间（ms）
    BATCH_DELAY: 100
  },
  
  // 性能测试配置
  PERFORMANCE_TEST: {
    ENABLED: true,
    LOG_DETAILS: true
  }
};

// ============================================================================
// 核心冲突检查函数
// ============================================================================

/**
 * ID冲突检查
 */
function handleIdConflictCheck(context) {
  const { headerRow, range } = context;
  const idColumns = findIdColumns(headerRow);
  
  if (idColumns.length === 0) return;
  
  const editedColumns = getEditedColumns(range);
  const relevantIdColumns = idColumns.filter(idCol => editedColumns.includes(idCol));
  
  if (relevantIdColumns.length === 0) return;
  
  checkRelevantIdColumns(context, relevantIdColumns);
}

/**
 * 优化的ID列查找
 */
function findIdColumns(headerRow) {
  const idColumns = [];
  const suffix = ID_CHECKER_CONFIG.ID_COLUMN_SUFFIX;
  
  for (let col = 0; col < headerRow.length; col++) {
    const header = headerRow[col];
    if (header && header.toString().endsWith(suffix)) {
      idColumns.push(col + 1);
    }
  }
  
  return idColumns;
}

/**
 * 优化的相关ID列检查
 */
function checkRelevantIdColumns(context, relevantIdColumns) {
  const { sheet, range } = context;
  
  // 批量处理所有ID列，而不是逐个处理
  for (const idCol of relevantIdColumns) {
    const idRange = sheet.getRange(range.getRow(), idCol, range.getNumRows(), 1);
    checkIdConflicts({
      sheet: sheet,
      range: idRange
    });
  }
}

/**
 * 单次ID冲突检查 - 用于实时编辑检测
 * @param {Object} params - 检查参数
 * @param {string} params.value - 要检查的ID值
 * @param {string} params.sheet - 表格名称
 * @param {number} params.row - 行号
 * @param {number} params.column - 列号
 * @param {string} params.columnName - 列标题
 * @param {boolean} params.useSingleRowOptimization - 是否使用单行优化检查（默认true）
 * @returns {Array} 冲突位置数组
 */
function checkSingleIdConflict({ value, sheet: sheetName, row, column, columnName, useSingleRowOptimization = true }) {
  const startTime = Date.now();
  console.log(`🔍 [单次检查] ${sheetName} 第${row}行 - "${value}"`);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentSheet = ss.getSheetByName(sheetName);
  
  // 1. 快速检查当前表格
  const currentConflict = checkCurrentSheetConflict(currentSheet, value, row, columnName);
  if (currentConflict) {
    console.log(`🚨 [当前表格冲突] 耗时: ${Date.now() - startTime}ms`);
    return [currentConflict];
  }
  
  // 2. 检查其他表格（使用单行优化检查）
  let crossConflict;
  if (useSingleRowOptimization && BATCH_CHECKER_CONFIG.ENABLE_SMART_SELECTION) {
    crossConflict = checkOtherSheetsConflictSingleRowSmart(ss, currentSheet, sheetName, value, row, columnName);
  } else {
    crossConflict = checkOtherSheetsConflictSmart(ss, currentSheet, sheetName, value, row, columnName);
  }
  
  console.log(`✅ [检查完成] 耗时: ${Date.now() - startTime}ms`);
  return crossConflict ? [crossConflict] : [];
}

/**
 * 检查当前表格内的ID冲突
 * @param {Sheet} sheet - 当前表格对象
 * @param {string} value - ID值
 * @param {number} currentRow - 当前行号
 * @param {string} columnName - 列标题
 * @returns {Object|null} 冲突信息或null
 */
function checkCurrentSheetConflict(sheet, value, currentRow, columnName) {
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
 * 检查其他表格的ID冲突
 * @param {Spreadsheet} ss - 电子表格对象
 * @param {Sheet} currentSheet - 当前表格
 * @param {string} currentSheetName - 当前表格名称
 * @param {string} value - ID值
 * @param {number} currentRow - 当前行号
 * @param {string} columnName - 列标题
 * @returns {Object|null} 冲突信息或null
 */
function checkOtherSheetsConflict(ss, currentSheet, currentSheetName, value, currentRow, columnName) {
  const sheets = ss.getSheets();
  let currentRowData = null;
  let currentHeaders = null;
  
  for (const sheet of sheets) {
    if (sheet.getName() === currentSheetName) continue;
    
    const conflict = findConflictInSheet(sheet, value, columnName);
    if (!conflict) continue;
    
    // 延迟加载当前行数据
    if (!currentRowData) {
      currentRowData = currentSheet.getRange(currentRow, 1, 1, currentSheet.getLastColumn()).getValues()[0];
      currentHeaders = currentSheet.getRange(1, 1, 1, currentSheet.getLastColumn()).getValues()[0];
    }
    
    // 检查数据是否一致
    const conflictRowData = sheet.getRange(conflict.row, 1, 1, sheet.getLastColumn()).getValues()[0];
    const conflictHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    if (!compareRowsData(currentRowData, conflictRowData, currentHeaders, conflictHeaders)) {
      return conflict;
    }
  }
  
  return null;
}

/**
 * 批量检查其他表格的ID冲突 - 核心优化函数
 * 使用批量预加载和缓存机制大幅提升性能
 * 
 * 优化策略：
 * 1. 批量预加载所有表格的ID列数据
 * 2. 快速查找潜在冲突位置
 * 3. 一次性读取所有冲突行数据
 * 4. 批量验证数据一致性
 * 
 * @param {Spreadsheet} ss - 电子表格对象
 * @param {Sheet} currentSheet - 当前表格
 * @param {string} currentSheetName - 当前表格名称
 * @param {string} value - ID值
 * @param {number} currentRow - 当前行号
 * @param {string} columnName - 列标题
 * @returns {Object|null} 冲突信息或null
 */
function checkOtherSheetsConflictBatch(ss, currentSheet, currentSheetName, value, currentRow, columnName) {
  console.log(`🚀 [批量检查] 开始批量检查其他表格的ID冲突: "${value}"`);
  const startTime = Date.now();
  
  // 第一步：批量预加载所有表格的ID列数据
  const sheetIdDataCache = buildSheetIdDataCache(ss, columnName);
  if (sheetIdDataCache.size === 0) {
    console.log(`⚠️ [批量检查] 未找到任何包含列 "${columnName}" 的表格`);
    return null;
  }
  
  // 第二步：快速查找所有包含该ID的表格位置
  const potentialConflicts = findPotentialConflicts(value, sheetIdDataCache, currentSheetName);
  if (potentialConflicts.length === 0) {
    console.log(`✅ [批量检查] 未发现任何潜在冲突，耗时: ${Date.now() - startTime}ms`);
    return null;
  }
  
  console.log(`🎯 [批量检查] 发现 ${potentialConflicts.length} 个潜在冲突位置，开始批量验证...`);
  
  // 第三步：一次性读取当前行数据和表头
  const currentRowData = currentSheet.getRange(currentRow, 1, 1, currentSheet.getLastColumn()).getValues()[0];
  const currentHeaders = currentSheet.getRange(1, 1, 1, currentSheet.getLastColumn()).getValues()[0];
  
  // 第四步：批量读取所有潜在冲突行的数据
  const conflictRowsData = batchReadConflictRowsData(ss, potentialConflicts);
  
  // 第五步：批量验证所有冲突行
  const conflict = validateAllConflicts(potentialConflicts, conflictRowsData, sheetIdDataCache, currentRowData, currentHeaders);
  
  const totalTime = Date.now() - startTime;
  if (conflict) {
    console.log(`🚨 [批量检查] 发现真实冲突: ${conflict.sheet} 第${conflict.row}行，总耗时: ${totalTime}ms`);
  } else {
    console.log(`✅ [批量检查] 所有潜在冲突都通过数据一致性验证，无真实冲突，总耗时: ${totalTime}ms`);
  }
  
  return conflict;
}

/**
 * 单行编辑跨页签检查 - 超级优化版本
 * 专门针对单行编辑后的ID冲突检查，实现真正的"一次性读取所有跨页签数据"
 * 
 * 超级优化特性：
 * 1. 一次性读取当前行数据和表头
 * 2. 一次性预加载所有表格的ID列数据
 * 3. 一次性读取所有跨页签的冲突行数据
 * 4. 智能内存管理和分批处理
 * 
 * @param {Spreadsheet} ss - 电子表格对象
 * @param {Sheet} currentSheet - 当前表格
 * @param {string} currentSheetName - 当前表格名称
 * @param {string} value - ID值
 * @param {number} currentRow - 当前行号
 * @param {string} columnName - 列标题
 * @returns {Object|null} 冲突信息或null
 */
function checkOtherSheetsConflictSingleRowOptimized(ss, currentSheet, currentSheetName, value, currentRow, columnName) {
  console.log(`🚀 [单行优化检查] 开始单行编辑跨页签检查: "${value}"`);
  const startTime = Date.now();
  
  // 第一步：一次性读取当前行数据和表头
  const currentRowData = currentSheet.getRange(currentRow, 1, 1, currentSheet.getLastColumn()).getValues()[0];
  const currentHeaders = currentSheet.getRange(1, 1, 1, currentSheet.getLastColumn()).getValues()[0];
  
  // 第二步：一次性预加载所有表格的ID列数据
  const sheetIdDataCache = buildSheetIdDataCache(ss, columnName);
  if (sheetIdDataCache.size === 0) {
    console.log(`⚠️ [单行优化检查] 未找到任何包含列 "${columnName}" 的表格`);
    return null;
  }
  
  // 第三步：快速查找所有包含该ID的表格位置
  const potentialConflicts = findPotentialConflicts(value, sheetIdDataCache, currentSheetName);
  if (potentialConflicts.length === 0) {
    console.log(`✅ [单行优化检查] 未发现任何潜在冲突，总耗时: ${Date.now() - startTime}ms`);
    return null;
  }
  
  console.log(`🎯 [单行优化检查] 发现 ${potentialConflicts.length} 个潜在冲突位置，开始超级批量验证...`);
  
  // 第四步：一次性读取所有跨页签的冲突行数据（带内存管理）
  const allConflictRowsData = readAllCrossSheetConflictRows(ss, potentialConflicts, sheetIdDataCache);
  
  // 第五步：批量验证所有冲突行
  const conflict = validateAllConflictsOptimized(potentialConflicts, allConflictRowsData, currentRowData, currentHeaders);
  
  const totalTime = Date.now() - startTime;
  if (conflict) {
    console.log(`🚨 [单行优化检查] 发现真实冲突: ${conflict.sheet} 第${conflict.row}行，总耗时: ${totalTime}ms`);
  } else {
    console.log(`✅ [单行优化检查] 所有潜在冲突都通过数据一致性验证，无真实冲突，总耗时: ${totalTime}ms`);
  }
  
  return conflict;
}

// ============================================================================
// 批量数据读取和验证 - 核心辅助函数
// ============================================================================

/**
 * 批量读取所有潜在冲突行的数据
 * 一次性读取所有需要验证的行，避免逐个调用API
 * 
 * @param {Spreadsheet} ss - 电子表格对象
 * @param {Array} potentialConflicts - 潜在冲突列表
 * @returns {Map} 冲突行数据映射 (key: "表格名_行号", value: 行数据)
 */
function batchReadConflictRowsData(ss, potentialConflicts) {
  console.log(`⚡ [批量读取] 开始批量读取 ${potentialConflicts.length} 行冲突数据...`);
  const startTime = Date.now();
  
  const conflictRowsData = new Map();
  
  // 按表格分组，减少API调用次数
  const sheetGroups = groupConflictsBySheet(potentialConflicts);
  console.log(`📊 [批量读取] 需要读取 ${sheetGroups.size} 个表格的数据`);
  
  // 逐个表格批量读取
  for (const [sheetName, conflicts] of sheetGroups) {
    try {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) continue;
      
      const lastCol = sheet.getLastColumn();
      if (lastCol === 0) continue;
      
      // 检查是否需要分批读取（内存管理）
      if (shouldUseBatchProcessing(potentialConflicts)) {
        console.log(`⚠️ [内存管理] ${sheetName} 需要分批读取，使用分批处理模式`);
        const batchData = readSheetRowsInBatches(sheet, conflicts, lastCol);
        batchData.forEach((value, key) => conflictRowsData.set(key, value));
        continue;
      }
      
      // 一次性读取该表格中所有相关行
      const rowsData = readSheetRowsOptimized(sheet, conflicts, lastCol);
      
      // 将读取的数据映射到对应的冲突位置
      mapRowsDataToConflicts(conflicts, rowsData, sheetName, conflictRowsData);
      
    } catch (error) {
      console.warn(`⚠️ [批量读取警告] 无法读取表格 ${sheetName} 的数据: ${error.message}`);
    }
  }
  
  const duration = Date.now() - startTime;
  console.log(`⚡ [批量读取完成] 成功读取 ${conflictRowsData.size}/${potentialConflicts.length} 行冲突数据，耗时: ${duration}ms`);
  
  return conflictRowsData;
}

/**
 * 一次性读取所有跨页签的冲突行数据
 * 这是单行编辑检查的核心优化：真正实现"一次性读取所有跨页签数据"
 * 
 * @param {Spreadsheet} ss - 电子表格对象
 * @param {Array} potentialConflicts - 潜在冲突列表
 * @param {Map} sheetIdDataCache - 表格ID列数据缓存
 * @returns {Map} 所有冲突行数据映射 (key: "表格名_行号", value: {data: 行数据, headers: 表头})
 */
function readAllCrossSheetConflictRows(ss, potentialConflicts, sheetIdDataCache) {
  console.log(`⚡ [跨页签读取] 开始一次性读取所有跨页签冲突行数据...`);
  const startTime = Date.now();
  
  // 内存管理：检查是否需要分批处理
  if (BATCH_CHECKER_CONFIG.MEMORY_MANAGEMENT.ENABLED && 
      shouldUseBatchProcessing(potentialConflicts)) {
    console.log(`⚠️ [内存管理] 检测到大量数据，启用分批处理模式`);
    return readAllCrossSheetConflictRowsBatched(ss, potentialConflicts, sheetIdDataCache);
  }
  
  const allConflictRowsData = new Map();
  
  // 按表格分组，准备批量读取
  const sheetGroups = groupConflictsBySheet(potentialConflicts);
  console.log(`📊 [跨页签读取] 需要读取 ${sheetGroups.size} 个跨页签的数据`);
  
  // 逐个表格进行超级批量读取
  for (const [sheetName, conflicts] of sheetGroups) {
    try {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) continue;
      
      const lastCol = sheet.getLastColumn();
      if (lastCol === 0) continue;
      
      // 检查单次读取的行数（内存管理）
      const rowCount = calculateRowCount(conflicts);
      if (BATCH_CHECKER_CONFIG.MEMORY_MANAGEMENT.ENABLED && 
          rowCount > BATCH_CHECKER_CONFIG.MEMORY_MANAGEMENT.MAX_ROWS_PER_BATCH) {
        console.log(`⚠️ [内存管理] ${sheetName} 需要读取 ${rowCount} 行，超过阈值，使用分批读取`);
        const batchData = readSheetRowsInBatches(sheet, conflicts, lastCol);
        batchData.forEach((value, key) => allConflictRowsData.set(key, value));
        continue;
      }
      
      // 一次性读取该表格中所有相关行 + 表头
      const rowsData = readSheetRowsOptimized(sheet, conflicts, lastCol);
      const headers = getSheetHeaders(sheet, sheetIdDataCache, sheetName, lastCol);
      
      // 将读取的数据和表头一起映射到对应的冲突位置
      mapRowsDataWithHeaders(conflicts, rowsData, headers, sheetName, allConflictRowsData);
      
    } catch (error) {
      console.warn(`⚠️ [跨页签读取警告] 无法读取表格 ${sheetName} 的数据: ${error.message}`);
    }
  }
  
  const duration = Date.now() - startTime;
  console.log(`⚡ [跨页签读取完成] 成功读取 ${allConflictRowsData.size}/${potentialConflicts.length} 行跨页签数据，耗时: ${duration}ms`);
  
  return allConflictRowsData;
}

// ============================================================================
// 数据验证和比较 - 核心逻辑函数
// ============================================================================

/**
 * 验证所有冲突行（批量方式）
 * @param {Array} potentialConflicts - 潜在冲突列表
 * @param {Map} conflictRowsData - 冲突行数据
 * @param {Map} sheetIdDataCache - 表格ID列数据缓存
 * @param {Array} currentRowData - 当前行数据
 * @param {Array} currentHeaders - 当前行表头
 * @returns {Object|null} 发现的冲突或null
 */
function validateAllConflicts(potentialConflicts, conflictRowsData, sheetIdDataCache, currentRowData, currentHeaders) {
  for (let i = 0; i < potentialConflicts.length; i++) {
    const conflict = potentialConflicts[i];
    const conflictRowData = conflictRowsData.get(conflict.sheet + '_' + conflict.row);
    const conflictHeaders = sheetIdDataCache.get(conflict.sheet).headers;
    
    if (conflictRowData && !compareRowsData(currentRowData, conflictRowData, currentHeaders, conflictHeaders)) {
      return conflict;
    }
  }
  return null;
}

/**
 * 验证所有冲突行（单行优化方式）
 * @param {Array} potentialConflicts - 潜在冲突列表
 * @param {Map} allConflictRowsData - 所有冲突行数据
 * @param {Array} currentRowData - 当前行数据
 * @param {Array} currentHeaders - 当前行表头
 * @returns {Object|null} 发现的冲突或null
 */
function validateAllConflictsOptimized(potentialConflicts, allConflictRowsData, currentRowData, currentHeaders) {
  for (let i = 0; i < potentialConflicts.length; i++) {
    const conflict = potentialConflicts[i];
    const conflictData = allConflictRowsData.get(conflict.sheet + '_' + conflict.row);
    
    if (conflictData && !compareRowsData(currentRowData, conflictData.data, currentHeaders, conflictData.headers)) {
      return conflict;
    }
  }
  return null;
}

// ============================================================================
// 数据读取辅助函数 - 优化和简化
// ============================================================================

/**
 * 检查是否需要使用分批处理
 * @param {Array} potentialConflicts - 潜在冲突列表
 * @returns {boolean} 是否需要分批处理
 */
function shouldUseBatchProcessing(potentialConflicts) {
  if (!BATCH_CHECKER_CONFIG.MEMORY_MANAGEMENT.ENABLED) {
    return false;
  }
  
  // 检查冲突数量
  if (potentialConflicts.length > 100) {
    console.log(`⚠️ [内存管理] 潜在冲突数量过多: ${potentialConflicts.length} > 100，建议分批处理`);
    return true;
  }
  
  // 检查表格数量
  const uniqueSheets = new Set(potentialConflicts.map(c => c.sheet));
  if (uniqueSheets.size > BATCH_CHECKER_CONFIG.MEMORY_MANAGEMENT.MAX_SHEETS_PER_BATCH) {
    console.log(`⚠️ [内存管理] 涉及表格数量过多: ${uniqueSheets.size} > ${BATCH_CHECKER_CONFIG.MEMORY_MANAGEMENT.MAX_SHEETS_PER_BATCH}，建议分批处理`);
    return true;
  }
  
  // 检查内存使用情况（如果可用）
  if (typeof Memory !== 'undefined') {
    try {
      const memoryUsage = Memory.getMemoryUsage();
      if (memoryUsage && memoryUsage.used > BATCH_CHECKER_CONFIG.MEMORY_MANAGEMENT.MEMORY_WARNING_THRESHOLD * 1024 * 1024) {
        console.log(`⚠️ [内存管理] 内存使用过高: ${(memoryUsage.used / 1024 / 1024).toFixed(2)}MB > ${BATCH_CHECKER_CONFIG.MEMORY_MANAGEMENT.MEMORY_WARNING_THRESHOLD}MB，建议分批处理`);
        return true;
      }
    } catch (e) {
      console.warn(`⚠️ [内存管理] 无法获取内存使用信息: ${e.message}`);
    }
  }
  
  return false;
}

/**
 * 分批读取表格行数据
 * @param {Sheet} sheet - 表格对象
 * @param {Array} conflicts - 冲突列表
 * @param {number} lastCol - 最后一列
 * @returns {Map} 分批读取的数据
 */
function readSheetRowsInBatches(sheet, conflicts, lastCol) {
  console.log(`📦 [分批读取] 开始分批读取表格 ${sheet.getName()} 的数据`);
  const batchData = new Map();
  
  // 按行号排序
  const sortedConflicts = conflicts.sort((a, b) => a.row - b.row);
  
  // 分批处理
  const batchSize = BATCH_CHECKER_CONFIG.MEMORY_MANAGEMENT.MAX_ROWS_PER_BATCH;
  let processedCount = 0;
  
  for (let i = 0; i < sortedConflicts.length; i += batchSize) {
    const batch = sortedConflicts.slice(i, i + batchSize);
    const minRow = batch[0].row;
    const maxRow = batch[batch.length - 1].row;
    const rowCount = maxRow - minRow + 1;
    
    console.log(`📦 [分批读取] 批次 ${Math.floor(i / batchSize) + 1}: 读取 ${minRow}-${maxRow} 行，共 ${rowCount} 行`);
    
    try {
      // 读取这一批的数据
      const range = sheet.getRange(minRow, 1, rowCount, lastCol);
      const rowsData = range.getValues();
      
      // 获取表头
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      
      // 映射数据
      for (const conflict of batch) {
        const rowIndex = conflict.row - minRow;
        if (rowIndex >= 0 && rowIndex < rowsData.length) {
          const key = `${sheet.getName()}_${conflict.row}`;
          batchData.set(key, {
            data: rowsData[rowIndex],
            headers: headers
          });
          
          // 更新conflict对象
          conflict.headers = headers;
        }
      }
      
      processedCount += batch.length;
      console.log(`✅ [分批读取] 批次 ${Math.floor(i / batchSize) + 1} 完成，已处理 ${processedCount}/${conflicts.length} 行`);
      
      // 添加延迟，避免API限制
      if (BATCH_CHECKER_CONFIG.MEMORY_MANAGEMENT.ENABLE_BATCH_PROCESSING && 
          i + batchSize < sortedConflicts.length) {
        Utilities.sleep(BATCH_CHECKER_CONFIG.MEMORY_MANAGEMENT.BATCH_DELAY);
      }
      
    } catch (error) {
      console.error(`❌ [分批读取错误] 批次 ${Math.floor(i / batchSize) + 1} 失败: ${error.message}`);
    }
  }
  
  console.log(`📦 [分批读取完成] 表格 ${sheet.getName()} 共处理 ${processedCount} 行数据`);
  return batchData;
}

/**
 * 分批处理版本的跨页签冲突行数据读取
 * @param {Spreadsheet} ss - 电子表格对象
 * @param {Array} potentialConflicts - 潜在冲突列表
 * @param {Map} sheetIdDataCache - 表格ID列数据缓存
 * @returns {Map} 所有冲突行数据映射
 */
function readAllCrossSheetConflictRowsBatched(ss, potentialConflicts, sheetIdDataCache) {
  console.log(`📦 [分批处理] 开始分批处理跨页签冲突行数据读取`);
  const startTime = Date.now();
  
  const allConflictRowsData = new Map();
  
  // 按表格分组
  const sheetGroups = groupConflictsBySheet(potentialConflicts);
  const totalSheets = sheetGroups.size;
  console.log(`📊 [分批处理] 需要处理 ${totalSheets} 个表格的数据`);
  
  // 分批处理表格
  const sheetsArray = Array.from(sheetGroups.entries());
  const batchSize = BATCH_CHECKER_CONFIG.MEMORY_MANAGEMENT.MAX_SHEETS_PER_BATCH;
  
  for (let i = 0; i < sheetsArray.length; i += batchSize) {
    const sheetBatch = sheetsArray.slice(i, i + batchSize);
    const batchNumber = Math.floor(i / batchSize) + 1;
    const totalBatches = Math.ceil(sheetsArray.length / batchSize);
    
    console.log(`📦 [分批处理] 处理表格批次 ${batchNumber}/${totalBatches}: ${sheetBatch.length} 个表格`);
    
    // 处理这一批表格
    for (const [sheetName, conflicts] of sheetBatch) {
      try {
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) continue;
        
        const lastCol = sheet.getLastColumn();
        if (lastCol === 0) continue;
        
        // 使用分批读取
        const batchData = readSheetRowsInBatches(sheet, conflicts, lastCol);
        batchData.forEach((value, key) => allConflictRowsData.set(key, value));
        
      } catch (error) {
        console.error(`❌ [分批处理错误] 表格 ${sheetName} 处理失败: ${error.message}`);
      }
    }
    
    console.log(`✅ [分批处理] 表格批次 ${batchNumber}/${totalBatches} 完成`);
    
    // 添加延迟，避免API限制
    if (BATCH_CHECKER_CONFIG.MEMORY_MANAGEMENT.ENABLE_BATCH_PROCESSING && 
        i + batchSize < sheetsArray.length) {
      console.log(`⏳ [分批处理] 等待 ${BATCH_CHECKER_CONFIG.MEMORY_MANAGEMENT.BATCH_DELAY}ms 后继续下一批...`);
      Utilities.sleep(BATCH_CHECKER_CONFIG.MEMORY_MANAGEMENT.BATCH_DELAY);
    }
  }
  
  const duration = Date.now() - startTime;
  console.log(`📦 [分批处理完成] 成功处理 ${allConflictRowsData.size}/${potentialConflicts.length} 行跨页签数据，总耗时: ${duration}ms`);
  
  return allConflictRowsData;
}

/**
 * 按表格分组冲突
 * @param {Array} conflicts - 冲突列表
 * @returns {Map} 按表格分组的冲突
 */
function groupConflictsBySheet(conflicts) {
  const sheetGroups = new Map();
  for (const conflict of conflicts) {
    if (!sheetGroups.has(conflict.sheet)) {
      sheetGroups.set(conflict.sheet, []);
    }
    sheetGroups.get(conflict.sheet).push(conflict);
  }
  return sheetGroups;
}

/**
 * 计算需要读取的行数
 * @param {Array} conflicts - 冲突列表
 * @returns {number} 行数
 */
function calculateRowCount(conflicts) {
  const rowNumbers = conflicts.map(c => c.row);
  const minRow = Math.min(...rowNumbers);
  const maxRow = Math.max(...rowNumbers);
  return maxRow - minRow + 1;
}

/**
 * 优化读取表格行数据
 * @param {Sheet} sheet - 表格对象
 * @param {Array} conflicts - 冲突列表
 * @param {number} lastCol - 最后一列
 * @returns {Array} 行数据数组
 */
function readSheetRowsOptimized(sheet, conflicts, lastCol) {
  const rowNumbers = conflicts.map(c => c.row);
  const minRow = Math.min(...rowNumbers);
  const maxRow = Math.max(...rowNumbers);
  const rowCount = maxRow - minRow + 1;
  
  const range = sheet.getRange(minRow, 1, rowCount, lastCol);
  return range.getValues();
}

/**
 * 获取表格表头
 * @param {Sheet} sheet - 表格对象
 * @param {Map} sheetIdDataCache - 表格ID列数据缓存
 * @param {string} sheetName - 表格名称
 * @param {number} lastCol - 最后一列
 * @returns {Array} 表头数组
 */
function getSheetHeaders(sheet, sheetIdDataCache, sheetName, lastCol) {
  let headers = sheetIdDataCache.get(sheetName).headers;
  if (!headers) {
    headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  }
  return headers;
}

/**
 * 映射行数据到冲突位置
 * @param {Array} conflicts - 冲突列表
 * @param {Array} rowsData - 行数据
 * @param {string} sheetName - 表格名称
 * @param {Map} conflictRowsData - 冲突行数据映射
 */
function mapRowsDataToConflicts(conflicts, rowsData, sheetName, conflictRowsData) {
  const rowNumbers = conflicts.map(c => c.row);
  const minRow = Math.min(...rowNumbers);
  
  for (const conflict of conflicts) {
    const rowIndex = conflict.row - minRow;
    if (rowIndex >= 0 && rowIndex < rowsData.length) {
      const key = `${sheetName}_${conflict.row}`;
      conflictRowsData.set(key, rowsData[rowIndex]);
    }
  }
}

/**
 * 映射行数据和表头到冲突位置
 * @param {Array} conflicts - 冲突列表
 * @param {Array} rowsData - 行数据
 * @param {Array} headers - 表头
 * @param {string} sheetName - 表格名称
 * @param {Map} allConflictRowsData - 所有冲突行数据映射
 */
function mapRowsDataWithHeaders(conflicts, rowsData, headers, sheetName, allConflictRowsData) {
  const rowNumbers = conflicts.map(c => c.row);
  const minRow = Math.min(...rowNumbers);
  
  for (const conflict of conflicts) {
    const rowIndex = conflict.row - minRow;
    if (rowIndex >= 0 && rowIndex < rowsData.length) {
      const key = `${sheetName}_${conflict.row}`;
      allConflictRowsData.set(key, {
        data: rowsData[rowIndex],
        headers: headers
      });
      
      // 同时更新conflict对象，避免后续重复查找
      conflict.headers = headers;
    }
  }
  
  console.log(`📊 [跨页签读取] ${sheetName}: 读取 ${minRow}-${minRow + rowsData.length - 1} 行，共 ${rowsData.length} 行数据 + 表头`);
}

/**
 * 构建表格ID列数据缓存
 * 一次性读取所有表格的ID列数据，避免重复API调用
 * @param {Spreadsheet} ss - 电子表格对象
 * @param {string} columnName - 列标题
 * @returns {Map} ID列数据缓存映射
 */
function buildSheetIdDataCache(ss, columnName) {
  const sheetIdDataCache = new Map();
  const sheets = ss.getSheets();
  
  console.log(`⚡ [批量预加载] 开始预加载所有表格的ID列数据，目标列: "${columnName}"...`);
  const startTime = Date.now();
  
  for (const sheet of sheets) {
    const sheetName = sheet.getName();
    try {
      const lastCol = sheet.getLastColumn();
      const lastRow = sheet.getLastRow();
      
      if (lastRow <= 1 || lastCol === 0) {
        console.log(`⏭️ [跳过表格] ${sheetName} - 无数据`);
        continue;
      }
      
      // 一次性读取表头和ID列数据
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      const columnIndex = headers.findIndex(header => header && header.toString() === columnName);
      
      if (columnIndex !== -1) {
        const idColumn = sheet.getRange(2, columnIndex + 1, lastRow - 1, 1).getValues().flat();
        sheetIdDataCache.set(sheetName, {
          headers,
          idColumn,
          columnIndex: columnIndex + 1,
          lastRow,
          lastCol
        });
        console.log(`📊 [缓存表格] ${sheetName} - ID列索引: ${columnIndex + 1}, 数据行数: ${idColumn.length}`);
      } else {
        console.log(`⏭️ [跳过表格] ${sheetName} - 未找到列 "${columnName}"`);
      }
    } catch (error) {
      console.warn(`⚠️ [缓存警告] 无法缓存表格 ${sheetName} 的ID列数据: ${error.message}`);
    }
  }
  
  const duration = Date.now() - startTime;
  console.log(`🚀 [批量预加载完成] 耗时: ${duration}ms, 成功缓存了 ${sheetIdDataCache.size} 个表格的ID列数据`);
  
  return sheetIdDataCache;
}

/**
 * 快速查找潜在冲突
 * 使用缓存的数据快速定位包含指定ID的所有表格位置
 * @param {string} value - ID值
 * @param {Map} sheetIdDataCache - ID列数据缓存
 * @param {string} currentSheetName - 当前表格名称
 * @returns {Array} 潜在冲突列表
 */
function findPotentialConflicts(value, sheetIdDataCache, currentSheetName) {
  console.log(`🔍 [快速查找] 在 ${sheetIdDataCache.size} 个表格中查找ID: "${value}"`);
  const potentialConflicts = [];
  
  for (const [sheetName, data] of sheetIdDataCache.entries()) {
    if (sheetName === currentSheetName) continue;
    
    const foundIndex = data.idColumn.findIndex(id => 
      id && id.toString().trim() && id.toString() === value.toString()
    );
    
    if (foundIndex !== -1) {
      potentialConflicts.push({
        sheet: sheetName,
        row: foundIndex + 2,
        column: data.columnIndex,
        sheetData: data
      });
      console.log(`🎯 [发现潜在冲突] ${sheetName} 第${foundIndex + 2}行第${data.columnIndex}列`);
    }
  }
  
  console.log(`📊 [快速查找完成] 找到 ${potentialConflicts.length} 个潜在冲突位置`);
  return potentialConflicts;
}

/**
 * 智能选择检查方式
 * 根据表格数量和性能需求自动选择最优的检查策略
 * @param {Spreadsheet} ss - 电子表格对象
 * @param {Sheet} currentSheet - 当前表格
 * @param {string} currentSheetName - 当前表格名称
 * @param {string} value - ID值
 * @param {number} currentRow - 当前行号
 * @param {string} columnName - 列标题
 * @returns {Object|null} 冲突信息或null
 */
function checkOtherSheetsConflictSmart(ss, currentSheet, currentSheetName, value, currentRow, columnName) {
  // 检查是否启用智能选择
  if (!BATCH_CHECKER_CONFIG.ENABLE_SMART_SELECTION) {
    console.log(`📊 [智能选择] 智能选择已禁用，使用传统检查方式`);
    return checkOtherSheetsConflict(ss, currentSheet, currentSheetName, value, currentRow, columnName);
  }
  
  const sheets = ss.getSheets();
  const totalSheets = sheets.length;
  
  // 智能选择策略：表格数量少时使用传统方式，多时使用批量方式
  if (totalSheets <= BATCH_CHECKER_CONFIG.SHEET_COUNT_THRESHOLD) {
    console.log(`📊 [智能选择] 表格数量较少(${totalSheets}个)，使用传统检查方式`);
    return checkOtherSheetsConflict(ss, currentSheet, currentSheetName, value, currentRow, columnName);
  } else {
    console.log(`📊 [智能选择] 表格数量较多(${totalSheets}个)，使用批量检查方式`);
    return checkOtherSheetsConflictBatch(ss, currentSheet, currentSheetName, value, currentRow, columnName);
  }
}

/**
 * 在指定表格中查找ID冲突
 * @param {Sheet} sheet - 目标表格
 * @param {string} value - ID值
 * @param {string} columnName - 列标题
 * @returns {Object|null} 冲突信息或null
 */
function findConflictInSheet(sheet, value, columnName) {
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

// ============================================================================
// 数据比较函数
// ============================================================================

/**
 * 比较两行数据是否一致（通用版本）
 * @param {Array} row1 - 第一行数据
 * @param {Array} row2 - 第二行数据
 * @param {Array} headers1 - 第一行表头
 * @param {Array} headers2 - 第二行表头
 * @returns {boolean} 是否一致
 */
function compareRowsData(row1, row2, headers1, headers2) {
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
 * 优化的行数据比较函数 - 用于批量处理
 * @param {Array} row1 - 第一行数据
 * @param {Array} row2 - 第二行数据
 * @param {Array} headers1 - 第一行表头
 * @param {Array} headers2 - 第二行表头
 * @returns {boolean} 是否一致
 */
function compareRowsDataOptimized(row1, row2, headers1, headers2) {
  // 快速长度检查
  if (!row1 || !row2 || !headers1 || !headers2) return false;
  
  const minLength = Math.min(row1.length, row2.length);
  
  // 快速检查：如果行长度差异很大，直接返回false
  if (Math.abs(row1.length - row2.length) > 5) return false;
  
  // 逐列比较，一旦发现差异就返回false
  for (let i = 0; i < minLength; i++) {
    const header1 = headers1[i];
    const header2 = headers2[i];
    
    // 只比较有意义的列（非空表头）
    if (header1 && header2 && header1.toString().trim() === header2.toString().trim()) {
      const val1 = row1[i] ? row1[i].toString().trim() : '';
      const val2 = row2[i] ? row2[i].toString().trim() : '';
      
      if (val1 !== val2) {
        return false; // 发现差异，立即返回
      }
    }
  }
  
  return true; // 所有比较的列都一致
}

// ============================================================================
// 批量冲突验证和清理
// ============================================================================

/**
 * 手动验证并清理所有冲突标记 - 批量处理优化版本
 * 这是一个兜底功能，用于清理可能过期的冲突标记
 * 
 * 性能优化特性：
 * - 批量数据读取
 * - 表头缓存机制
 * - 避免重复检查
 * - 智能冲突验证
 * 
 * @returns {Object} 清理结果统计
 */
function validateAndClearConflictMarks() {
  console.log(`🚀 [批量清理] 开始验证当前表格冲突标记`);
  
  const startTime = new Date().getTime();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentSheet = ss.getActiveSheet();
  
  // 初始化统计变量
  const stats = {
    totalSheets: 1,
    totalCells: 0,
    conflictCells: 0,
    clearedConflicts: 0,
    validConflicts: 0
  };
  
  const results = {
    sheets: [],
    summary: {}
  };
  
  try {
    const sheetName = currentSheet.getName();
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
      const lastRow = currentSheet.getLastRow();
      const lastColumn = currentSheet.getLastColumn();
      
      if (lastRow <= 1 || lastColumn === 0) {
        console.log(`⏭️ [跳过表格] ${sheetName} - 无数据`);
        sheetResult.totalCells = 0;
        results.sheets.push(sheetResult);
      } else {
        console.log(`📊 [表格范围] ${sheetName} - ${lastRow}行 × ${lastColumn}列`);
        
        // 🚀 批量数据读取优化
        const batchData = readSheetDataBatch(currentSheet, lastRow, lastColumn);
        sheetResult.totalCells = lastRow * lastColumn;
        stats.totalCells += sheetResult.totalCells;
        
        // 🔍 冲突单元格检测
        const conflictCells = detectConflictCells(batchData);
        sheetResult.conflictCells = conflictCells.length;
        stats.conflictCells += conflictCells.length;
        
        console.log(`🎯 [冲突统计] 发现 ${conflictCells.length} 个冲突标记单元格`);
        
        if (conflictCells.length > 0) {
                  // ✅ 批量冲突验证（启用超级批量验证）
        const validationResults = validateConflictsBatch(conflictCells, sheetName, batchData.headers, true, true);
          
          // 🔄 批量单元格更新
          const updateResults = updateCellsBatch(validationResults, currentSheet);
          
          // 统计结果
          sheetResult.clearedConflicts = updateResults.clearedCount;
          sheetResult.validConflicts = updateResults.validCount;
          stats.clearedConflicts += updateResults.clearedCount;
          stats.validConflicts += updateResults.validCount;
          
          console.log(`📊 [更新统计] 清理: ${updateResults.clearedCount}, 保留: ${updateResults.validCount}`);
        }
        
        console.log(`✅ [表格完成] ${sheetName} - 冲突标记: ${sheetResult.conflictCells}, 清理: ${sheetResult.clearedConflicts}, 保留: ${sheetResult.validConflicts}`);
      }
      
    } catch (sheetError) {
      console.error(`❌ [表格处理错误] ${sheetName}:`, sheetError);
      sheetResult.errors.push(`表格处理失败: ${sheetError.message}`);
    }
    
    results.sheets.push(sheetResult);
    
    // 汇总结果
    const duration = new Date().getTime() - startTime;
    results.summary = {
      ...stats,
      duration,
      success: true
    };
    
    console.log(`🏁 [清理完成] 总耗时: ${duration}ms`);
    console.log(`📊 [最终统计] 表格: ${stats.totalSheets}, 单元格: ${stats.totalCells}, 冲突标记: ${stats.conflictCells}, 清理: ${stats.clearedConflicts}, 保留: ${stats.validConflicts}`);
    
    // 显示结果给用户
    showCleanupResults(results);
    
    return results;
    
  } catch (error) {
    console.error(`💥 [清理异常] 批量清理过程中出现错误:`, error);
    
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
 * 批量读取表格数据
 * @param {Sheet} sheet - 表格对象
 * @param {number} lastRow - 最后一行
 * @param {number} lastColumn - 最后一列
 * @returns {Object} 批量数据对象
 */
function readSheetDataBatch(sheet, lastRow, lastColumn) {
  console.log(`⚡ [批量读取] 开始批量读取表格数据...`);
  const batchStartTime = Date.now();
  
  const range = sheet.getRange(1, 1, lastRow, lastColumn);
  const data = {
    backgrounds: range.getBackgrounds(),
    values: range.getValues(),
    notes: range.getNotes(),
    headers: sheet.getRange(1, 1, 1, lastColumn).getValues()[0]
  };
  
  const batchReadTime = Date.now() - batchStartTime;
  console.log(`⚡ [批量读取完成] 耗时: ${batchReadTime}ms`);
  
  return data;
}

/**
 * 检测冲突标记单元格
 * @param {Object} batchData - 批量数据对象
 * @returns {Array} 冲突单元格列表
 */
function detectConflictCells(batchData) {
  console.log(`🔍 [冲突检测] 开始检测冲突标记单元格...`);
  
  const { notes, values, headers } = batchData;
  const conflictCells = [];
  
  for (let row = 0; row < notes.length; row++) {
    for (let col = 0; col < notes[row].length; col++) {
      const note = notes[row][col];
      if (note && NoteManager.getSystemNote(note, NOTE_CONSTANTS.TYPES.CONFLICT)) {
        conflictCells.push({
          row: row + 1,
          col: col + 1,
          value: values[row][col],
          note: note,
          header: headers[col]
        });
      }
    }
  }
  
  return conflictCells;
}

/**
 * 批量验证冲突状态
 * @param {Array} conflictCells - 冲突单元格列表
 * @param {string} sheetName - 表格名称
 * @param {Array} headers - 表头
 * @param {boolean} useBatchValidation - 是否使用批量验证优化（默认true）
 * @param {boolean} useSuperBatch - 是否使用超级批量验证（默认false，仅在冲突单元格数量很多时启用）
 * @returns {Array} 验证结果数组
 */
function validateConflictsBatch(conflictCells, sheetName, headers, useBatchValidation = true, useSuperBatch = false) {
  console.log(`✅ [批量验证] 开始验证 ${conflictCells.length} 个冲突单元格`);
  
  // 智能选择验证方式
  if (useSuperBatch && BATCH_CHECKER_CONFIG.ENABLE_SUPER_BATCH_VALIDATION && conflictCells.length > BATCH_CHECKER_CONFIG.SUPER_BATCH_THRESHOLD) {
    console.log(`🚀 [智能选择] 冲突单元格数量很多(${conflictCells.length}个)，使用超级批量验证优化`);
    return validateCellConflictsSuperBatch(conflictCells, sheetName, headers);
  } else if (useBatchValidation && BATCH_CHECKER_CONFIG.ENABLE_BATCH_VALIDATION && conflictCells.length > BATCH_CHECKER_CONFIG.CONFLICT_CELL_THRESHOLD) {
    console.log(`🚀 [智能选择] 冲突单元格数量较多(${conflictCells.length}个)，使用批量验证优化`);
    return validateCellConflictsBatch(conflictCells, sheetName, headers);
  } else {
    console.log(`📊 [智能选择] 冲突单元格数量较少(${conflictCells.length}个)，使用传统验证方式`);
    return validateConflictsBatchTraditional(conflictCells, sheetName, headers);
  }
}

/**
 * 传统批量验证冲突状态（保持向后兼容）
 * @param {Array} conflictCells - 冲突单元格列表
 * @param {string} sheetName - 表格名称
 * @param {Array} headers - 表头
 * @returns {Array} 验证结果数组
 */
function validateConflictsBatchTraditional(conflictCells, sheetName, headers) {
  console.log(`📊 [传统验证] 开始传统方式验证 ${conflictCells.length} 个冲突单元格`);
  
  const results = [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentSheet = ss.getSheetByName(sheetName);
  
  // 预加载所有表格的表头信息，避免重复读取
  const sheetHeadersCache = buildSheetHeadersCache(ss);
  
  // 批量验证每个冲突单元格
  for (let i = 0; i < conflictCells.length; i++) {
    const cell = conflictCells[i];
    console.log(`🔍 [验证进度] ${i + 1}/${conflictCells.length} - ${sheetName} 第${cell.row}行第${cell.col}列`);
    
    try {
      // 快速验证冲突状态
      const conflicts = validateCellConflict(cell, sheetName, headers, sheetHeadersCache);
      const stillHasConflict = conflicts.length > 0;
      
      results.push({
        ...cell,
        stillHasConflict,
        conflicts, // 保存冲突信息，避免重复检查
        validationTime: Date.now()
      });
      
    } catch (error) {
      console.error(`❌ [验证错误] ${sheetName} 第${cell.row}行第${cell.col}列: ${error.message}`);
      results.push({
        ...cell,
        stillHasConflict: false, // 出错时默认清除标记
        error: error.message,
        validationTime: Date.now()
      });
    }
  }
  
  console.log(`✅ [传统验证完成] 共验证 ${results.length} 个单元格`);
  return results;
}

/**
 * 构建表格表头缓存
 * @param {Spreadsheet} ss - 电子表格对象
 * @returns {Map} 表头缓存映射
 */
function buildSheetHeadersCache(ss) {
  const sheetHeadersCache = new Map();
  const sheets = ss.getSheets();
  
  for (const sheet of sheets) {
    const sheetName = sheet.getName();
    if (!sheetHeadersCache.has(sheetName)) {
      try {
        const lastCol = sheet.getLastColumn();
        if (lastCol > 0) {
          const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
          sheetHeadersCache.set(sheetName, headers);
        }
      } catch (error) {
        console.warn(`⚠️ [缓存警告] 无法缓存表格 ${sheetName} 的表头: ${error.message}`);
      }
    }
  }
  
  return sheetHeadersCache;
}

/**
 * 验证单个单元格的冲突状态
 * @param {Object} cell - 单元格信息
 * @param {string} sheetName - 表格名称
 * @param {Array} currentHeaders - 当前表格表头
 * @param {Map} sheetHeadersCache - 表头缓存
 * @returns {Array} 冲突信息数组
 */
function validateCellConflict(cell, sheetName, currentHeaders, sheetHeadersCache) {
  const { value, header } = cell;
  
  if (!value || !value.toString().trim()) {
    return [];
  }
  
  // 检查是否是ID列
  if (!header || !header.toString().endsWith(ID_CHECKER_CONFIG.ID_COLUMN_SUFFIX)) {
    return [];
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  // 快速检查当前表格
  const currentSheet = ss.getSheetByName(sheetName);
  const currentConflict = checkCurrentSheetConflict(currentSheet, value, cell.row, header);
  if (currentConflict) {
    return [currentConflict];
  }
  
  // 检查其他表格
  let foundConflicts = [];
  
  for (const sheet of sheets) {
    if (sheet.getName() === sheetName) continue;
    
    const cachedHeaders = sheetHeadersCache.get(sheet.getName());
    if (!cachedHeaders) continue;
    
    const conflict = findConflictInSheet(sheet, value, header);
    if (conflict) {
      // 🚀 重要：不仅要检查ID是否相同，还要验证整行数据是否一致
      const currentRowData = currentSheet.getRange(cell.row, 1, 1, currentSheet.getLastColumn()).getValues()[0];
      const conflictRowData = sheet.getRange(conflict.row, 1, 1, sheet.getLastColumn()).getValues()[0];
      
      // 使用优化的行比较逻辑验证数据一致性
      if (compareRowsDataOptimized(currentRowData, conflictRowData, currentHeaders, cachedHeaders)) {
        console.log(`🔍 [数据一致] ID相同但整行数据一致，不构成冲突`);
        continue; // 数据一致，继续检查其他表格
      } else {
        console.log(`🚨 [发现冲突] ID相同且整行数据不一致，构成真实冲突`);
        foundConflicts.push(conflict);
      }
    }
  }
  
  return foundConflicts;
}

/**
 * 批量验证单元格冲突状态 - 优化版本
 * 使用批量预加载和缓存机制大幅提升批量验证性能
 * @param {Array} cells - 单元格信息数组
 * @param {string} sheetName - 表格名称
 * @param {Array} currentHeaders - 当前表格表头
 * @returns {Array} 验证结果数组
 */
function validateCellConflictsBatch(cells, sheetName, currentHeaders) {
  console.log(`🚀 [批量验证] 开始批量验证 ${cells.length} 个单元格的冲突状态`);
  const startTime = Date.now();
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentSheet = ss.getSheetByName(sheetName);
  
  // 过滤出需要检查的ID列单元格
  const idCells = cells.filter(cell => 
    cell.header && cell.header.toString().endsWith(ID_CHECKER_CONFIG.ID_COLUMN_SUFFIX)
  );
  
  if (idCells.length === 0) {
    console.log(`⏭️ [批量验证] 没有需要检查的ID列单元格`);
    return cells.map(cell => ({ ...cell, stillHasConflict: false, conflicts: [] }));
  }
  
  console.log(`🎯 [批量验证] 需要检查 ${idCells.length} 个ID列单元格`);
  
  // 🚀 批量预加载所有表格的ID列数据
  const allColumnNames = [...new Set(idCells.map(cell => cell.header))];
  const columnDataCache = new Map();
  
  for (const columnName of allColumnNames) {
    const cache = buildSheetIdDataCache(ss, columnName);
    columnDataCache.set(columnName, cache);
  }
  
  // 批量验证每个单元格
  const results = [];
  for (let i = 0; i < cells.length; i++) {
    const cell = cells[i];
    console.log(`🔍 [验证进度] ${i + 1}/${cells.length} - ${sheetName} 第${cell.row}行第${cell.col}列`);
    
    try {
      let conflicts = [];
      
      if (idCells.some(idCell => idCell.row === cell.row && idCell.col === cell.col)) {
        // 这是ID列单元格，需要详细检查
        conflicts = validateCellConflictBatch(cell, sheetName, currentHeaders, columnDataCache);
      }
      
      const stillHasConflict = conflicts.length > 0;
      
      results.push({
        ...cell,
        stillHasConflict,
        conflicts,
        validationTime: Date.now()
      });
      
    } catch (error) {
      console.error(`❌ [验证错误] ${sheetName} 第${cell.row}行第${cell.col}列: ${error.message}`);
      results.push({
        ...cell,
        stillHasConflict: false,
        error: error.message,
        validationTime: Date.now()
      });
    }
  }
  
  const duration = Date.now() - startTime;
  console.log(`✅ [批量验证完成] 共验证 ${results.length} 个单元格，耗时: ${duration}ms`);
  
  return results;
}

/**
 * 批量验证单个单元格的冲突状态
 * @param {Object} cell - 单元格信息
 * @param {string} sheetName - 表格名称
 * @param {Array} currentHeaders - 当前表格表头
 * @param {Map} columnDataCache - 列数据缓存
 * @returns {Array} 冲突信息数组
 */
function validateCellConflictBatch(cell, sheetName, currentHeaders, columnDataCache) {
  const { value, header } = cell;
  
  if (!value || !value.toString().trim()) {
    return [];
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentSheet = ss.getSheetByName(sheetName);
  
  // 快速检查当前表格
  const currentConflict = checkCurrentSheetConflict(currentSheet, value, cell.row, header);
  if (currentConflict) {
    return [currentConflict];
  }
  
  // 使用缓存的列数据快速检查其他表格
  const columnCache = columnDataCache.get(header);
  if (!columnCache) {
    return [];
  }
  
  // 快速查找潜在冲突
  const potentialConflicts = findPotentialConflicts(value, columnCache, sheetName);
  
  if (potentialConflicts.length === 0) {
    return [];
  }
  
  // 🚀 关键优化：一次性读取当前行数据
  const currentRowData = currentSheet.getRange(cell.row, 1, 1, currentSheet.getLastColumn()).getValues()[0];
  
  // 🚀 关键优化：批量读取所有潜在冲突行的数据
  const conflictRowsData = batchReadConflictRowsData(ss, potentialConflicts);
  
  // 批量验证冲突
  const foundConflicts = [];
  for (const conflict of potentialConflicts) {
    const conflictRowData = conflictRowsData.get(conflict.sheet + '_' + conflict.row);
    if (!conflictRowData) continue;
    
    const conflictHeaders = columnCache.get(conflict.sheet).headers;
    
    if (!compareRowsDataOptimized(currentRowData, conflictRowData, currentHeaders, conflictHeaders)) {
      foundConflicts.push(conflict);
    }
  }
  
  return foundConflicts;
}

/**
 * 超级批量验证单元格冲突状态 - 终极优化版本
 * 一次性读取所有需要验证的数据，最大化减少API调用
 * @param {Array} cells - 单元格信息数组
 * @param {string} sheetName - 表格名称
 * @param {Array} currentHeaders - 当前表格表头
 * @returns {Array} 验证结果数组
 */
function validateCellConflictsSuperBatch(cells, sheetName, currentHeaders) {
  console.log(`🚀 [超级批量验证] 开始超级批量验证 ${cells.length} 个单元格的冲突状态`);
  const startTime = Date.now();
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentSheet = ss.getSheetByName(sheetName);
  
  // 过滤出需要检查的ID列单元格
  const idCells = cells.filter(cell => 
    cell.header && cell.header.toString().endsWith(ID_CHECKER_CONFIG.ID_COLUMN_SUFFIX)
  );
  
  if (idCells.length === 0) {
    console.log(`⏭️ [超级批量验证] 没有需要检查的ID列单元格`);
    return cells.map(cell => ({ ...cell, stillHasConflict: false, conflicts: [] }));
  }
  
  console.log(`🎯 [超级批量验证] 需要检查 ${idCells.length} 个ID列单元格`);
  
  // 🚀 超级优化：一次性读取当前表格中所有相关行
  const currentRowsData = batchReadCurrentSheetRows(currentSheet, idCells);
  console.log(`⚡ [超级批量验证] 一次性读取当前表格 ${idCells.length} 行数据`);
  
  // 🚀 超级优化：一次性预加载所有表格的ID列数据
  const allColumnNames = [...new Set(idCells.map(cell => cell.header))];
  const columnDataCache = new Map();
  
  for (const columnName of allColumnNames) {
    const cache = buildSheetIdDataCache(ss, columnName);
    columnDataCache.set(columnName, cache);
  }
  
  // 🚀 超级优化：收集所有需要验证的潜在冲突位置
  const allPotentialConflicts = collectAllPotentialConflicts(idCells, columnDataCache, sheetName);
  console.log(`🎯 [超级批量验证] 总共发现 ${allPotentialConflicts.length} 个潜在冲突位置`);
  
  // 🚀 超级优化：一次性读取所有潜在冲突行的数据
  const allConflictRowsData = batchReadConflictRowsData(ss, allPotentialConflicts);
  console.log(`⚡ [超级批量验证] 一次性读取所有冲突行数据`);
  
  // 批量验证每个单元格
  const results = [];
  for (let i = 0; i < cells.length; i++) {
    const cell = cells[i];
    console.log(`🔍 [验证进度] ${i + 1}/${cells.length} - ${sheetName} 第${cell.row}行第${cell.col}列`);
    
    try {
      let conflicts = [];
      
      if (idCells.some(idCell => idCell.row === cell.row && idCell.col === cell.col)) {
        // 这是ID列单元格，需要详细检查
        conflicts = validateCellConflictSuperBatch(
          cell, 
          sheetName, 
          currentHeaders, 
          columnDataCache, 
          currentRowsData, 
          allConflictRowsData
        );
      }
      
      const stillHasConflict = conflicts.length > 0;
      
      results.push({
        ...cell,
        stillHasConflict,
        conflicts,
        validationTime: Date.now()
      });
      
    } catch (error) {
      console.error(`❌ [验证错误] ${sheetName} 第${cell.row}行第${cell.col}列: ${error.message}`);
      results.push({
        ...cell,
        stillHasConflict: false,
        error: error.message,
        validationTime: Date.now()
      });
    }
  }
  
  const duration = Date.now() - startTime;
  console.log(`✅ [超级批量验证完成] 共验证 ${results.length} 个单元格，耗时: ${duration}ms`);
  
  return results;
}

/**
 * 批量读取当前表格中所有相关行数据
 * @param {Sheet} currentSheet - 当前表格
 * @param {Array} idCells - ID列单元格列表
 * @returns {Map} 行数据映射 (key: 行号, value: 行数据)
 */
function batchReadCurrentSheetRows(currentSheet, idCells) {
  const currentRowsData = new Map();
  
  if (idCells.length === 0) return currentRowsData;
  
  // 获取所有需要读取的行号
  const rowNumbers = idCells.map(cell => cell.row);
  const minRow = Math.min(...rowNumbers);
  const maxRow = Math.max(...rowNumbers);
  
  // 一次性读取所有相关行
  const lastCol = currentSheet.getLastColumn();
  if (lastCol > 0) {
    const range = currentSheet.getRange(minRow, 1, maxRow - minRow + 1, lastCol);
    const rowsData = range.getValues();
    
    // 映射到对应的行号
    for (let i = 0; i < rowsData.length; i++) {
      const actualRow = minRow + i;
      currentRowsData.set(actualRow, rowsData[i]);
    }
  }
  
  return currentRowsData;
}

/**
 * 收集所有潜在冲突位置
 * @param {Array} idCells - ID列单元格列表
 * @param {Map} columnDataCache - 列数据缓存
 * @param {string} sheetName - 当前表格名称
 * @returns {Array} 所有潜在冲突位置
 */
function collectAllPotentialConflicts(idCells, columnDataCache, sheetName) {
  const allPotentialConflicts = [];
  
  for (const cell of idCells) {
    const columnCache = columnDataCache.get(cell.header);
    if (!columnCache) continue;
    
    const potentialConflicts = findPotentialConflicts(cell.value, columnCache, sheetName);
    allPotentialConflicts.push(...potentialConflicts);
  }
  
  // 去重（同一个表格的同一行可能被多个单元格引用）
  const uniqueConflicts = new Map();
  for (const conflict of allPotentialConflicts) {
    const key = `${conflict.sheet}_${conflict.row}`;
    if (!uniqueConflicts.has(key)) {
      uniqueConflicts.set(key, conflict);
    }
  }
  
  return Array.from(uniqueConflicts.values());
}

/**
 * 超级批量验证单个单元格的冲突状态
 * @param {Object} cell - 单元格信息
 * @param {string} sheetName - 表格名称
 * @param {Array} currentHeaders - 当前表格表头
 * @param {Map} columnDataCache - 列数据缓存
 * @param {Map} currentRowsData - 当前表格行数据缓存
 * @param {Map} allConflictRowsData - 所有冲突行数据缓存
 * @returns {Array} 冲突信息数组
 */
function validateCellConflictSuperBatch(cell, sheetName, currentHeaders, columnDataCache, currentRowsData, allConflictRowsData) {
  const { value, header, row } = cell;
  
  if (!value || !value.toString().trim()) {
    return [];
  }
  
  // 快速检查当前表格
  const currentRowData = currentRowsData.get(row);
  if (!currentRowData) {
    console.warn(`⚠️ [超级批量验证] 无法获取当前行数据: ${sheetName} 第${row}行`);
    return [];
  }
  
  // 使用缓存的列数据快速检查其他表格
  const columnCache = columnDataCache.get(header);
  if (!columnCache) {
    return [];
  }
  
  // 快速查找潜在冲突
  const potentialConflicts = findPotentialConflicts(value, columnCache, sheetName);
  
  if (potentialConflicts.length === 0) {
    return [];
  }
  
  // 批量验证冲突（使用预加载的数据）
  const foundConflicts = [];
  for (const conflict of potentialConflicts) {
    const conflictRowData = allConflictRowsData.get(conflict.sheet + '_' + conflict.row);
    if (!conflictRowData) continue;
    
    const conflictHeaders = columnCache.get(conflict.sheet).headers;
    
    if (!compareRowsDataOptimized(currentRowData, conflictRowData, currentHeaders, conflictHeaders)) {
      foundConflicts.push(conflict);
    }
  }
  
  return foundConflicts;
}

/**
 * 批量更新单元格
 * @param {Array} validationResults - 验证结果数组
 * @param {Sheet} currentSheet - 当前表格
 * @returns {Object} 更新结果统计
 */
function updateCellsBatch(validationResults, currentSheet) {
  console.log(`🔄 [批量更新] 开始批量更新 ${validationResults.length} 个单元格`);
  
  let clearedCount = 0;
  let validCount = 0;
  let errorCount = 0;
  
  // 分类处理单元格
  const { cellsToClear, cellsToUpdate } = categorizeCells(validationResults);
  
  // 批量清除冲突标记
  if (cellsToClear.length > 0) {
    clearedCount = clearConflictMarks(cellsToClear, currentSheet);
  }
  
  // 批量更新有效冲突的注释
  if (cellsToUpdate.length > 0) {
    validCount = updateConflictNotes(cellsToUpdate, currentSheet);
  }
  
  console.log(`🔄 [批量更新完成] 清除: ${clearedCount}, 更新: ${validCount}, 错误: ${errorCount}`);
  
  return {
    clearedCount,
    validCount,
    errorCount
  };
}

/**
 * 分类单元格处理类型
 * @param {Array} validationResults - 验证结果数组
 * @returns {Object} 分类结果
 */
function categorizeCells(validationResults) {
  const cellsToClear = [];
  const cellsToUpdate = [];
  
  for (const result of validationResults) {
    if (result.error) {
      // 有错误的单元格默认清除标记
      cellsToClear.push(result);
    } else if (!result.stillHasConflict) {
      // 无冲突的单元格清除标记
      cellsToClear.push(result);
    } else {
      // 仍有冲突的单元格更新注释
      cellsToUpdate.push(result);
    }
  }
  
  return { cellsToClear, cellsToUpdate };
}

/**
 * 清除冲突标记
 * @param {Array} cellsToClear - 需要清除的单元格列表
 * @param {Sheet} currentSheet - 当前表格
 * @returns {number} 成功清除的数量
 */
function clearConflictMarks(cellsToClear, currentSheet) {
  console.log(`🧹 [批量清除] 开始清除 ${cellsToClear.length} 个过期冲突标记`);
  
  let clearedCount = 0;
  
  for (const cell of cellsToClear) {
    try {
      const cellRange = currentSheet.getRange(cell.row, cell.col);
      cellRange.setBackground(null);
      NoteManager.removeMarkFromCell(cellRange, NOTE_CONSTANTS.TYPES.CONFLICT);
      
      console.log(`✅ [清除完成] ${currentSheet.getName()} 第${cell.row}行第${cell.col}列`);
      clearedCount++;
    } catch (error) {
      console.error(`❌ [清除失败] ${currentSheet.getName()} 第${cell.row}行第${cell.col}列: ${error.message}`);
    }
  }
  
  return clearedCount;
}

/**
 * 更新冲突注释
 * @param {Array} cellsToUpdate - 需要更新的单元格列表
 * @param {Sheet} currentSheet - 当前表格
 * @returns {number} 成功更新的数量
 */
function updateConflictNotes(cellsToUpdate, currentSheet) {
  console.log(`📝 [批量更新] 开始更新 ${cellsToUpdate.length} 个有效冲突的注释`);
  
  let updatedCount = 0;
  
  for (const cell of cellsToUpdate) {
    try {
      const cellRange = currentSheet.getRange(cell.row, cell.col);
      
      // 🚀 性能优化：使用验证阶段已经收集的冲突信息，避免重复检查
      if (cell.conflicts && cell.conflicts.length > 0) {
        // 使用验证阶段收集的冲突信息更新注释
        const conflictLocations = cell.conflicts.map(loc => `${loc.sheet} 第${loc.row}行`).join('\n');
        const userNote = `在以下位置重复:\n${conflictLocations}`;
        
        const currentNote = cellRange.getNote();
        const updatedNote = NoteManager.addSystemNote(currentNote, NOTE_CONSTANTS.TYPES.CONFLICT, userNote);
        cellRange.setNote(updatedNote);
        
        console.log(`📝 [更新完成] ${currentSheet.getName()} 第${cell.row}行第${cell.col}列 - 使用验证阶段信息`);
        updatedCount++;
      } else {
        console.log(`📝 [更新完成] ${currentSheet.getName()} 第${cell.row}行第${cell.col}列 - 无冲突信息`);
      }
    } catch (error) {
      console.error(`❌ [更新失败] ${currentSheet.getName()} 第${cell.row}行第${cell.col}列: ${error.message}`);
    }
  }
  
  return updatedCount;
}

// ============================================================================
// 兼容性函数（保持向后兼容）
// ============================================================================

/**
 * 重新检查单个单元格的冲突状态（兼容性函数）
 * @param {Range} cellRange - 单元格范围
 * @param {string} sheetName - 表格名称
 * @returns {boolean} 是否仍有冲突
 */
function recheckCellConflictStatusCurrentSheet(cellRange, sheetName) {
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
  const conflicts = checkSingleIdConflict({
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

// ============================================================================
// 结果展示函数
// ============================================================================

/**
 * 显示清理结果给用户
 * @param {Object} results - 清理结果对象
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

// ============================================================================
// 实时冲突检查函数（保持原有功能）
// ============================================================================

/**
 * 检查ID冲突并标记（实时编辑检测）
 * @param {Object} editedCell - 编辑的单元格信息
 */
function checkIdConflicts(editedCell) {
  console.log(`🚀 [实时检查] 开始处理编辑单元格`);
  
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
    const conflicts = checkSingleIdConflict({
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
        NOTE_CONSTANTS.TYPES.CONFLICT,
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
 * 性能测试函数 - 比较传统方式和批量方式的性能差异
 * @param {string} testId - 测试用的ID值
 * @param {string} columnName - 列标题
 * @param {string} sheetName - 表格名称
 * @param {number} row - 行号
 * @returns {Object} 性能测试结果
 */
function performanceTestIdChecker(testId, columnName, sheetName, row) {
  console.log(`🧪 [性能测试] 开始测试ID检查器性能`);
  console.log(`📋 [测试参数] ID: "${testId}", 列: "${columnName}", 表格: "${sheetName}", 行: ${row}`);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentSheet = ss.getSheetByName(sheetName);
  
  if (!currentSheet) {
    console.error(`❌ [测试错误] 找不到表格: ${sheetName}`);
    return { error: `找不到表格: ${sheetName}` };
  }
  
  const results = {
    testId,
    columnName,
    sheetName,
    row,
    traditional: {},
    batch: {},
    improvement: {}
  };
  
  try {
    // 测试传统方式
    console.log(`🔄 [测试1] 传统检查方式`);
    const traditionalStart = Date.now();
    const traditionalResult = checkOtherSheetsConflict(ss, currentSheet, sheetName, testId, row, columnName);
    const traditionalTime = Date.now() - traditionalStart;
    
    results.traditional = {
      result: traditionalResult,
      time: traditionalTime,
      success: true
    };
    
    console.log(`✅ [传统方式] 耗时: ${traditionalTime}ms, 结果:`, traditionalResult);
    
    // 测试批量方式
    console.log(`🔄 [测试2] 批量检查方式`);
    const batchStart = Date.now();
    const batchResult = checkOtherSheetsConflictBatch(ss, currentSheet, sheetName, testId, row, columnName);
    const batchTime = Date.now() - batchStart;
    
    results.batch = {
      result: batchResult,
      time: batchTime,
      success: true
    };
    
    console.log(`✅ [批量方式] 耗时: ${batchTime}ms, 结果:`, batchResult);
    
    // 计算性能提升
    if (traditionalTime > 0) {
      const timeImprovement = ((traditionalTime - batchTime) / traditionalTime * 100).toFixed(2);
      const speedup = (traditionalTime / batchTime).toFixed(2);
      
      results.improvement = {
        timeImprovement: `${timeImprovement}%`,
        speedup: `${speedup}x`,
        traditionalTime,
        batchTime
      };
      
      console.log(`🚀 [性能提升] 时间减少: ${timeImprovement}%, 速度提升: ${speedup}倍`);
      
      // 显示结果给用户
      let message = `🧪 性能测试完成！\n\n`;
      message += `📊 测试结果:\n`;
      message += `• 传统方式: ${traditionalTime}ms\n`;
      message += `• 批量方式: ${batchTime}ms\n`;
      message += `• 性能提升: ${timeImprovement}%\n`;
      message += `• 速度倍数: ${speedup}x\n\n`;
      
      if (timeImprovement > 0) {
        message += `🎉 批量方式性能更优！`;
      } else {
        message += `⚠️ 传统方式性能更优（可能是数据量较小）`;
      }
      
      SpreadsheetApp.getActiveSpreadsheet().toast(message, '性能测试结果', 8);
    }
    
  } catch (error) {
    console.error(`💥 [测试异常] 性能测试过程中出现错误:`, error);
    results.error = error.message;
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `性能测试失败: ${error.message}`,
      '测试错误',
      5
    );
  }
  
  console.log(`🏁 [性能测试完成] 测试结果:`, results);
  return results;
}

/**
 * 批量性能测试 - 测试多个ID的检查性能
 * @param {Array} testCases - 测试用例数组，每个包含 {id, columnName, sheetName, row}
 * @returns {Object} 批量测试结果
 */
function batchPerformanceTest(testCases) {
  console.log(`🧪 [批量性能测试] 开始测试 ${testCases.length} 个用例`);
  
  const results = {
    totalCases: testCases.length,
    successful: 0,
    failed: 0,
    totalTraditionalTime: 0,
    totalBatchTime: 0,
    cases: []
  };
  
  for (let i = 0; i < testCases.length; i++) {
    const testCase = testCases[i];
    console.log(`🔍 [测试进度] ${i + 1}/${testCases.length} - ID: "${testCase.id}"`);
    
    try {
      const caseResult = performanceTestIdChecker(
        testCase.id,
        testCase.columnName,
        testCase.sheetName,
        testCase.row
      );
      
      if (caseResult.error) {
        results.failed++;
        results.cases.push({ ...testCase, error: caseResult.error });
      } else {
        results.successful++;
        results.totalTraditionalTime += caseResult.traditional.time;
        results.totalBatchTime += caseResult.batch.time;
        results.cases.push({ ...testCase, result: caseResult });
      }
      
    } catch (error) {
      results.failed++;
      results.cases.push({ ...testCase, error: error.message });
    }
  }
  
  // 计算总体性能提升
  if (results.totalTraditionalTime > 0) {
    const totalTimeImprovement = ((results.totalTraditionalTime - results.totalBatchTime) / results.totalTraditionalTime * 100).toFixed(2);
    const totalSpeedup = (results.totalTraditionalTime / results.totalBatchTime).toFixed(2);
    
    results.summary = {
      totalTimeImprovement: `${totalTimeImprovement}%`,
      totalSpeedup: `${totalSpeedup}x`,
      averageTraditionalTime: (results.totalTraditionalTime / results.successful).toFixed(2),
      averageBatchTime: (results.totalBatchTime / results.successful).toFixed(2)
    };
    
    console.log(`📊 [批量测试总结] 成功: ${results.successful}, 失败: ${results.failed}`);
    console.log(`🚀 [总体性能] 时间减少: ${totalTimeImprovement}%, 速度提升: ${totalSpeedup}倍`);
    
    // 显示总结结果
    let message = `🧪 批量性能测试完成！\n\n`;
    message += `📊 总体结果:\n`;
    message += `• 成功用例: ${results.successful}/${results.totalCases}\n`;
    message += `• 传统方式总耗时: ${results.totalTraditionalTime}ms\n`;
    message += `• 批量方式总耗时: ${results.totalBatchTime}ms\n`;
    message += `• 总体性能提升: ${totalTimeImprovement}%\n`;
    message += `• 总体速度倍数: ${totalSpeedup}x\n\n`;
    message += `📈 平均性能:\n`;
    message += `• 传统方式平均: ${results.summary.averageTraditionalTime}ms\n`;
    message += `• 批量方式平均: ${results.summary.averageBatchTime}ms`;
    
    SpreadsheetApp.getActiveSpreadsheet().toast(message, '批量测试完成', 10);
  }
  
  return results;
}

/**
 * 性能对比测试 - 比较三种验证方式的性能差异
 * @param {Array} testCells - 测试用的冲突单元格列表
 * @param {string} sheetName - 表格名称
 * @param {Array} headers - 表头
 * @returns {Object} 性能对比结果
 */
function performanceComparisonTest(testCells, sheetName, headers) {
  console.log(`🧪 [性能对比测试] 开始对比三种验证方式的性能`);
  console.log(`📋 [测试参数] 冲突单元格数量: ${testCells.length}, 表格: ${sheetName}`);
  
  if (testCells.length === 0) {
    console.log(`⏭️ [测试跳过] 没有测试数据`);
    return { error: '没有测试数据' };
  }
  
  const results = {
    testCellsCount: testCells.length,
    sheetName,
    traditional: {},
    batch: {},
    superBatch: {},
    comparison: {}
  };
  
  try {
    // 测试1：传统验证方式
    console.log(`🔄 [测试1] 传统验证方式`);
    const traditionalStart = Date.now();
    const traditionalResult = validateConflictsBatchTraditional(testCells, sheetName, headers);
    const traditionalTime = Date.now() - traditionalStart;
    
    results.traditional = {
      result: traditionalResult,
      time: traditionalTime,
      success: true
    };
    
    console.log(`✅ [传统方式] 耗时: ${traditionalTime}ms, 验证结果: ${traditionalResult.length} 个`);
    
    // 测试2：批量验证方式
    console.log(`🔄 [测试2] 批量验证方式`);
    const batchStart = Date.now();
    const batchResult = validateCellConflictsBatch(testCells, sheetName, headers);
    const batchTime = Date.now() - batchStart;
    
    results.batch = {
      result: batchResult,
      time: batchTime,
      success: true
    };
    
    console.log(`✅ [批量方式] 耗时: ${batchTime}ms, 验证结果: ${batchResult.length} 个`);
    
    // 测试3：超级批量验证方式
    console.log(`🔄 [测试3] 超级批量验证方式`);
    const superBatchStart = Date.now();
    const superBatchResult = validateCellConflictsSuperBatch(testCells, sheetName, headers);
    const superBatchTime = Date.now() - superBatchStart;
    
    results.superBatch = {
      result: superBatchResult,
      time: superBatchTime,
      success: true
    };
    
    console.log(`✅ [超级批量方式] 耗时: ${superBatchTime}ms, 验证结果: ${superBatchResult.length} 个`);
    
    // 计算性能对比
    if (traditionalTime > 0) {
      const batchImprovement = ((traditionalTime - batchTime) / traditionalTime * 100).toFixed(2);
      const superBatchImprovement = ((traditionalTime - superBatchTime) / traditionalTime * 100).toFixed(2);
      const batchSpeedup = (traditionalTime / batchTime).toFixed(2);
      const superBatchSpeedup = (traditionalTime / superBatchTime).toFixed(2);
      
      results.comparison = {
        batchImprovement: `${batchImprovement}%`,
        superBatchImprovement: `${superBatchImprovement}%`,
        batchSpeedup: `${batchSpeedup}x`,
        superBatchSpeedup: `${superBatchSpeedup}x`,
        traditionalTime,
        batchTime,
        superBatchTime
      };
      
      console.log(`🚀 [性能提升] 批量方式: ${batchImprovement}%, 超级批量方式: ${superBatchImprovement}%`);
      console.log(`📈 [速度倍数] 批量方式: ${batchSpeedup}x, 超级批量方式: ${superBatchSpeedup}x`);
      
      // 显示结果给用户
      let message = `🧪 性能对比测试完成！\n\n`;
      message += `📊 测试结果:\n`;
      message += `• 传统方式: ${traditionalTime}ms\n`;
      message += `• 批量方式: ${batchTime}ms (提升${batchImprovement}%)\n`;
      message += `• 超级批量: ${superBatchTime}ms (提升${superBatchImprovement}%)\n\n`;
      
      if (superBatchTime < batchTime && superBatchTime < traditionalTime) {
        message += `🏆 超级批量方式性能最佳！`;
      } else if (batchTime < traditionalTime) {
        message += `🥇 批量方式性能更优！`;
      } else {
        message += `⚠️ 传统方式性能更优（可能是数据量较小）`;
      }
      
      SpreadsheetApp.getActiveSpreadsheet().toast(message, '性能对比结果', 10);
    }
    
  } catch (error) {
    console.error(`💥 [测试异常] 性能对比测试过程中出现错误:`, error);
    results.error = error.message;
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `性能对比测试失败: ${error.message}`,
      '测试错误',
      5
    );
  }
  
  console.log(`🏁 [性能对比测试完成] 测试结果:`, results);
  return results;
}

/**
 * 快速性能测试 - 使用当前表格的冲突标记进行测试
 * @returns {Object} 快速测试结果
 */
function quickPerformanceTest() {
  console.log(`🚀 [快速性能测试] 开始使用当前表格的冲突标记进行性能测试`);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentSheet = ss.getActiveSheet();
  const sheetName = currentSheet.getName();
  
  // 获取当前表格的冲突标记
  const lastRow = currentSheet.getLastRow();
  const lastColumn = currentSheet.getLastColumn();
  
  if (lastRow <= 1 || lastColumn === 0) {
    console.log(`⏭️ [快速测试] 当前表格无数据`);
    return { error: '当前表格无数据' };
  }
  
  // 批量读取表格数据
  const batchData = readSheetDataBatch(currentSheet, lastRow, lastColumn);
  
  // 检测冲突标记单元格
  const conflictCells = detectConflictCells(batchData);
  
  if (conflictCells.length === 0) {
    console.log(`⏭️ [快速测试] 当前表格无冲突标记`);
    return { error: '当前表格无冲突标记' };
  }
  
  console.log(`🎯 [快速测试] 发现 ${conflictCells.length} 个冲突标记，开始性能测试`);
  
  // 进行性能对比测试
  const results = performanceComparisonTest(conflictCells, sheetName, batchData.headers);
  
  return results;
}

/**
 * 单行编辑检查的智能选择函数
 * 根据表格数量和性能需求自动选择最优的检查策略
 * @param {Spreadsheet} ss - 电子表格对象
 * @param {Sheet} currentSheet - 当前表格
 * @param {string} currentSheetName - 当前表格名称
 * @param {string} value - ID值
 * @param {number} currentRow - 当前行号
 * @param {string} columnName - 列标题
 * @returns {Object|null} 冲突信息或null
 */
function checkOtherSheetsConflictSingleRowSmart(ss, currentSheet, currentSheetName, value, currentRow, columnName) {
  // 检查是否启用智能选择
  if (!BATCH_CHECKER_CONFIG.ENABLE_SMART_SELECTION) {
    console.log(`📊 [单行智能选择] 智能选择已禁用，使用传统检查方式`);
    return checkOtherSheetsConflict(ss, currentSheet, currentSheetName, value, currentRow, columnName);
  }
  
  const sheets = ss.getSheets();
  const totalSheets = sheets.length;
  
  // 智能选择策略：针对单行编辑检查的优化
  if (totalSheets <= 2) {
    console.log(`📊 [单行智能选择] 表格数量很少(${totalSheets}个)，使用传统检查方式`);
    return checkOtherSheetsConflict(ss, currentSheet, currentSheetName, value, currentRow, columnName);
  } else if (totalSheets <= 5) {
    console.log(`📊 [单行智能选择] 表格数量较少(${totalSheets}个)，使用批量检查方式`);
    return checkOtherSheetsConflictBatch(ss, currentSheet, currentSheetName, value, currentRow, columnName);
  } else {
    console.log(`📊 [单行智能选择] 表格数量较多(${totalSheets}个)，使用单行优化检查方式`);
    return checkOtherSheetsConflictSingleRowOptimized(ss, currentSheet, currentSheetName, value, currentRow, columnName);
  }
}

/**
 * 单行编辑检查性能测试 - 专门测试跨页签检查的性能
 * @param {string} testId - 测试用的ID值
 * @param {string} columnName - 列标题
 * @param {string} sheetName - 表格名称
 * @param {number} row - 行号
 * @returns {Object} 性能测试结果
 */
function singleRowPerformanceTest(testId, columnName, sheetName, row) {
  console.log(`🧪 [单行性能测试] 开始测试单行编辑跨页签检查性能`);
  console.log(`📋 [测试参数] ID: "${testId}", 列: "${columnName}", 表格: "${sheetName}", 行: ${row}`);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentSheet = ss.getSheetByName(sheetName);
  
  if (!currentSheet) {
    console.error(`❌ [测试错误] 找不到表格: ${sheetName}`);
    return { error: `找不到表格: ${sheetName}` };
  }
  
  const results = {
    testId,
    columnName,
    sheetName,
    row,
    traditional: {},
    batch: {},
    singleRowOptimized: {},
    improvement: {}
  };
  
  try {
    // 测试1：传统检查方式
    console.log(`🔄 [测试1] 传统检查方式`);
    const traditionalStart = Date.now();
    const traditionalResult = checkOtherSheetsConflict(ss, currentSheet, sheetName, testId, row, columnName);
    const traditionalTime = Date.now() - traditionalStart;
    
    results.traditional = {
      result: traditionalResult,
      time: traditionalTime,
      success: true
    };
    
    console.log(`✅ [传统方式] 耗时: ${traditionalTime}ms, 结果:`, traditionalResult);
    
    // 测试2：批量检查方式
    console.log(`🔄 [测试2] 批量检查方式`);
    const batchStart = Date.now();
    const batchResult = checkOtherSheetsConflictBatch(ss, currentSheet, sheetName, testId, row, columnName);
    const batchTime = Date.now() - batchStart;
    
    results.batch = {
      result: batchResult,
      time: batchTime,
      success: true
    };
    
    console.log(`✅ [批量方式] 耗时: ${batchTime}ms, 结果:`, batchResult);
    
    // 测试3：单行优化检查方式
    console.log(`🔄 [测试3] 单行优化检查方式`);
    const singleRowStart = Date.now();
    const singleRowResult = checkOtherSheetsConflictSingleRowOptimized(ss, currentSheet, sheetName, testId, row, columnName);
    const singleRowTime = Date.now() - singleRowStart;
    
    results.singleRowOptimized = {
      result: singleRowResult,
      time: singleRowTime,
      success: true
    };
    
    console.log(`✅ [单行优化方式] 耗时: ${singleRowTime}ms, 结果:`, singleRowResult);
    
    // 计算性能提升
    if (traditionalTime > 0) {
      const batchImprovement = ((traditionalTime - batchTime) / traditionalTime * 100).toFixed(2);
      const singleRowImprovement = ((traditionalTime - singleRowTime) / traditionalTime * 100).toFixed(2);
      const batchSpeedup = (traditionalTime / batchTime).toFixed(2);
      const singleRowSpeedup = (traditionalTime / singleRowTime).toFixed(2);
      
      results.improvement = {
        batchImprovement: `${batchImprovement}%`,
        singleRowImprovement: `${singleRowImprovement}%`,
        batchSpeedup: `${batchSpeedup}x`,
        singleRowSpeedup: `${singleRowSpeedup}x`,
        traditionalTime,
        batchTime,
        singleRowTime
      };
      
      console.log(`🚀 [性能提升] 批量方式: ${batchImprovement}%, 单行优化方式: ${singleRowImprovement}%`);
      console.log(`📈 [速度倍数] 批量方式: ${batchSpeedup}x, 单行优化方式: ${singleRowSpeedup}x`);
      
      // 显示结果给用户
      let message = `🧪 单行编辑检查性能测试完成！\n\n`;
      message += `📊 测试结果:\n`;
      message += `• 传统方式: ${traditionalTime}ms\n`;
      message += `• 批量方式: ${batchTime}ms (提升${batchImprovement}%)\n`;
      message += `• 单行优化: ${singleRowTime}ms (提升${singleRowImprovement}%)\n\n`;
      
      if (singleRowTime < batchTime && singleRowTime < traditionalTime) {
        message += `🏆 单行优化方式性能最佳！`;
      } else if (batchTime < traditionalTime) {
        message += `🥇 批量方式性能更优！`;
      } else {
        message += `⚠️ 传统方式性能更优（可能是数据量较小）`;
      }
      
      SpreadsheetApp.getActiveSpreadsheet().toast(message, '单行性能测试结果', 8);
    }
    
  } catch (error) {
    console.error(`💥 [测试异常] 单行性能测试过程中出现错误:`, error);
    results.error = error.message;
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `单行性能测试失败: ${error.message}`,
      '测试错误',
      5
    );
  }
  
  console.log(`🏁 [单行性能测试完成] 测试结果:`, results);
  return results;
}

/**
 * 跨页签检查性能分析 - 分析不同检查方式的性能特点
 * @returns {Object} 性能分析结果
 */
function crossSheetPerformanceAnalysis() {
  console.log(`📊 [性能分析] 开始分析跨页签检查的性能特点`);
  
  const analysis = {
    traditional: {
      description: "传统检查方式",
      advantages: [
        "实现简单，易于理解",
        "适合表格数量少的情况",
        "内存占用小"
      ],
      disadvantages: [
        "API调用次数多",
        "性能随表格数量线性下降",
        "不适合大量数据"
      ],
      bestFor: "表格数量 ≤ 2个，数据量小"
    },
    batch: {
      description: "批量检查方式",
      advantages: [
        "批量预加载ID列数据",
        "减少重复API调用",
        "性能提升明显"
      ],
      disadvantages: [
        "仍需要逐个读取冲突行",
        "内存占用增加",
        "实现复杂度中等"
      ],
      bestFor: "表格数量 3-5个，中等数据量"
    },
    singleRowOptimized: {
      description: "单行优化检查方式",
      advantages: [
        "一次性读取所有跨页签数据",
        "最大化减少API调用",
        "性能提升最显著",
        "专门针对单行编辑优化"
      ],
      disadvantages: [
        "内存占用最大",
        "实现复杂度高",
        "初始化开销较大"
      ],
      bestFor: "表格数量 > 5个，大量数据，频繁编辑"
    }
  };
  
  console.log(`📊 [性能分析完成] 分析结果:`, analysis);
  
  // 显示分析结果给用户
  let message = `📊 跨页签检查性能分析\n\n`;
  message += `🏆 单行优化方式:\n`;
  message += `• 一次性读取所有跨页签数据\n`;
  message += `• 性能提升最显著\n`;
  message += `• 适合表格数量多的情况\n\n`;
  
  message += `🥇 批量方式:\n`;
  message += `• 批量预加载ID列数据\n`;
  message += `• 性能提升明显\n`;
  message += `• 适合中等规模数据\n\n`;
  
  message += `📋 传统方式:\n`;
  message += `• 实现简单，易于理解\n`;
  message += `• 适合表格数量少的情况`;
  
  SpreadsheetApp.getActiveSpreadsheet().toast(message, '性能分析结果', 10);
  
  return analysis;
}

/**
 * 内存使用监控 - 监控当前内存使用情况
 * @returns {Object} 内存使用信息
 */
function monitorMemoryUsage() {
  console.log(`📊 [内存监控] 开始监控内存使用情况`);
  
  const memoryInfo = {
    timestamp: new Date().toISOString(),
    available: false,
    details: {}
  };
  
  try {
    // 尝试获取内存使用信息（仅在支持的环境中可用）
    if (typeof Memory !== 'undefined') {
      const memoryUsage = Memory.getMemoryUsage();
      if (memoryUsage) {
        memoryInfo.available = true;
        memoryInfo.details = {
          used: memoryUsage.used,
          usedMB: (memoryUsage.used / 1024 / 1024).toFixed(2),
          total: memoryUsage.total,
          totalMB: (memoryUsage.total / 1024 / 1024).toFixed(2),
          percentage: ((memoryUsage.used / memoryUsage.total) * 100).toFixed(2)
        };
        
        console.log(`📊 [内存监控] 内存使用: ${memoryInfo.details.usedMB}MB / ${memoryInfo.details.totalMB}MB (${memoryInfo.details.percentage}%)`);
        
        // 检查是否超过警告阈值
        if (memoryUsage.used > BATCH_CHECKER_CONFIG.MEMORY_MANAGEMENT.MEMORY_WARNING_THRESHOLD * 1024 * 1024) {
          console.warn(`⚠️ [内存警告] 内存使用过高: ${memoryInfo.details.usedMB}MB > ${BATCH_CHECKER_CONFIG.MEMORY_MANAGEMENT.MEMORY_WARNING_THRESHOLD}MB`);
        }
      }
    }
    
    // 获取当前配置信息
    memoryInfo.config = {
      maxRowsPerBatch: BATCH_CHECKER_CONFIG.MEMORY_MANAGEMENT.MAX_ROWS_PER_BATCH,
      maxSheetsPerBatch: BATCH_CHECKER_CONFIG.MEMORY_MANAGEMENT.MAX_SHEETS_PER_BATCH,
      memoryWarningThreshold: BATCH_CHECKER_CONFIG.MEMORY_MANAGEMENT.MEMORY_WARNING_THRESHOLD,
      batchProcessingEnabled: BATCH_CHECKER_CONFIG.MEMORY_MANAGEMENT.ENABLE_BATCH_PROCESSING
    };
    
    // 显示内存信息给用户
    if (memoryInfo.available) {
      let message = `📊 内存使用监控\n\n`;
      message += `💾 内存使用: ${memoryInfo.details.usedMB}MB / ${memoryInfo.details.totalMB}MB\n`;
      message += `📈 使用率: ${memoryInfo.details.percentage}%\n\n`;
      message += `⚙️ 当前配置:\n`;
      message += `• 单次最大行数: ${memoryInfo.config.maxRowsPerBatch}\n`;
      message += `• 单次最大表格数: ${memoryInfo.config.maxSheetsPerBatch}\n`;
      message += `• 内存警告阈值: ${memoryInfo.config.memoryWarningThreshold}MB\n`;
      message += `• 分批处理: ${memoryInfo.config.batchProcessingEnabled ? '启用' : '禁用'}`;
      
      if (parseFloat(memoryInfo.details.percentage) > 80) {
        message += `\n\n⚠️ 内存使用率较高，建议调整配置或启用分批处理`;
      }
      
      SpreadsheetApp.getActiveSpreadsheet().toast(message, '内存监控结果', 8);
    } else {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        '内存监控不可用，但分批处理机制已启用以保护系统资源',
        '内存监控',
        5
      );
    }
    
  } catch (error) {
    console.error(`❌ [内存监控错误] 监控过程中出现错误: ${error.message}`);
    memoryInfo.error = error.message;
  }
  
  return memoryInfo;
}

/**
 * 动态调整内存管理配置 - 根据当前情况自动调整参数
 * @param {Object} options - 调整选项
 * @returns {Object} 调整后的配置
 */
function adjustMemoryManagementConfig(options = {}) {
  console.log(`⚙️ [配置调整] 开始动态调整内存管理配置`);
  
  const currentConfig = { ...BATCH_CHECKER_CONFIG.MEMORY_MANAGEMENT };
  const adjustments = {};
  
  try {
    // 获取当前内存使用情况
    let memoryUsage = null;
    if (typeof Memory !== 'undefined') {
      try {
        memoryUsage = Memory.getMemoryUsage();
      } catch (e) {
        console.warn(`⚠️ [配置调整] 无法获取内存使用信息: ${e.message}`);
      }
    }
    
    // 根据内存使用情况调整配置
    if (memoryUsage) {
      const usedMB = memoryUsage.used / 1024 / 1024;
      const percentage = (memoryUsage.used / memoryUsage.total) * 100;
      
      console.log(`📊 [配置调整] 当前内存使用: ${usedMB.toFixed(2)}MB (${percentage.toFixed(2)}%)`);
      
      // 内存使用率高的调整
      if (percentage > 80) {
        adjustments.maxRowsPerBatch = Math.max(100, Math.floor(currentConfig.MAX_ROWS_PER_BATCH * 0.5));
        adjustments.maxSheetsPerBatch = Math.max(3, Math.floor(currentConfig.MAX_SHEETS_PER_BATCH * 0.5));
        adjustments.batchDelay = Math.min(500, currentConfig.BATCH_DELAY * 2);
        
        console.log(`⚠️ [配置调整] 内存使用率过高，降低批处理参数`);
      }
      // 内存使用率中等的调整
      else if (percentage > 60) {
        adjustments.maxRowsPerBatch = Math.max(500, Math.floor(currentConfig.MAX_ROWS_PER_BATCH * 0.8));
        adjustments.maxSheetsPerBatch = Math.max(5, Math.floor(currentConfig.MAX_SHEETS_PER_BATCH * 0.8));
        
        console.log(`📊 [配置调整] 内存使用率中等，适度降低批处理参数`);
      }
      // 内存使用率低的调整
      else if (percentage < 30) {
        adjustments.maxRowsPerBatch = Math.min(2000, Math.floor(currentConfig.MAX_ROWS_PER_BATCH * 1.5));
        adjustments.maxSheetsPerBatch = Math.min(20, Math.floor(currentConfig.MAX_SHEETS_PER_BATCH * 1.5));
        adjustments.batchDelay = Math.max(50, Math.floor(currentConfig.BATCH_DELAY * 0.8));
        
        console.log(`✅ [配置调整] 内存使用率较低，提高批处理参数`);
      }
    }
    
    // 应用用户指定的调整
    if (options.maxRowsPerBatch) {
      adjustments.maxRowsPerBatch = options.maxRowsPerBatch;
    }
    if (options.maxSheetsPerBatch) {
      adjustments.maxSheetsPerBatch = options.maxSheetsPerBatch;
    }
    if (options.batchDelay) {
      adjustments.batchDelay = options.batchDelay;
    }
    if (options.memoryWarningThreshold) {
      adjustments.memoryWarningThreshold = options.memoryWarningThreshold;
    }
    
    // 应用调整
    Object.keys(adjustments).forEach(key => {
      const upperKey = key.toUpperCase();
      if (BATCH_CHECKER_CONFIG.MEMORY_MANAGEMENT.hasOwnProperty(upperKey)) {
        BATCH_CHECKER_CONFIG.MEMORY_MANAGEMENT[upperKey] = adjustments[key];
        console.log(`⚙️ [配置调整] ${key}: ${currentConfig[upperKey]} → ${adjustments[key]}`);
      }
    });
    
    // 显示调整结果
    let message = `⚙️ 内存管理配置已调整\n\n`;
    message += `📊 调整详情:\n`;
    Object.keys(adjustments).forEach(key => {
      const upperKey = key.toUpperCase();
      if (currentConfig[upperKey] !== undefined) {
        message += `• ${key}: ${currentConfig[upperKey]} → ${adjustments[key]}\n`;
      }
    });
    
    if (Object.keys(adjustments).length === 0) {
      message += `• 无需调整，当前配置已是最优`;
    }
    
    SpreadsheetApp.getActiveSpreadsheet().toast(message, '配置调整完成', 6);
    
  } catch (error) {
    console.error(`❌ [配置调整错误] 调整过程中出现错误: ${error.message}`);
    adjustments.error = error.message;
  }
  
  return {
    original: currentConfig,
    adjustments: adjustments,
    current: BATCH_CHECKER_CONFIG.MEMORY_MANAGEMENT
  };
}

/**
 * 内存压力测试 - 测试系统在不同数据量下的内存使用情况
 * @param {number} maxSheets - 最大表格数量
 * @param {number} maxRows - 最大行数
 * @returns {Object} 压力测试结果
 */
function memoryStressTest(maxSheets = 20, maxRows = 5000) {
  console.log(`🧪 [内存压力测试] 开始测试系统内存承受能力`);
  console.log(`📋 [测试参数] 最大表格数: ${maxSheets}, 最大行数: ${maxRows}`);
  
  const testResults = {
    maxSheets,
    maxRows,
    tests: [],
    recommendations: []
  };
  
  try {
    // 测试不同规模的配置
    const testScenarios = [
      { sheets: 5, rows: 1000, description: "小规模测试" },
      { sheets: 10, rows: 2000, description: "中等规模测试" },
      { sheets: 15, rows: 3000, description: "大规模测试" },
      { sheets: maxSheets, rows: maxRows, description: "极限规模测试" }
    ];
    
    for (const scenario of testScenarios) {
      console.log(`🧪 [压力测试] 开始 ${scenario.description}: ${scenario.sheets} 表格 × ${scenario.rows} 行`);
      
      const testStart = Date.now();
      const memoryBefore = monitorMemoryUsage();
      
      // 模拟大量数据处理
      const mockConflicts = generateMockConflicts(scenario.sheets, scenario.rows);
      
      // 测试分批处理逻辑
      const shouldBatch = shouldUseBatchProcessing(mockConflicts);
      
      const testEnd = Date.now();
      const memoryAfter = monitorMemoryUsage();
      
      const testResult = {
        scenario: scenario.description,
        sheets: scenario.sheets,
        rows: scenario.rows,
        conflicts: mockConflicts.length,
        shouldBatch: shouldBatch,
        duration: testEnd - testStart,
        memoryBefore: memoryBefore,
        memoryAfter: memoryAfter
      };
      
      testResults.tests.push(testResult);
      console.log(`✅ [压力测试] ${scenario.description} 完成，建议分批处理: ${shouldBatch}`);
      
      // 添加延迟，避免系统压力过大
      Utilities.sleep(1000);
    }
    
    // 生成建议
    testResults.recommendations = generateMemoryRecommendations(testResults);
    
    // 显示测试结果
    let message = `🧪 内存压力测试完成！\n\n`;
    message += `📊 测试结果:\n`;
    testResults.tests.forEach(test => {
      message += `• ${test.scenario}: ${test.sheets}表格×${test.rows}行 → 建议分批: ${test.shouldBatch ? '是' : '否'}\n`;
    });
    
    if (testResults.recommendations.length > 0) {
      message += `\n💡 优化建议:\n`;
      testResults.recommendations.forEach(rec => {
        message += `• ${rec}\n`;
      });
    }
    
    SpreadsheetApp.getActiveSpreadsheet().toast(message, '压力测试完成', 10);
    
  } catch (error) {
    console.error(`💥 [压力测试异常] 测试过程中出现错误: ${error.message}`);
    testResults.error = error.message;
  }
  
  return testResults;
}

/**
 * 生成模拟冲突数据用于压力测试
 * @param {number} sheetCount - 表格数量
 * @param {number} rowCount - 行数量
 * @returns {Array} 模拟冲突数据
 */
function generateMockConflicts(sheetCount, rowCount) {
  const conflicts = [];
  
  for (let i = 0; i < sheetCount; i++) {
    const sheetName = `TestSheet${i + 1}`;
    const conflictsPerSheet = Math.ceil(rowCount / sheetCount);
    
    for (let j = 0; j < conflictsPerSheet; j++) {
      conflicts.push({
        sheet: sheetName,
        row: Math.floor(Math.random() * rowCount) + 2,
        column: Math.floor(Math.random() * 10) + 1
      });
    }
  }
  
  return conflicts;
}

/**
 * 根据压力测试结果生成内存优化建议
 * @param {Object} testResults - 测试结果
 * @returns {Array} 优化建议列表
 */
function generateMemoryRecommendations(testResults) {
  const recommendations = [];
  
  // 分析测试结果
  const largeScaleTests = testResults.tests.filter(t => t.sheets > 10 || t.rows > 2000);
  const batchRecommendations = testResults.tests.filter(t => t.shouldBatch);
  
  if (largeScaleTests.length > 0) {
    recommendations.push("对于大规模数据（>10表格或>2000行），建议启用分批处理");
  }
  
  if (batchRecommendations.length > 0) {
    recommendations.push("系统已自动建议分批处理，当前配置适合大多数场景");
  }
  
  // 根据测试结果调整配置
  const maxSheetsTested = Math.max(...testResults.tests.map(t => t.sheets));
  const maxRowsTested = Math.max(...testResults.tests.map(t => t.rows));
  
  if (maxSheetsTested > BATCH_CHECKER_CONFIG.MEMORY_MANAGEMENT.MAX_SHEETS_PER_BATCH) {
    recommendations.push(`建议将单次最大表格数调整为 ${Math.ceil(maxSheetsTested * 0.8)}`);
  }
  
  if (maxRowsTested > BATCH_CHECKER_CONFIG.MEMORY_MANAGEMENT.MAX_ROWS_PER_BATCH) {
    recommendations.push(`建议将单次最大行数调整为 ${Math.ceil(maxRowsTested * 0.8)}`);
  }
  
  return recommendations;
}