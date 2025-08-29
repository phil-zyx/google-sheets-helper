/**
 * Google Sheets ID Conflict Checker 
 * Core Principles:
 * 1. Unified Pipeline: All checks (real-time, batch) use the same core logic.
 * 2. Batch Everything: All reads and writes to the spreadsheet are batched to minimize API calls.
 * 3. Separation of Concerns: Logic is split into Data Loading, In-Memory Validation, and Sheet Updating.
 */

/**
 * Global cache for sheet headers to avoid repeated API calls
 */
const HEADER_CACHE = new Map();

/**
 * Global cache for sheet dimensions to avoid repeated getLastRow/getLastColumn calls
 */
const DIMENSION_CACHE = new Map();

/**
 * sheet data reader that minimizes API calls
 * @param {Sheet} sheet - The sheet to read data from
 * @returns {Object} All necessary sheet data in one operation
 */
function readSheetData(sheet) {
  const sheetName = sheet.getName();
  console.log(`📊 开始读取表格: ${sheetName}`);
  
  try {
    // Use getDataRange() to automatically get the data bounds - single API call
    const dataRange = sheet.getDataRange();
    const startTime = new Date().getTime();
    
    // Batch read all data in one operation - 4 API calls instead of multiple
    const [values, notes, backgrounds] = [
      dataRange.getValues(),
      dataRange.getNotes(),
      dataRange.getBackgrounds()
    ];
    
    const readTime = new Date().getTime() - startTime;
    console.log(`⚡ [数据读取] ${sheetName}: ${values.length}行 × ${values[0]?.length || 0}列, 耗时: ${readTime}ms`);
    
    // Cache dimensions for future use
    const dimensions = {
      lastRow: dataRange.getLastRow(),
      lastColumn: dataRange.getLastColumn()
    };
    DIMENSION_CACHE.set(sheetName, dimensions);
    
    // Cache headers for future use
    if (values.length > 0) {
      HEADER_CACHE.set(sheetName, values[0]);
    }
    
    return {
      values,
      notes,
      backgrounds,
      lastRow: dimensions.lastRow,
      lastColumn: dimensions.lastColumn,
      headers: values[0] || [],
      sheetName
    };
  } catch (error) {
    console.error(`❌ [优化读取失败] ${sheetName}: ${error.message}`);
    // Fallback to traditional method
    return readSheetDataFallback(sheet);
  }
}

/**
 * Fallback method for sheet data reading when optimization fails
 * @param {Sheet} sheet - The sheet to read data from
 * @returns {Object} Sheet data using traditional methods
 */
function readSheetDataFallback(sheet) {
  const sheetName = sheet.getName();
  console.log(`⚠️ [回退读取] 使用传统方法读取: ${sheetName}`);
  
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  
  if (lastRow <= 1 || lastColumn === 0) {
    return {
      values: [],
      notes: [],
      backgrounds: [],
      lastRow: 0,
      lastColumn: 0,
      headers: [],
      sheetName
    };
  }
  
  const range = sheet.getRange(1, 1, lastRow, lastColumn);
  return {
    values: range.getValues(),
    notes: range.getNotes(),
    backgrounds: range.getBackgrounds(),
    lastRow,
    lastColumn,
    headers: sheet.getRange(1, 1, 1, lastColumn).getValues()[0],
    sheetName
  };
}

/**
 * Get sheet headers with caching to avoid repeated API calls
 * @param {Sheet} sheet - The sheet to get headers from
 * @returns {Array} Array of header values
 */
function getSheetHeadersCached(sheet) {
  const sheetName = sheet.getName();
  
  if (!HEADER_CACHE.has(sheetName)) {
    console.log(`📝 [表头缓存] 首次读取表头: ${sheetName}`);
    const lastCol = sheet.getLastColumn();
    if (lastCol > 0) {
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      HEADER_CACHE.set(sheetName, headers);
    } else {
      HEADER_CACHE.set(sheetName, []);
    }
  }
  
  return HEADER_CACHE.get(sheetName);
}

/**
 * Get sheet dimensions with caching to avoid repeated API calls
 * @param {Sheet} sheet - The sheet to get dimensions from
 * @returns {Object} Object with lastRow and lastColumn
 */
function getSheetDimensionsCached(sheet) {
  const sheetName = sheet.getName();
  
  if (!DIMENSION_CACHE.has(sheetName)) {
    console.log(`📏 [尺寸缓存] 首次读取尺寸: ${sheetName}`);
    const dimensions = {
      lastRow: sheet.getLastRow(),
      lastColumn: sheet.getLastColumn()
    };
    DIMENSION_CACHE.set(sheetName, dimensions);
  }
  
  return DIMENSION_CACHE.get(sheetName);
}

/**
 * Optimized batch update using getRangeList for multiple non-contiguous cells
 * @param {Sheet} sheet - The sheet to update
 * @param {Array} updates - Array of update objects with row, col, and properties
 * @returns {Object} Update statistics
 */
function batchUpdateCells(sheet, updates) {
  if (updates.length === 0) {
    return { success: true, updatedCount: 0, message: "没有需要更新的单元格" };
  }
  
  const sheetName = sheet.getName();
  console.log(`🚀 [批量更新]: ${sheetName}, 共${updates.length}个单元格`);
  
  const startTime = new Date().getTime();
  let successCount = 0;
  let errorCount = 0;
  
  try {
    // Group updates by operation type for batch processing
    const updatesByType = groupUpdatesByType(updates);
    
    // 使用 Set 来跟踪实际更新的单元格，避免重复计算
    const updatedCells = new Set();
    
    // Process background updates in batch
    if (updatesByType.backgrounds.length > 0) {
      const bgResult = batchUpdateBackgrounds(sheet, updatesByType.backgrounds);
      // 记录成功更新的单元格
      bgResult.successCells.forEach(cell => updatedCells.add(`${cell.row}-${cell.col}`));
      errorCount += bgResult.errorCount;
    }
    
    // Process note updates in batch
    if (updatesByType.notes.length > 0) {
      const noteResult = batchUpdateNotes(sheet, updatesByType.notes);
      // 记录成功更新的单元格
      noteResult.successCells.forEach(cell => updatedCells.add(`${cell.row}-${cell.col}`));
      errorCount += noteResult.errorCount;
    }
    
    // 返回实际更新的单元格数量，而不是操作次数
    const actualUpdatedCount = updatedCells.size;
    
    const duration = new Date().getTime() - startTime;
    console.log(`✅ [批量更新完成] ${sheetName}: 成功${actualUpdatedCount}个单元格, 失败${errorCount}个, 耗时: ${duration}ms`);
    
    return {
      success: true,
      updatedCount: actualUpdatedCount,  // 这里返回实际的单元格数量
      errorCount,
      duration,
      message: `批量更新完成: 成功${actualUpdatedCount}个单元格, 失败${errorCount}个`
    };
    
  } catch (error) {
    console.error(`❌ [批量更新异常] ${sheetName}: ${error.message}`);
    // Fallback to individual updates
    return batchUpdateCellsFallback(sheet, updates);
  }
}

/**
 * Group updates by operation type for efficient batch processing
 * @param {Array} updates - Array of update objects
 * @returns {Object} Grouped updates by type
 */
function groupUpdatesByType(updates) {
  const grouped = {
    backgrounds: [],
    notes: []
  };
  
  updates.forEach(update => {
    if (update.background !== undefined) {
      grouped.backgrounds.push(update);
    }
    if (update.note !== undefined) {
      grouped.notes.push(update);
    }
  });
  
  return grouped;
}

/**
 * background updates using contiguous ranges
 * @param {Sheet} sheet - The sheet to update
 * @param {Array} backgroundUpdates - Array of background update objects
 * @returns {Object} Update result
 */
function batchUpdateBackgrounds(sheet, backgroundUpdates) {
  try {
    console.log(`🎨 [背景更新] 开始处理 ${backgroundUpdates.length} 个背景更新`);
    
    // Sort updates by row and column for better batching
    const sortedUpdates = [...backgroundUpdates].sort((a, b) => {
      if (a.row !== b.row) return a.row - b.row;
      return a.col - b.col;
    });
    
    let successCount = 0;
    let errorCount = 0;
    const successCells = []; // 新增：记录成功更新的单元格
    
    // Group updates into contiguous ranges for batch processing
    let currentRange = null;
    const ranges = [];
    
    for (const update of sortedUpdates) {
      if (!currentRange) {
        currentRange = {
          startRow: update.row,
          endRow: update.row,
          startCol: update.col,
          endCol: update.col,
          background: update.background,
          updates: [update]
        };
      } else if (
        update.row === currentRange.endRow &&
        update.col === currentRange.endCol + 1 &&
        update.background === currentRange.background
      ) {
        // Extend current range horizontally
        currentRange.endCol = update.col;
        currentRange.updates.push(update);
      } else if (
        update.row === currentRange.endRow + 1 &&
        update.col === currentRange.startCol &&
        update.background === currentRange.background
      ) {
        // Extend current range vertically
        currentRange.endRow = update.row;
        currentRange.updates.push(update);
      } else {
        // Start new range
        ranges.push(currentRange);
        currentRange = {
          startRow: update.row,
          endRow: update.row,
          startCol: update.col,
          endCol: update.col,
          background: update.background,
          updates: [update]
        };
      }
    }
    
    // Add the last range
    if (currentRange) {
      ranges.push(currentRange);
    }
    
    console.log(`📊 [范围分组] 将 ${backgroundUpdates.length} 个更新分组为 ${ranges.length} 个连续范围`);
    
    // Process each range in batch
    for (const range of ranges) {
      try {
        const numRows = range.endRow - range.startRow + 1;
        const numCols = range.endCol - range.startCol + 1;
        
        // Use single API call for each contiguous range
        const sheetRange = sheet.getRange(range.startRow, range.startCol, numRows, numCols);
        sheetRange.setBackground(range.background);
        
        // 记录成功更新的单元格
        range.updates.forEach(update => {
          successCells.push({ row: update.row, col: update.col });
        });
        
        successCount += range.updates.length;
      } catch (err) {
        console.error(`❌ [范围更新失败] 行${range.startRow}-${range.endRow}, 列${range.startCol}-${range.endCol}: ${err.message}`);
        errorCount += range.updates.length;
      }
    }
    
    console.log(`🎨 [背景更新完成] 成功: ${successCount}个, 失败: ${errorCount}个`);
    return { successCells, errorCount }; // 修改返回值结构
    
  } catch (error) {
    console.error(`❌ [背景更新失败]: ${error.message}`);
    return { successCells: [], errorCount: backgroundUpdates.length };
  }
}

/**
 * batch note updates using contiguous ranges
 * @param {Sheet} sheet - The sheet to update
 * @param {Array} noteUpdates - Array of note update objects
 * @returns {Object} Update result
 */
function batchUpdateNotes(sheet, noteUpdates) {
  try {
    console.log(`📝 [注释更新] 开始处理 ${noteUpdates.length} 个注释更新`);
    
    // Sort updates by row and column for better batching
    const sortedUpdates = [...noteUpdates].sort((a, b) => {
      if (a.row !== b.row) return a.row - b.row;
      return a.col - b.col;
    });
    
    let successCount = 0;
    let errorCount = 0;
    const successCells = []; // 新增：记录成功更新的单元格
    
    // Group updates into contiguous ranges for batch processing
    let currentRange = null;
    const ranges = [];
    
    for (const update of sortedUpdates) {
      if (!currentRange) {
        currentRange = {
          startRow: update.row,
          endRow: update.row,
          startCol: update.col,
          endCol: update.col,
          notes: [update.note],
          updates: [update]
        };
      } else if (
        update.row === currentRange.endRow &&
        update.col === currentRange.endCol + 1
      ) {
        // Extend current range horizontally
        currentRange.endCol = update.col;
        currentRange.notes.push(update.note);
        currentRange.updates.push(update);
      } else if (
        update.row === currentRange.endRow + 1 &&
        update.col === currentRange.startCol
      ) {
        // Extend current range vertically
        currentRange.endRow = update.row;
        currentRange.notes.push(update.note);
        currentRange.updates.push(update);
      } else {
        // Start new range
        ranges.push(currentRange);
        currentRange = {
          startRow: update.row,
          endRow: update.row,
          startCol: update.col,
          endCol: update.col,
          notes: [update.note],
          updates: [update]
        };
      }
    }
    
    // Add the last range
    if (currentRange) {
      ranges.push(currentRange);
    }
    
    console.log(`📊 [范围分组] 将 ${noteUpdates.length} 个更新分组为 ${ranges.length} 个连续范围`);
    
    // Process each range in batch
    for (const range of ranges) {
      try {
        const numRows = range.endRow - range.startRow + 1;
        const numCols = range.endCol - range.startCol + 1;
        
        // Use single API call for each contiguous range
        const sheetRange = sheet.getRange(range.startRow, range.startCol, numRows, numCols);
        
        // Create 2D array for notes
        const notesArray = [];
        for (let row = 0; row < numRows; row++) {
          const rowNotes = [];
          for (let col = 0; col < numCols; col++) {
            const index = row * numCols + col;
            rowNotes.push(range.notes[index] || '');
          }
          notesArray.push(rowNotes);
        }
        
        sheetRange.setNotes(notesArray);
        
        // 记录成功更新的单元格
        range.updates.forEach(update => {
          successCells.push({ row: update.row, col: update.col });
        });
        
        successCount += range.updates.length;
      } catch (err) {
        console.error(`❌ [范围更新失败] 行${range.startRow}-${range.endRow}, 列${range.startCol}-${range.endCol}: ${err.message}`);
        errorCount += range.updates.length;
      }
    }
    
    console.log(`📝 [注释更新完成] 成功: ${successCount}个, 失败: ${errorCount}个`);
    return { successCells, errorCount }; // 修改返回值结构
    
  } catch (error) {
    console.error(`❌ [注释更新失败]: ${error.message}`);
    return { successCells: [], errorCount: noteUpdates.length };
  }
}

/**
 * Fallback method for batch updates when optimization fails
 * @param {Sheet} sheet - The sheet to update
 * @param {Array} updates - Array of update objects
 * @returns {Object} Update result using individual operations
 */
function batchUpdateCellsFallback(sheet, updates) {
  console.log(`⚠️ [回退更新] 使用传统方法更新: ${sheet.getName()}`);
  
  let successCount = 0;
  let errorCount = 0;
  
  for (const update of updates) {
    try {
      const range = sheet.getRange(update.row, update.col);
      
      if (update.background !== undefined) {
        range.setBackground(update.background);
      }
      if (update.note !== undefined) {
        range.setNote(update.note);
      }
      
      successCount++;
    } catch (error) {
      console.error(`❌ [单个更新失败] 行${update.row}列${update.col}: ${error.message}`);
      errorCount++;
    }
  }
  
  return {
    success: true,
    updatedCount: successCount,
    errorCount,
    message: `回退更新完成: 成功${successCount}个, 失败${errorCount}个`
  };
}

/**
 * Clear all caches to free memory
 */
function clearAllCaches() {
  HEADER_CACHE.clear();
  DIMENSION_CACHE.clear();
  console.log(`🧹 [缓存清理] 所有缓存已清理`);
}

// ============================================================================ 
// Public API / Entry Points (Hooks for Google Sheets)
// ============================================================================ 

/**
 * Handles the onEdit event to trigger a real-time conflict check.
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
 * Checks all edited cells in the relevant ID columns.
 */
function checkRelevantIdColumns(context, relevantIdColumns) {
  const { sheet, range } = context;
  for (const idCol of relevantIdColumns) {
    const idRange = sheet.getRange(range.getRow(), idCol, range.getNumRows(), 1);
    checkIdConflicts({
      sheet: sheet,
      range: idRange
    });
  }
}

/**
 * Finds all columns in the header row that are designated as ID columns.
 */
function findIdColumns(headerRow) {
  const idColumns = [];
  const suffix = ID_CHECKER_CONFIG.ID_COLUMN_SUFFIX;
  if (!suffix) return [];

  for (let col = 0; col < headerRow.length; col++) {
    const header = headerRow[col];
    if (header && header.toString().endsWith(suffix)) {
      idColumns.push(col + 1);
    }
  }
  return idColumns;
}

/**
 * 🚀 PHASE 1 OPTIMIZATION: Improved ID column detection with strict matching
 * This function ensures only columns that EXACTLY end with '_INT_id' are considered ID columns
 */
function findIdColumnsStrict(headerRow) {
  const idColumns = [];
  const suffix = ID_CHECKER_CONFIG.ID_COLUMN_SUFFIX;
  if (!suffix) return [];

  for (let col = 0; col < headerRow.length; col++) {
    const header = headerRow[col];
    if (header && typeof header === 'string') {
      const headerStr = header.toString().trim();
      
      // Strict matching: must end with '_INT_id' and not contain '_INT_' in the middle
      if (headerStr.endsWith(suffix)) {
        // Additional check: ensure it's not a false positive like 'A_INT_base_activity_id'
        const beforeSuffix = headerStr.slice(0, -suffix.length);
        
        // Valid ID column names should not contain '_INT_' before the suffix
        // Examples of valid names: 'user_INT_id', 'product_INT_id', 'order_INT_id'
        // Examples of invalid names: 'A_INT_base_activity_id', 'base_INT_activity_id'
        if (!beforeSuffix.includes('_INT_')) {
          idColumns.push(col + 1);
          console.log(`✅ [ID列检测] 发现有效ID列: "${headerStr}" (列${col + 1})`);
        } else {
          console.log(`⚠️ [ID列检测] 跳过假阳性: "${headerStr}" (列${col + 1}) - 包含中间_INT_`);
        }
      }
    }
  }
  
  console.log(`🎯 [ID列检测完成] 共发现 ${idColumns.length} 个有效ID列`);
  return idColumns;
}

/**
 * 🚀 PHASE 1 OPTIMIZATION: Helper function to validate if a column header is a valid ID column
 * @param {string} headerValue - The column header value to check
 * @returns {boolean} True if it's a valid ID column
 */
function isValidIdColumn(headerValue) {
  if (!headerValue || typeof headerValue !== 'string') return false;
  
  const headerStr = headerValue.trim();
  const suffix = ID_CHECKER_CONFIG.ID_COLUMN_SUFFIX;
  
  if (!suffix || !headerStr.endsWith(suffix)) return false;
  
  // Additional check: ensure it's not a false positive like 'A_INT_base_activity_id'
  const beforeSuffix = headerStr.slice(0, -suffix.length);
  
  // Valid ID column names should not contain '_INT_' before the suffix
  // Examples of valid names: 'user_INT_id', 'product_INT_id', 'order_INT_id'
  // Examples of invalid names: 'A_INT_base_activity_id', 'base_INT_activity_id'
  return !beforeSuffix.includes('_INT_');
}

/**
 * Gets the column numbers that were edited in a given range.
 */
function getEditedColumns(range) {
  const startCol = range.getColumn();
  const endCol = range.getLastColumn();
  const columns = [];
  for (let i = startCol; i <= endCol; i++) {
    columns.push(i);
  }
  return columns;
}


/**
 * Main entry point for real-time conflict checking of a single edited cell.
 * @param {Object} editedCell - The cell being edited {sheet, range}.
 */
function checkIdConflicts(editedCell) {
  console.log(`🚀 [实时检查] 开始处理编辑单元格`);
  const { sheet, range } = editedCell;
  const value = range.getValue();

  console.log(`📍 [编辑信息] 表: "${sheet.getName()}", 位置: 第${range.getRow()}行第${range.getColumn()}列, 值: "${value}"`);

  // Clear previous markings first
  range.setBackground(null);
  NoteManager.removeMarkFromCell(range, NOTE_CONSTANTS.TYPES.CONFLICT);

  if (!value || !value.toString().trim()) {
    console.log(`🧹 [清除标记] 值为空，不执行检查`);
    return;
  }

  const headerValue = sheet.getRange(1, range.getColumn()).getValue();
  
  // Only check if it's a valid ID column
  if (!headerValue || !isValidIdColumn(headerValue.toString())) {
    return;
  }

  try {
    const conflicts = unifiedCheckSingleIdConflict({
      value,
      sheet: sheet.getName(),
      row: range.getRow(),
      column: range.getColumn(),
      header: headerValue
    });

    if (conflicts.length > 0) {
      console.log(`🚨 [发现冲突] ${sheet.getName()} 第${range.getRow()}行与 ${conflicts.length} 处冲突`);
      range.setBackground(ID_CHECKER_CONFIG.COLORS.CONFLICT);
      const conflictLocations = conflicts.map(loc => `${loc.sheet} 第${loc.row}行`).join('\n');
      const userNote = `在以下位置重复:\n${conflictLocations}`;
      const currentNote = range.getNote();
      const updatedNote = NoteManager.addSystemNote(currentNote, NOTE_CONSTANTS.TYPES.CONFLICT, userNote);
      range.setNote(updatedNote);
    } else {
      console.log(`✅ [无冲突]`);
    }
  } catch (e) {
    console.error(`💥 [实时检查异常] ${e.message}`);
    SpreadsheetApp.getActiveSpreadsheet().toast(`ID冲突检查时发生错误: ${e.message}`);
  }
}

/**
 * Main entry point for manually validating and clearing all conflict marks on the active sheet.
 */
function validateAndClearConflictMarks() {
  console.log(`🚀 [批量清理] 开始验证当前表格所有冲突标记`);
  const startTime = new Date().getTime();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentSheet = ss.getActiveSheet();
  const sheetName = currentSheet.getName();

  const stats = { totalCells: 0, conflictCells: 0, clearedConflicts: 0, validConflicts: 0 };

  try {
    const batchData = readSheetData(currentSheet);
    stats.totalCells = batchData.lastRow * batchData.lastColumn;

    if (stats.totalCells === 0) {
      SpreadsheetApp.getActiveSpreadsheet().toast('当前工作表没有数据可供检查。');
      return;
    }

    // 2. Detect cells with conflict notes (in-memory)
    const conflictCells = detectConflictCells(batchData);
    stats.conflictCells = conflictCells.length;
    console.log(`🎯 [冲突统计] 发现 ${conflictCells.length} 个冲突标记单元格`);

    if (conflictCells.length > 0) {
      // 3. Unified validation for all detected cells
      const validationResults = unifiedValidateCells(conflictCells, sheetName);
      
      // 4. Use optimized batch updates
      const updateStats = updateCellsBatch(validationResults, currentSheet);
      stats.clearedConflicts = updateStats.clearedCount;
      stats.validConflicts = updateStats.validCount;
    }

    const duration = new Date().getTime() - startTime;
    const results = { summary: { ...stats, duration, success: true }, sheets: [] };
    
    // 🚀 PHASE 1 OPTIMIZATION: Clear caches after operation
    clearAllCaches();
    
    showCleanupResults(results);
    
    console.log(`🎉 [验证完成] 总耗时: ${duration}ms`);

  } catch (error) {
    console.error(`💥 [清理异常] ${error.stack}`);
    SpreadsheetApp.getActiveSpreadsheet().toast(`冲突标记清理失败: ${error.message}`);
    // Clear caches even on error
    clearAllCaches();
  }
}


// ============================================================================ 
// UNIFIED CONFLICT VALIDATION PIPELINE
// ============================================================================ 

/**
 * Builds a cache of ID column data for a specific column name across all sheets.
 * @param {Spreadsheet} ss The spreadsheet object.
 * @param {string} columnName The header of the ID column to cache.
 * @returns {Map<string, Object>} Map where key is sheet name, value is {headers, idColumn, ...}.
 */
function buildSheetIdDataCache(ss, columnName) {
  const sheetIdDataCache = new Map();
  const sheets = ss.getSheets();

  for (const sheet of sheets) {
    const sheetName = sheet.getName();
    
    // 🚀 PHASE 1 OPTIMIZATION: 跳过预览表和临时表
    if (isPreviewOrTemporarySheet(sheetName)) {
      console.log(`⏭️ [跳过预览表] ${sheetName}`);
      continue;
    }
    
    try {
      const lastCol = sheet.getLastColumn();
      const lastRow = sheet.getLastRow();
      if (lastRow <= 1 || lastCol === 0) continue;

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
      }
    } catch (error) {
      console.warn(`⚠️ [缓存警告] 无法缓存表格 ${sheetName} 的ID列数据: ${error.message}`);
    }
  }
  return sheetIdDataCache;
}

/**
 * ID column cache builder using new caching system
 * @param {Spreadsheet} ss The spreadsheet object.
 * @param {string} columnName The header of the ID column to cache.
 * @returns {Map<string, Object>} Map where key is sheet name, value is {headers, idColumn, ...}.
 */
function buildSheetIdDataCacheOptimized(ss, columnName) {
  const sheetIdDataCache = new Map();
  const sheets = ss.getSheets();
  
  console.log(`🚀 [优化缓存] 开始构建ID列缓存: ${columnName}`);

  for (const sheet of sheets) {
    const sheetName = sheet.getName();
    
    // 🚀 PHASE 1 OPTIMIZATION: 跳过预览表和临时表
    if (isPreviewOrTemporarySheet(sheetName)) {
      console.log(`⏭️ [跳过预览表] ${sheetName}`);
      continue;
    }
    
    try {
      // Use cached dimensions if available
      let dimensions;
      if (DIMENSION_CACHE.has(sheetName)) {
        dimensions = DIMENSION_CACHE.get(sheetName);
      } else {
        dimensions = getSheetDimensionsCached(sheet);
      }
      
      const { lastRow, lastCol } = dimensions;
      if (lastRow <= 1 || lastCol === 0) continue;

      // Use cached headers if available
      let headers;
      if (HEADER_CACHE.has(sheetName)) {
        headers = HEADER_CACHE.get(sheetName);
      } else {
        headers = getSheetHeadersCached(sheet);
      }
      
      // Use strict ID column detection to avoid false positives
      const columnIndex = headers.findIndex(header => header && isValidIdColumn(header.toString()));

      if (columnIndex !== -1) {
        // Read only the ID column data instead of the entire sheet
        const idColumn = sheet.getRange(2, columnIndex + 1, lastRow - 1, 1).getValues().flat();
        sheetIdDataCache.set(sheetName, {
          headers,
          idColumn,
          columnIndex: columnIndex + 1,
          lastRow,
          lastCol
        });
        
        console.log(`✅ [缓存成功] ${sheetName}: ID列位置${columnIndex + 1}, 数据行数${idColumn.length}`);
      }
    } catch (error) {
      console.warn(`⚠️ [优化缓存警告] 无法缓存表格 ${sheetName} 的ID列数据: ${error.message}`);
    }
  }
  
  console.log(`🎯 [缓存完成] 成功缓存 ${sheetIdDataCache.size} 个表格的ID列数据`);
  return sheetIdDataCache;
}

/**
 * Finds all potential conflicts for a given set of tasks using the pre-built cache.
 * @param {Array<Object>} tasks The validation tasks.
 * @param {Map} columnDataCache The cache of all required ID columns.
 * @param {string} currentSheetName The name of the sheet where the check is initiated.
 * @returns {Array<Object>} A list of all potential conflict locations.
 */
function collectAllPotentialConflicts(tasks, columnDataCache, currentSheetName) {
  const allPotentialConflicts = new Map(); // Use Map to auto-deduplicate conflicts

  for (const task of tasks) {
    const { value, header } = task;
    if (!value || !value.toString().trim()) continue;

    const cache = columnDataCache.get(header);
    if (!cache) continue;

    for (const [sheetName, data] of cache.entries()) {
      // Do not check against itself in the same sheet
      if (sheetName === currentSheetName) continue;

      data.idColumn.forEach((id, index) => {
        if (id && id.toString().trim() && id.toString() === value.toString()) {
          const conflict = {
            sheet: sheetName,
            row: index + 2, // +2 because idColumn is 0-indexed and starts from row 2
            column: data.columnIndex
          };
          const key = `${conflict.sheet}_${conflict.row}`;
          if (!allPotentialConflicts.has(key)) {
            allPotentialConflicts.set(key, conflict);
          }
        }
      });
    }
  }
  return Array.from(allPotentialConflicts.values());
}

/**
 * Reads all specified row data across multiple sheets in a highly optimized way.
 * @param {Spreadsheet} ss The spreadsheet object.
 * @param {Array<Object>} rowsToRead A list of {sheet, row} objects.
 * @returns {Map<string, Array>} Map where key is "sheetName_rowNum" and value is the row data array.
 */
function batchReadRowsData(ss, rowsToRead) {
  const allRowsData = new Map();
  const groupedBySheet = rowsToRead.reduce((acc, { sheet, row }) => {
    if (!acc[sheet]) acc[sheet] = [];
    acc[sheet].push(row);
    return acc;
  }, {});

  for (const sheetName in groupedBySheet) {
    try {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) continue;
      const lastCol = sheet.getLastColumn();
      if (lastCol === 0) continue;

      const rows = groupedBySheet[sheetName].sort((a, b) => a - b);
      if (rows.length === 0) continue;

      // Read ranges in chunks to be efficient
      let startRow = rows[0];
      let endRow = rows[0];

      for (let i = 1; i < rows.length; i++) {
        // If the next row is not contiguous, read the current chunk
        if (rows[i] > endRow + 1) {
          const range = sheet.getRange(startRow, 1, endRow - startRow + 1, lastCol);
          const values = range.getValues();
          values.forEach((rowData, index) => {
            allRowsData.set(`${sheetName}_${startRow + index}`, rowData);
          });
          startRow = rows[i];
        }
        endRow = rows[i];
      }
      // Read the last chunk
      const range = sheet.getRange(startRow, 1, endRow - startRow + 1, lastCol);
      const values = range.getValues();
      values.forEach((rowData, index) => {
        allRowsData.set(`${sheetName}_${startRow + index}`, rowData);
      });
    } catch (e) {
      console.error(`Error batch reading rows from ${sheetName}: ${e.message}`);
    }
  }
  return allRowsData;
}

/**
 * The main data pre-loading engine. It orchestrates caching and batch reading.
 * @param {Array<Object>} tasks A list of validation tasks.
 * @param {Spreadsheet} ss The spreadsheet object.
 * @returns {Object} A context object with all pre-loaded data.
 */
function preloadValidationData(tasks, ss) {
  const currentSheetName = ss.getActiveSheet().getName();
  const requiredHeaders = [...new Set(tasks.map(task => task.header).filter(h => h))];

  // 1. Build cache for all required ID columns
  const columnDataCache = new Map();
  for (const header of requiredHeaders) {
    columnDataCache.set(header, buildSheetIdDataCache(ss, header));
  }

  // 2. Find all potential conflicts
  const allPotentialConflicts = collectAllPotentialConflicts(tasks, columnDataCache, currentSheetName);

  // 3. Collect all unique rows that need to be read
  const rowsToReadMap = new Map();
  const addTaskRow = (task) => {
    const key = `${task.sheet}_${task.row}`;
    if (!rowsToReadMap.has(key)) rowsToReadMap.set(key, { sheet: task.sheet, row: task.row });
  };
  tasks.forEach(addTaskRow);
  allPotentialConflicts.forEach(conflict => {
     const key = `${conflict.sheet}_${conflict.row}`;
     if (!rowsToReadMap.has(key)) rowsToReadMap.set(key, { sheet: conflict.sheet, row: conflict.row });
  });

  // 4. Batch read all row data
  const allRowsData = batchReadRowsData(ss, Array.from(rowsToReadMap.values()));

  return { columnDataCache, allRowsData, currentSheetName };
}

// STAGE 2: IN-MEMORY VALIDATION

/**
 * Compares two rows of data for equality based on common headers.
 * @returns {boolean} True if rows are considered identical.
 */
function compareRowsData(row1, row2, headers1, headers2) {
  if (!row1 || !row2 || !headers1 || !headers2) return false;

  const map1 = new Map(headers1.map((h, i) => [h, row1[i] ? row1[i].toString().trim() : '']));
  const map2 = new Map(headers2.map((h, i) => [h, row2[i] ? row2[i].toString().trim() : '']));

  const commonKeys = [...map1.keys()].filter(key => map2.has(key));
  if (commonKeys.length === 0) return false;

  return commonKeys.every(key => map1.get(key) === map2.get(key));
}

/**
 * Performs validation for a single task using only the pre-loaded context.
 * @returns {Array<Object>} A list of verified conflict locations.
 */
function performInMemoryValidation(task, context) {
  const { columnDataCache, allRowsData } = context;
  const { value, row, header, sheet } = task; // `sheet` is the current sheet name

  const cache = columnDataCache.get(header);
  if (!cache) return [];
  
  const currentRowData = allRowsData.get(`${sheet}_${row}`);
  if (!currentRowData) {
    console.warn(`Could not find row data for ${sheet}_${row} in cache.`);
    return [];
  }

  // --- Step 1: Check for conflicts within the CURRENT sheet (ID match only) ---
  const currentSheetCache = cache.get(sheet);
  if (currentSheetCache) {
    const currentSheetConflicts = [];
    for (let i = 0; i < currentSheetCache.idColumn.length; i++) {
      const id = currentSheetCache.idColumn[i];
      const actualRow = i + 2; // +2 because array is 0-indexed and data starts at row 2

      // For the current sheet, any ID match on a different row is a conflict, regardless of data.
      if (actualRow !== row && id && id.toString().trim() === value.toString()) {
        currentSheetConflicts.push({ sheet, row: actualRow, column: currentSheetCache.columnIndex });
      }
    }
    // If any conflicts are found in the current sheet, return them immediately.
    if (currentSheetConflicts.length > 0) {
      return currentSheetConflicts;
    }
  }

  // --- Step 2: If no conflicts in current sheet, check OTHER sheets (ID + Data Inconsistency) ---
  const otherSheetConflicts = [];
  const currentHeaders = currentSheetCache ? currentSheetCache.headers : [];

  for (const [sheetName, sheetData] of cache.entries()) {
    if (sheetName === sheet) continue; // Skip current sheet

    for (let i = 0; i < sheetData.idColumn.length; i++) {
      const id = sheetData.idColumn[i];
      if (id && id.toString().trim() === value.toString()) {
        const conflictRow = i + 2;
        const conflictRowData = allRowsData.get(`${sheetName}_${conflictRow}`);
        const conflictHeaders = sheetData.headers;
        
        // For other sheets, check for data inconsistency before flagging a conflict.
        if (conflictRowData && !compareRowsData(currentRowData, conflictRowData, currentHeaders, conflictHeaders)) {
          otherSheetConflicts.push({ sheet: sheetName, row: conflictRow, column: sheetData.columnIndex });
        }
      }
    }
  }

  return otherSheetConflicts;
}

// STAGE 3: UNIFIED PUBLIC-FACING VALIDATORS

/**
 * Unified entry point for a single real-time conflict check.
 * @param {Object} task - The validation task {value, sheet, row, column, header}.
 * @returns {Array<Object>} A list of verified conflicts.
 */
function unifiedCheckSingleIdConflict(task) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const context = preloadValidationData([task], ss);
  return performInMemoryValidation(task, context);
}

/**
 * Unified entry point to validate a list of cells for the manual cleanup process.
 * @param {Array<Object>} cellsToValidate List of cell info objects.
 * @param {string} sheetName The name of the sheet being validated.
 * @returns {Array<Object>} A list of validation result objects.
 */
function unifiedValidateCells(cellsToValidate, sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tasks = cellsToValidate.map(cell => ({ ...cell, sheet: sheetName }));
  if (tasks.length === 0) return [];

  const context = preloadValidationData(tasks, ss);

  const results = [];
  for (const task of tasks) {
    const conflicts = performInMemoryValidation(task, context);
    results.push({
      ...task,
      stillHasConflict: conflicts.length > 0,
      conflicts
    });
  }
  return results;
}

// ============================================================================ 
// BATCH UPDATE & CLEANUP HELPERS
// ============================================================================ 

/**
 * Reads all necessary data from a sheet for the cleanup process.
 */
function readSheetDataBatch(sheet, lastRow, lastColumn) {
  const range = sheet.getRange(1, 1, lastRow, lastColumn);
  return {
    backgrounds: range.getBackgrounds(),
    values: range.getValues(),
    notes: range.getNotes(),
    headers: sheet.getRange(1, 1, 1, lastColumn).getValues()[0]
  };
}

/**
 * Detects cells with conflict notes from the pre-read data.
 */
function detectConflictCells(batchData) {
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
 * batch update using new batch update functions
 * @returns {Object} Statistics of the update operation.
 */
function updateCellsBatch(validationResults, currentSheet) {
  const { cellsToClear, cellsToUpdate } = categorizeCells(validationResults);
  
  console.log(`🚀 [批量更新] 开始处理: 清除${cellsToClear.length}个, 更新${cellsToUpdate.length}个`);

  // 使用优化的批量操作
  const clearedCount = clearConflictMarksBatch(cellsToClear, currentSheet);
  const validCount = updateConflictNotesBatchOptimized(cellsToUpdate, currentSheet);

  return { 
    clearedCount, 
    validCount: cellsToUpdate.length  // 直接使用输入的数量，而不是 batchUpdateCells 的统计
  };
}

/**
 * Categorizes cells into those that need clearing vs. those that need updating.
 */
function categorizeCells(validationResults) {
  const cellsToClear = [];
  const cellsToUpdate = [];
  validationResults.forEach(result => {
    if (result.stillHasConflict) {
      cellsToUpdate.push(result);
    } else {
      cellsToClear.push(result);
    }
  });
  return { cellsToClear, cellsToUpdate };
}

/**
 * batch clear using new batch update system
 * @returns {number} The number of cells successfully cleared.
 */
function clearConflictMarksBatch(cellsToClear, currentSheet) {
  if (cellsToClear.length === 0) return 0;
  
  console.log(`🧹 [批量清除] 开始清除 ${cellsToClear.length} 个过期冲突标记`);
  
  // Prepare updates for the optimized batch system
  const updates = cellsToClear.map(cell => ({
    row: cell.row,
    col: cell.col,
    background: null,
    note: '' // Clear note
  }));
  
  // Use optimized batch update
  const result = batchUpdateCells(currentSheet, updates);
  
  if (result.success) {
    console.log(`✅ [批量清除完成] 成功清除 ${result.updatedCount} 个单元格`);
    return result.updatedCount;
  } else {
    console.error(`❌ [批量清除失败] 回退到传统方法`);
    return clearConflictMarksBatch(cellsToClear, currentSheet);
  }
}

/**
 * Updates the notes for a list of cells that still have valid conflicts.
 * @returns {number} The number of cells successfully updated.
 */
function updateConflictNotesBatch(cellsToUpdate, currentSheet) {
  if (cellsToUpdate.length === 0) return 0;
  console.log(`📝 [批量更新] 开始更新 ${cellsToUpdate.length} 个有效冲突的注释和背景`);
  let updatedCount = 0;
  // setNotes is tricky for non-contiguous ranges. Looping here is often the most robust approach.
  for (const cell of cellsToUpdate) {
    try {
      const cellRange = currentSheet.getRange(cell.row, cell.col);
      const conflictLocations = cell.conflicts.map(loc => `${loc.sheet} 第${loc.row}行`).join('\n');
      const userNote = `在以下位置重复:\n${conflictLocations}`;
      const currentNote = cellRange.getNote();
      const updatedNote = NoteManager.addSystemNote(currentNote, NOTE_CONSTANTS.TYPES.CONFLICT, userNote);
      cellRange.setNote(updatedNote);
      cellRange.setBackground(ID_CHECKER_CONFIG.COLORS.CONFLICT);
      updatedCount++;
    } catch (e) {
      console.error(`❌ [更新注释失败] ${currentSheet.getName()} 第${cell.row}行: ${e.message}`);
    }
  }
  return updatedCount;
}

/**
 * 🚀 PHASE 1 OPTIMIZATION: Optimized batch note update using new batch update system
 * @returns {number} The number of cells successfully updated.
 */
function updateConflictNotesBatchOptimized(cellsToUpdate, currentSheet) {
  if (cellsToUpdate.length === 0) return 0;
  
  console.log(`📝 [批量更新] 开始更新 ${cellsToUpdate.length} 个有效冲突的注释和背景`);
  
  // Prepare updates for the optimized batch system
  const updates = cellsToUpdate.map(cell => {
    const conflictLocations = cell.conflicts.map(loc => `${loc.sheet} 第${loc.row}行`).join('\n');
    const userNote = `在以下位置重复:\n${conflictLocations}`;
    const currentNote = currentSheet.getRange(cell.row, cell.col).getNote();
    const updatedNote = NoteManager.addSystemNote(currentNote, NOTE_CONSTANTS.TYPES.CONFLICT, userNote);
    
    return {
      row: cell.row,
      col: cell.col,
      background: ID_CHECKER_CONFIG.COLORS.CONFLICT,
      note: updatedNote
    };
  });
  
  // Use optimized batch update
  const result = batchUpdateCells(currentSheet, updates);
  
  if (result.success) {
    console.log(`✅ [优化批量更新完成] 成功更新 ${result.updatedCount} 个单元格`);
    return result.updatedCount;
  } else {
    console.error(`❌ [优化批量更新失败] 回退到传统方法`);
    return updateConflictNotesBatch(cellsToUpdate, currentSheet);
  }
}

// ============================================================================ 
// UI / UTILITY
// ============================================================================ 

/**
 * Displays the results of the cleanup operation in a toast message.
 */
function showCleanupResults(results) {
  const { summary } = results;
  if (!summary.success) {
    SpreadsheetApp.getActiveSpreadsheet().toast(`清理失败: ${summary.error}`, '清理错误', 10);
    return;
  }

  let message;
  if (summary.conflictCells === 0) {
    message = '未发现任何冲突标记，表格状态良好！';
  } else if (summary.clearedConflicts > 0) {
    message = `成功清理了 ${summary.clearedConflicts} 个过期的冲突标记！保留 ${summary.validConflicts} 个有效冲突。`;
  } else {
    message = `所有 ${summary.validConflicts} 个冲突标记都是有效的，无需清理。`;
  }
  
  SpreadsheetApp.getActiveSpreadsheet().toast(message, '清理完成', 8);
  console.log(`📋 [清理报告] ${message}`);
}

/**
 * 检查表格是否为预览表或临时表
 * @param {string} sheetName 表格名称
 * @returns {boolean} 如果是预览表或临时表则返回true
 */
function isPreviewOrTemporarySheet(sheetName) {
  // 合并预览表模式：源表 -> 目标表 合并预览
  if (sheetName.includes(' -> ') && sheetName.includes(' 合并预览')) {
    return true;
  }
  
  // 对比预览表模式：表1 vs 表2 比较结果
  if (sheetName.includes(' vs ') && sheetName.includes(' 比较结果')) {
    return true;
  }
  
  // 其他可能的预览表模式
  if (sheetName.includes('预览') || sheetName.includes('preview') || sheetName.includes('temp')) {
    return true;
  }
  
  return false;
}