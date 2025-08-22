/**
 * Google Sheets ID Conflict Checker - Refactored Version
 * Provides efficient ID conflict detection, validation, and cleanup.
 * This version is refactored to reduce redundancy and improve maintainability
 * by using a unified data processing pipeline.
 *
 * Core Principles:
 * 1. Unified Pipeline: All checks (real-time, batch) use the same core logic.
 * 2. Batch Everything: All reads and writes to the spreadsheet are batched to minimize API calls.
 * 3. Separation of Concerns: Logic is split into Data Loading, In-Memory Validation, and Sheet Updating.
 */

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
  
  // Only check if it\'s a valid ID column
  if (!headerValue || !headerValue.toString().endsWith(ID_CHECKER_CONFIG.ID_COLUMN_SUFFIX)) {
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
    const lastRow = currentSheet.getLastRow();
    const lastColumn = currentSheet.getLastColumn();

    if (lastRow <= 1 || lastColumn === 0) {
      SpreadsheetApp.getActiveSpreadsheet().toast('当前工作表没有数据可供检查。');
      return;
    }

    // 1. Batch read all sheet data at once
    const batchData = readSheetDataBatch(currentSheet, lastRow, lastColumn);
    stats.totalCells = lastRow * lastColumn;

    // 2. Detect cells with conflict notes (in-memory)
    const conflictCells = detectConflictCells(batchData);
    stats.conflictCells = conflictCells.length;
    console.log(`🎯 [冲突统计] 发现 ${conflictCells.length} 个冲突标记单元格`);

    if (conflictCells.length > 0) {
      // 3. Unified validation for all detected cells
      const validationResults = unifiedValidateCells(conflictCells, sheetName);
      
      // 4. Batch update cells based on validation results
      const updateStats = updateCellsBatch(validationResults, currentSheet);
      stats.clearedConflicts = updateStats.clearedCount;
      stats.validConflicts = updateStats.validCount;
    }

    const duration = new Date().getTime() - startTime;
    const results = { summary: { ...stats, duration, success: true }, sheets: [] }; // Simplified results
    showCleanupResults(results);

  } catch (error) {
    console.error(`💥 [清理异常] ${error.stack}`);
    SpreadsheetApp.getActiveSpreadsheet().toast(`冲突标记清理失败: ${error.message}`);
  }
}


// ============================================================================ 
// UNIFIED CONFLICT VALIDATION PIPELINE
// ============================================================================ 

// STAGE 1: DATA PREPARATION

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
 * Orchestrates the batch update of cells after validation.
 * @returns {Object} Statistics of the update operation.
 */
function updateCellsBatch(validationResults, currentSheet) {
  const { cellsToClear, cellsToUpdate } = categorizeCells(validationResults);

  const clearedCount = clearConflictMarksBatch(cellsToClear, currentSheet);
  const validCount = updateConflictNotesBatch(cellsToUpdate, currentSheet);

  return { clearedCount, validCount };
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
 * Clears conflict markings (background and note) from a list of cells using batch operations.
 * @returns {number} The number of cells successfully cleared.
 */
function clearConflictMarksBatch(cellsToClear, currentSheet) {
  if (cellsToClear.length === 0) return 0;
  console.log(`🧹 [批量清除] 开始清除 ${cellsToClear.length} 个过期冲突标记`);
  try {
    const a1Notations = cellsToClear.map(cell => currentSheet.getRange(cell.row, cell.col).getA1Notation());
    const rangeList = currentSheet.getRangeList(a1Notations);
    rangeList.setBackground(null);
    rangeList.clearNote();
    return cellsToClear.length;
  } catch (e) {
    console.error(`❌ [批量清除失败]: ${e.message}. 回退到逐个清除。`);
    let count = 0;
    for(const cell of cellsToClear) {
        try {
            const range = currentSheet.getRange(cell.row, cell.col);
            range.setBackground(null);
            range.clearNote();
            count++;
        } catch (err) { /* ignore single error */ }
    }
    return count;
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