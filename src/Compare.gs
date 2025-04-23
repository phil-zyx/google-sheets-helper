// 显示配置对话框
function showCompareDialog() {
  // 每次打开对话框时清除缓存，确保获取最新数据
  var cache = CacheService.getScriptCache();
  cache.remove('sheet_info');
  
  var html = HtmlService.createHtmlOutputFromFile('CompareDialog')
    .setWidth(400)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, '表格比较配置');
}

// 执行比较操作 - 优化数据处理
function compareSheets(config) {
  if (!config || !config.sheet1 || !config.sheet2) {
    return {
      success: false,
      message: "配置参数无效"
    };
  }

  // 检查当前表格是否为对比结果表
  var currentSheet = SpreadsheetApp.getActiveSheet();
  if (currentSheet.getName().includes(" vs ") && currentSheet.getName().endsWith("比较结果")) {
    return {
      success: false,
      message: "对比结果表不能作为对比的源表格，请切换到其他表格后再试"
    };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = ss.getSheetByName(config.sheet1);
  var sheet2 = ss.getSheetByName(config.sheet2);
  
  if (!sheet1 || !sheet2) {
    return {
      success: false,
      message: "未找到指定的表格，请检查表格名称"
    };
  }

  try {
    // 创建比较表格
    var previewName = `${config.sheet1} vs ${config.sheet2} 比较结果`;
    var previewSheet = createCompareSheet(previewName);

    // 批量获取数据以减少API调用
    var range1 = sheet1.getDataRange();
    var range2 = sheet2.getDataRange();
    
    var data1 = range1.getValues();
    var data2 = range2.getValues();
    
    // 获取表头数据
    var headers1 = data1[0];
    var headers2 = data2[0];
    
    // 使用对象来存储表头映射
    var headerMap = {};
    var unmatchedHeaders1 = [];
    var unmatchedHeaders2 = [];
    
    // 记录表头1中的列索引
    headers1.forEach((header, index) => {
      headerMap[header] = {sheet1Index: index, sheet2Index: -1};
    });
    
    // 查找表头2中对应的列索引
    headers2.forEach((header, index) => {
      if (headerMap[header]) {
        headerMap[header].sheet2Index = index;
      } else {
        unmatchedHeaders2.push({header: header, index: index});
      }
    });
    
    // 找出表头1中未匹配的列
    headers1.forEach((header, index) => {
      if (headerMap[header].sheet2Index === -1) {
        unmatchedHeaders1.push({header: header, index: index});
      }
    });

    // 找到ID列（使用常量定义的后缀）
    var idCol1 = headers1.findIndex(header => 
      header.toString().endsWith(ID_CHECKER_CONFIG.ID_COLUMN_SUFFIX));
    var idCol2 = headers2.findIndex(header => 
      header.toString().endsWith(ID_CHECKER_CONFIG.ID_COLUMN_SUFFIX));
    
    if (idCol1 === -1 || idCol2 === -1) {
      return {
        success: false,
        message: `未找到ID列（以'${ID_CHECKER_CONFIG.ID_COLUMN_SUFFIX}'结尾的列）`
      };
    }

    // 准备预览表数据
    var previewData = [headers1];
    var previewColors = [new Array(headers1.length).fill(null)];
    var previewNotes = [new Array(headers1.length).fill('')];

    // 使用Map存储表2的数据，以ID为键
    var data2Map = new Map();
    for (var i = 1; i < data2.length; i++) {
      var id = data2[i][idCol2].toString();
      if (!data2Map.has(id)) {
        data2Map.set(id, []);
      }
      data2Map.get(id).push({
        rowIndex: i,
        data: data2[i]
      });
    }

    var differences = {
      total: 0,
      modified: 0,
      added: 0,
      removed: 0,
      duplicate: 0,
      headerDiff: unmatchedHeaders1.length + unmatchedHeaders2.length
    };

    // 处理表1的数据行
    var processedIds = new Set();
    for (var i = 1; i < data1.length; i++) {
      var id = data1[i][idCol1].toString();
      var hasChanges = false;
      var rowColors = new Array(headers1.length).fill(null);
      var rowNotes = new Array(headers1.length).fill('');
      var rowData = [...data1[i]];

      var matchingRows = data2Map.get(id) || [];
      processedIds.add(id);

      if (matchingRows.length === 0) {
        // ID在表2中不存在，标记为新增行
        rowColors.fill(COMPARE_CONSTANTS.COLORS.ADDED);
        // 移除注释，只使用颜色标记
        hasChanges = true;
        differences.added++;
      } else if (matchingRows.length > 1) {
        // ID在表2中有重复
        rowColors.fill(COMPARE_CONSTANTS.COLORS.MODIFIED);
        rowNotes = rowNotes.map(note => 
          NoteManager.addSystemNote(note, NOTE_CONSTANTS.TYPES.VERSION, 
            `在对比表格中发现 ${matchingRows.length} 条重复记录`)
        );
        hasChanges = true;
        differences.duplicate++;
      } else {
        // 比较每个单元格
        for (var j = 0; j < headers1.length; j++) {
          var sheet2Col = headerMap[headers1[j]].sheet2Index;
          if (sheet2Col === -1) {
            rowColors[j] = COMPARE_CONSTANTS.COLORS.ADDED;
            rowNotes[j] = NoteManager.addSystemNote('', NOTE_CONSTANTS.TYPES.VERSION, 
              "此列在对比表格中不存在");
            hasChanges = true;
            differences.added++;
          } else {
            var value2 = matchingRows[0].data[sheet2Col];
            if (data1[i][j] !== value2) {
              rowColors[j] = COMPARE_CONSTANTS.COLORS.MODIFIED;
              rowNotes[j] = NoteManager.addSystemNote('', NOTE_CONSTANTS.TYPES.VERSION,
                `当前表格: ${data1[i][j]}\n对比表格: ${value2}`);
              hasChanges = true;
              differences.modified++;
            }
          }
        }
      }

      if (hasChanges) {
        previewData.push(rowData);
        previewColors.push(rowColors);
        previewNotes.push(rowNotes);
        differences.total++;
      }
    }

    // 检查表2中存在而表1中不存在的ID
    for (let [id, rows] of data2Map) {
      if (!processedIds.has(id)) {
        // 对于每个未处理的ID，添加一行到预览表
        var rowData = new Array(headers1.length).fill('');
        var rowColors = new Array(headers1.length).fill(COMPARE_CONSTANTS.COLORS.REMOVED);
        var rowNotes = new Array(headers1.length).fill('').map(note => 
          NoteManager.addSystemNote(note, NOTE_CONSTANTS.TYPES.VERSION, '此ID在基准表中不存在')
        );

        // 填充能对应的数据
        headers1.forEach((header, index) => {
          var sheet2Col = headerMap[header].sheet2Index;
          if (sheet2Col !== -1) {
            rowData[index] = rows[0].data[sheet2Col];
          }
        });

        if (rows.length > 1) {
          rowNotes = rowNotes.map(note => 
            NoteManager.addSystemNote(
              NoteManager.removeSystemNote(note, NOTE_CONSTANTS.TYPES.VERSION),
              NOTE_CONSTANTS.TYPES.VERSION,
              `此ID在基准表中不存在\n(在对比表格中有 ${rows.length} 条重复记录)`
            )
          );
        }

        previewData.push(rowData);
        previewColors.push(rowColors);
        previewNotes.push(rowNotes);
        differences.removed++;
        differences.total++;
      }
    }

    // 更新预览表
    var previewRange = previewSheet.getRange(1, 1, previewData.length, headers1.length);
    previewRange.setValues(previewData);
    previewRange.setBackgrounds(previewColors);
    previewRange.setNotes(previewNotes);

    // 添加比较信息到A1单元格
    var compareInfo = `对比表格: ${config.sheet2}\n` +
                     `差异总数: ${differences.total}\n` +
                     `└─ 值不同: ${differences.modified}\n` +
                     `└─ 新增项: ${differences.added}\n` +
                     `└─ 删除项: ${differences.removed}\n` +
                     `└─ 重复ID: ${differences.duplicate}\n` +
                     `└─ 表头差异: ${differences.headerDiff}\n` +
                     `比较时间: ${new Date().toLocaleString()}`;
    
    // 使用 NoteManager 添加系统注释
    var a1Cell = previewSheet.getRange(1, 1);
    var currentNote = a1Cell.getNote();
    var newNote = NoteManager.addSystemNote(
      currentNote,
      NOTE_CONSTANTS.TYPES.VERSION,
      compareInfo
    );
    a1Cell.setNote(newNote);

    // 激活预览表
    previewSheet.activate();

    // 记录比较完成的日志
    LogManager.addLog(
      LOG_CONSTANTS.TYPES.COMPARE,
      config.sheet1,
      "比较完成",
      `对比表格：${config.sheet2}\n` +
      `差异总数：${differences.total}\n` +
      `└─ 值不同：${differences.modified}\n` +
      `└─ 新增项：${differences.added}\n` +
      `└─ 删除项：${differences.removed}\n` +
      `└─ 重复ID：${differences.duplicate}\n` +
      `└─ 表头差异：${differences.headerDiff}`
    );

    return {
      success: true,
      message: `比较完成！\n发现 ${differences.total} 处差异\n${differences.modified} 处值不同\n${differences.added} 处新增\n${differences.removed} 处删除\n${differences.duplicate} 处重复ID\n${differences.headerDiff} 处表头差异`,
      differences: differences
    };
  } catch (error) {
    return {
      success: false,
      message: "发生错误: " + error.toString()
    };
  }
}

/**
 * 创建比较表格
 * @param {string} previewName 预览表格名称
 * @returns {Sheet} 比较表格
 */
function createCompareSheet(previewName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var existingSheet = ss.getSheetByName(previewName);
  
  if (existingSheet) {
    ss.deleteSheet(existingSheet);
  }
  
  return ss.insertSheet(previewName);
}

// 清除所有高亮和注释
function clearAllHighlights(showConfirm = true) {
  var ui = SpreadsheetApp.getUi();
  
  if (showConfirm) {
    var response = ui.alert(
      '确认清除',
      '是否要清除当前表格中所有的比较标记？',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      return;
    }
  }
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getDataRange();
  var backgrounds = range.getBackgrounds();
  var notes = range.getNotes();
  var values = range.getValues();
  
  var newBackgrounds = [];
  var newNotes = [];
  var newValues = [];
  var rowsToKeep = [];
  
  // 检查每一行，标记需要保留的行
  for (var i = 0; i < backgrounds.length; i++) {
    var isDeletedRow = true;
    var hasHighlight = false;
    
    // 检查这一行是否是比较时新增的行（通过检查背景色和注释）
    for (var j = 0; j < backgrounds[i].length; j++) {
      var currentBg = backgrounds[i][j];
      var currentNote = notes[i][j];
      
      if (currentBg === COMPARE_CONSTANTS.COLORS.REMOVED) {
        hasHighlight = true;
      }
      
      if (currentNote && currentNote.includes("此行在基准表中不存在")) {
        hasHighlight = true;
      }
      
      // 如果这一行有任何非高亮的单元格，说明不是新增的行
      if (currentBg !== COMPARE_CONSTANTS.COLORS.REMOVED && 
          currentBg !== COMPARE_CONSTANTS.COLORS.MODIFIED && 
          currentBg !== COMPARE_CONSTANTS.COLORS.ADDED && 
          currentBg !== COMPARE_CONSTANTS.COLORS.HEADER_MODIFIED) {
        isDeletedRow = false;
      }
    }
    
    // 如果这一行不是新增的行，或者是第一行（表头），就保留它
    if (!isDeletedRow || !hasHighlight || i === 0) {
      rowsToKeep.push(i);
      
      var backgroundRow = [];
      var noteRow = [];
      
      for (var j = 0; j < backgrounds[i].length; j++) {
        var currentBg = backgrounds[i][j];
        var currentNote = notes[i][j];
        
        // 清除所有比较标记的背景色
        if (currentBg === COMPARE_CONSTANTS.COLORS.MODIFIED || 
            currentBg === COMPARE_CONSTANTS.COLORS.ADDED || 
            currentBg === COMPARE_CONSTANTS.COLORS.REMOVED || 
            currentBg === COMPARE_CONSTANTS.COLORS.HEADER_MODIFIED) {
          backgroundRow.push(null);
        } else {
          backgroundRow.push(currentBg);
        }
        
        // 清除所有比较相关的注释
        if (currentNote) {
          // 移除版本信息注释
          currentNote = NoteManager.removeSystemNote(currentNote, NOTE_CONSTANTS.TYPES.VERSION);
          noteRow.push(currentNote || '');
        } else {
          noteRow.push('');
        }
      }
      
      newBackgrounds.push(backgroundRow);
      newNotes.push(noteRow);
      newValues.push(values[i]);
    }
  }
  
  // 如果有行被删除，更新表格
  if (rowsToKeep.length < backgrounds.length) {
    var newRange = sheet.getRange(1, 1, newBackgrounds.length, backgrounds[0].length);
    newRange.setBackgrounds(newBackgrounds);
    newRange.setNotes(newNotes);
    newRange.setValues(newValues);
    
    // 删除多余的行
    if (backgrounds.length > newBackgrounds.length) {
      sheet.deleteRows(newBackgrounds.length + 1, backgrounds.length - newBackgrounds.length);
    }
  } else {
    // 如果没有行被删除，只更新背景色和注释
    range.setBackgrounds(newBackgrounds);
    range.setNotes(newNotes);
  }
}

/**
 * 基于当前表格创建新的页签
 */
function createNewSheetTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var currentSheet = ss.getActiveSheet();
  
  // 弹出对话框让用户输入新页签名称
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    '新建页签',
    '请输入新页签名称：\n(将基于当前页签 "' + currentSheet.getName() + '" 创建)',
    ui.ButtonSet.OK_CANCEL
  );

  // 处理用户输入
  if (response.getSelectedButton() == ui.Button.OK) {
    var newSheetName = response.getResponseText().trim();
    
    // 验证输入的名称
    if (newSheetName === '') {
      ui.alert('错误', '页签名称不能为空', ui.ButtonSet.OK);
      return;
    }
    
    // 检查是否已存在同名页签
    if (ss.getSheetByName(newSheetName)) {
      ui.alert('错误', '已存在同名页签："' + newSheetName + '"', ui.ButtonSet.OK);
      return;
    }
    
    try {
      // 复制当前页签
      var newSheet = currentSheet.copyTo(ss);
      newSheet.setName(newSheetName);
      
      // 将新页签移动到当前页签后面
      var sheets = ss.getSheets();
      var currentIndex = sheets.findIndex(function(sheet) {
        return sheet.getName() === currentSheet.getName();
      });
      ss.setActiveSheet(newSheet);
      ss.moveActiveSheet(currentIndex + 2);
      
      // 清除所有标记和系统注释
      clearAllMarks(false);  // 传入 false 以跳过确认对话框
      
      // 创建用户友好的创建信息
      var creationInfo = {
        '创建者': Session.getActiveUser().getEmail(),
        '创建时间': new Date().toLocaleString(),
        '来源页签': currentSheet.getName()
      };
      
      var a1Cell = newSheet.getRange("A1");
      var currentNote = a1Cell.getNote() || '';
      var newNote = NoteManager.addSystemNote(
        currentNote,
        NOTE_CONSTANTS.TYPES.SHEET_CREATION,
        Object.entries(creationInfo)
          .map(([key, value]) => `${key}：${value}`)
          .join('\n')
      );
      a1Cell.setNote(newNote);
      
      // 清除缓存以确保getSheetInfo()返回最新数据
      var cache = CacheService.getScriptCache();
      cache.remove('sheet_info');
      
      // 记录创建成功日志
      LogManager.addLog(
        LOG_CONSTANTS.TYPES.SHEET_CREATE,
        newSheetName,
        "创建页签成功",
        `来源页签：${currentSheet.getName()}`
      );
      
      ui.alert('成功', '已创建新页签："' + newSheetName + '"', ui.ButtonSet.OK);
    } catch (error) {
      ui.alert('错误', '创建页签失败：' + error.toString(), ui.ButtonSet.OK);
    }
  }
}

/**
 * 使用原生UI实现比较功能
 * 不使用HTML对话框，而是使用Google Sheets内置的UI组件
 */
function nativeCompare() {
  try {
    // 获取当前活动页签
    var activeSheet = SpreadsheetApp.getActiveSheet();
    var currentSheetName = activeSheet.getName();
    var ui = SpreadsheetApp.getUi();
    
    // 检查当前表格是否为比较结果表
    if (currentSheetName.includes(" vs ") && currentSheetName.endsWith("比较结果")) {
      ui.alert(
        '无法比较',
        '当前表格是比较结果表，不能用作比较源表格。请先切换到其他表格。',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // 获取所有表格
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets().map(sheet => sheet.getName());
    
    // 过滤掉当前表格和比较结果表
    var targetSheets = sheets.filter(name => 
      name !== currentSheetName && 
      !(name.includes(" vs ") && name.endsWith("比较结果"))
    );
    
    if (targetSheets.length === 0) {
      ui.alert(
        '无法比较',
        '没有可用的目标表格进行比较。',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // 构建选择列表
    var sheetList = targetSheets.map((name, index) => `${index + 1}. ${name}`).join('\n');
    
    // 显示目标表格选择对话框
    var response = ui.prompt(
      '选择对比表格',
      '请输入要与当前表格进行对比的表格编号:\n\n' + sheetList + '\n\n当前表格: ' + currentSheetName,
      ui.ButtonSet.OK_CANCEL
    );
    
    // 处理用户选择
    if (response.getSelectedButton() == ui.Button.OK) {
      var selection = response.getResponseText().trim();
      var index = parseInt(selection) - 1;
      
      // 验证输入
      if (isNaN(index) || index < 0 || index >= targetSheets.length) {
        ui.alert('错误', '请输入有效的表格编号 (1-' + targetSheets.length + ')', ui.ButtonSet.OK);
        return;
      }
      
      var targetSheetName = targetSheets[index];
      
      // 确认比较操作
      var confirmResponse = ui.alert(
        '确认比较操作',
        '是否比较 "' + currentSheetName + '" 和 "' + targetSheetName + '"?\n\n' + 
        '比较结果将显示在新创建的表格中。',
        ui.ButtonSet.YES_NO
      );
      
      if (confirmResponse == ui.Button.YES) {
        // 显示处理中提示
        ui.alert(
          '正在处理',
          '正在执行比较操作，这可能需要一些时间，请稍候...\n' +
          '点击"确定"后，操作将在后台继续，完成后会显示结果。',
          ui.ButtonSet.OK
        );
        
        // 执行比较
        var result = compareSheets({
          sheet1: currentSheetName,
          sheet2: targetSheetName
        });
        
        if (result.success) {
          // 激活比较结果页签
          var compareResultName = `${currentSheetName} vs ${targetSheetName} 比较结果`;
          var compareSheet = ss.getSheetByName(compareResultName);
          if (compareSheet) {
            ss.setActiveSheet(compareSheet);
          }
          
          // 显示成功信息
          ui.alert(
            '比较完成',
            result.message,
            ui.ButtonSet.OK
          );
        } else {
          // 显示错误信息
          ui.alert('错误', result.message, ui.ButtonSet.OK);
        }
      }
    }
  } catch (error) {
    console.error('执行比较操作出错:', error);
    SpreadsheetApp.getUi().alert('错误', '执行比较操作时出错: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}