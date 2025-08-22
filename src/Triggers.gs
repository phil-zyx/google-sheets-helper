/**
 * 安装时触发，设置必要的触发器
 */
function onInstall(e) {
  onOpen(e);
  createEditTrigger(false);
}

/**
 * 卸载时触发，清理所有触发器
 */
function onUninstall(e) {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const scriptId = ScriptApp.getScriptId();
    
    triggers.forEach(trigger => {
      // 只删除由当前脚本创建的触发器
      if (trigger.getScriptId() === scriptId) {
        ScriptApp.deleteTrigger(trigger);
      }
    });
  } catch (error) {
    console.error('Error cleaning up triggers:', error);
  }
}

/**
 * 设置编辑触发器
 * @param {boolean} showToast 是否显示提示，默认为true
 */
function createEditTrigger(showToast = true) {
  try {
    // First try to get authorization
    try {
      ScriptApp.getProjectTriggers();
    } catch (authError) {
      // If we get a permissions error, show a more user-friendly message
      if (showToast) {
        SpreadsheetApp.getActive().toast(
          '需要额外授权来安装触发器。请重新运行此脚本并接受权限请求。',
          '需要授权',
          10
        );
      }
      return;
    }

    // Original trigger creation logic
    const triggers = ScriptApp.getProjectTriggers();
    let hasEditTrigger = false;
    
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'onEditHandler') {
        hasEditTrigger = true;
      }
    });

    if (!hasEditTrigger) {
      const ss = SpreadsheetApp.getActive();
      ScriptApp.newTrigger('onEditHandler')
        .forSpreadsheet(ss)
        .onEdit()
        .create();
      
      if (showToast) {
        SpreadsheetApp.getActive().toast('编辑触发器已安装', '提示', 3);
      }
    }
  } catch (error) {
    console.error('Error creating edit trigger:', error);
    if (showToast) {
      SpreadsheetApp.getActive().toast('触发器安装失败: ' + error.message, '错误', 5);
    }
  }
}

/**
 * 打开文档时的触发器
 */
function onOpen(e) {
  try {
    SpreadsheetApp.getUi()
      .createAddonMenu()
      .addItem('新建页签', 'createNewSheetTab')
      .addItem('比较差异', 'nativeCompare')  // 使用原生UI的比较功能
      .addItem('合并表格', 'smartNativeMerge')  // 使用智能合并函数
      // .addItem('比较差异', 'showCompareDialog')
      // .addItem('合并表格', 'showMergeDialog')
      .addItem('配置表检查(开发中)', 'checkConfigTable')  // 新增配置检查功能
      .addItem('清除所有标记', 'clearAllMarks')
      .addItem('手动更新冲突标记', 'validateAndClearConflictMarks')  // 🆕 新增手动清理功能
      .addItem('刷新触发器', 'createEditTrigger')
      .addToUi();
  } catch (error) {
    console.error('Error creating menu: ' + error.toString());
  }
}

/**
 * 当编辑表格时的触发器（优化版）
 */
function onEditHandler(e) {
  const startTime = Date.now();
  
  try {
    // 快速预检查：是否需要处理
    // if (!quickPreCheck(e)) {
    //  console.log(`⚡ [快速退出] 无需处理 - 耗时: ${Date.now() - startTime}ms`);
    //  return;
    // }
    
    const context = createEditContextOptimized(e);
    if (!context.hasIdColumn) {
      console.log(`⚡ [无ID列] 跳过处理 - 耗时: ${Date.now() - startTime}ms`);
      return;
    }

    // 并行处理值追踪和ID冲突检查
    handleValueTrackingOptimized(context);
    handleIdConflictCheck(context);
    
    console.log(`✅ [触发器完成] 总耗时: ${Date.now() - startTime}ms`);
  } catch (error) {
    console.error('onEditHandler触发器出错:', error);
    console.log(`❌ [触发器异常] 耗时: ${Date.now() - startTime}ms`);
  }
}

/**
 * 快速预检查
 */
function quickPreCheck(e) {
  // 跳过大范围编辑（如复制粘贴大量数据）
  if (e.range.getNumRows() > 10 || e.range.getNumColumns() > 10) {
    return false;
  }
  
  // 跳过表头编辑
  if (e.range.getRow() === 1) {
    return false;
  }
  
  return true;
}

/**
 * 优化的编辑上下文创建
 */
function createEditContextOptimized(e) {
  const sheet = e.range.getSheet();
  
  // 只读取必要的表头范围
  const lastCol = sheet.getLastColumn();
  const headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  
  // 快速检查是否有ID列
  const hasIdColumn = headerRow.some(header => 
    header && header.toString().endsWith(ID_CHECKER_CONFIG.ID_COLUMN_SUFFIX)
  );

  return {
    sheet,
    range: e.range,
    oldValue: e.oldValue,
    newValue: e.range.getValue(),
    headerRow,
    hasIdColumn,
    lastCol
  };
}

/**
 * 优化的值追踪处理
 */
function handleValueTrackingOptimized(context) {
  const { range, oldValue, newValue } = context;
  
  // 批量获取当前状态，减少API调用
  const currentBg = range.getBackground();

  // 如果已经是新增状态，直接返回
  if (currentBg === SHEET_CONSTANTS.COLORS.ADDED) return;

  if (oldValue !== undefined) {
    // 修改状态
    const isAlreadyModified = currentBg === SHEET_CONSTANTS.COLORS.MODIFIED;
    range.setBackground(SHEET_CONSTANTS.COLORS.MODIFIED);
    
    // 只在首次修改时记录基准值（异步处理）
    if (!isAlreadyModified) {
      recordBaseValueAsync(range, oldValue);
    }
  } else if (newValue && newValue.toString().trim() !== '') {
    // 新增状态
    range.setBackground(SHEET_CONSTANTS.COLORS.ADDED);
  }
}

/**
 * 异步记录基准值（不阻塞主流程）
 */
function recordBaseValueAsync(range, oldValue) {
  // 使用时间触发器异步处理注释更新
  const trigger = ScriptApp.newTrigger('updateBaseValueNote')
    .timeBased()
    .after(100)
    .create();
  
  // 存储参数
  const params = {
    sheetName: range.getSheet().getName(),
    row: range.getRow(),
    column: range.getColumn(),
    baseValue: oldValue.toString(),
    triggerId: trigger.getUniqueId()
  };
  
  PropertiesService.getScriptProperties().setProperty(
    `base_value_${trigger.getUniqueId()}`, 
    JSON.stringify(params)
  );
}

/**
 * 延迟执行的基准值注释更新
 */
function updateBaseValueNote() {
  const triggers = ScriptApp.getProjectTriggers();
  const currentTrigger = triggers.find(t => t.getHandlerFunction() === 'updateBaseValueNote');
  
  if (!currentTrigger) return;
  
  try {
    const paramsJson = PropertiesService.getScriptProperties()
      .getProperty(`base_value_${currentTrigger.getUniqueId()}`);
    
    if (paramsJson) {
      const params = JSON.parse(paramsJson);
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(params.sheetName);
      const range = sheet.getRange(params.row, params.column);
      
      const note = range.getNote();
      const newNote = NoteManager.addSystemNote(
        note,
        NOTE_CONSTANTS.TYPES.BASE_VALUE,
        params.baseValue
      );
      range.setNote(newNote);
      
      // 清理
      PropertiesService.getScriptProperties().deleteProperty(`base_value_${currentTrigger.getUniqueId()}`);
    }
  } catch (error) {
    console.error('更新基准值注释失败:', error);
  } finally {
    // 删除触发器
    ScriptApp.deleteTrigger(currentTrigger);
  }
}

/**
 * 获取被编辑的列数组
 */
function getEditedColumns(range) {
  const editedColStart = range.getColumn();
  const editedColEnd = editedColStart + range.getNumColumns() - 1;
  return Array.from({ length: editedColEnd - editedColStart + 1 }, (_, i) => editedColStart + i);
}

/**
 * 清除所有标记和系统注释
 * @param {boolean} showConfirm 是否显示确认对话框
 * @returns {Object} 操作结果
 */
function clearAllMarks(showConfirm = true) {
  try {
    if (showConfirm) {
      const ui = SpreadsheetApp.getUi();
      const response = ui.alert(
        '确认清除',
        '是否要清除当前表格中所有的比较标记和系统注释？',
        ui.ButtonSet.YES_NO
      );

      if (response !== ui.Button.YES) {
        return { success: false, message: "操作已取消" };
      }
    }

    const sheet = SpreadsheetApp.getActiveSheet();
    const range = sheet.getDataRange();
    const [backgrounds, notes] = [range.getBackgrounds(), range.getNotes()];
    
    const newBackgrounds = backgrounds.map(row => 
      row.map(bg => {
        // 检查是否是系统背景色
        if ([
          COMPARE_CONSTANTS.COLORS.MODIFIED,
          COMPARE_CONSTANTS.COLORS.ADDED,
          COMPARE_CONSTANTS.COLORS.REMOVED,
          COMPARE_CONSTANTS.COLORS.HEADER_MODIFIED,
          SHEET_CONSTANTS.COLORS.MODIFIED,
          SHEET_CONSTANTS.COLORS.ADDED,
          MERGE_CONSTANTS.COLORS.NEW,
          MERGE_CONSTANTS.COLORS.CONFLICT,
          MERGE_CONSTANTS.COLORS.UPDATED,
          MERGE_CONSTANTS.COLORS.MERGED,
          MERGE_CONSTANTS.COLORS.RESOLVED,
          ID_CHECKER_CONFIG.COLORS.CONFLICT
        ].includes(bg)) {
          return null;  // 清除系统背景色
        }
        return bg;  // 保持非系统背景色不变
      })
    );

    const newNotes = notes.map(row =>
      row.map(note => note ? NoteManager.removeAllSystemNotes(note) : '')
    );

    // 更新表格
    range.setBackgrounds(newBackgrounds);
    range.setNotes(newNotes);

    return {
      success: true,
      message: "已清除所有标记和系统注释"
    };
  } catch (error) {
    console.error('清除标记和注释失败:', error);
    return {
      success: false,
      message: "清除失败: " + error.toString()
    };
  }
}

/**
 * Shows the add-on's homepage card when opened from the add-on menu
 * @param {Object} e The event object
 * @return {CardService.Card} The homepage card
 */
function onHomepage(e) {
  // 创建卡片UI
  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('Google Sheets Helper')
      .setImageUrl('https://writesome.oss-cn-chengdu.aliyuncs.com/logo.jpeg'))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText('选择以下操作:'))
      .addWidget(CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText('新建页签')
          .setOnClickAction(CardService.newAction().setFunctionName('cardCreateNewSheetTab')))
        .addButton(CardService.newTextButton()
          .setText('比较差异')
          .setOnClickAction(CardService.newAction().setFunctionName('cardShowCompareDialog')))
        .addButton(CardService.newTextButton()
          .setText('合并表格')
          .setOnClickAction(CardService.newAction().setFunctionName('cardShowMergeDialog')))
        .addButton(CardService.newTextButton()
          .setText('清除所有标记')
          .setOnClickAction(CardService.newAction().setFunctionName('cardClearAllMarks')))))
    .build();
  
  return card;
}

/**
 * 从卡片UI调用新建页签功能
 */
function cardCreateNewSheetTab(e) {
  // 将卡片UI操作转换为传统UI操作
  try {
    createNewSheetTab();
    return createSuccessCard('新建页签操作已启动');
  } catch (error) {
    return createErrorCard('新建页签失败: ' + error.toString());
  }
}

/**
 * 从卡片UI调用比较差异功能
 */
function cardShowCompareDialog(e) {
  try {
    showCompareDialog();
    return createSuccessCard('比较差异对话框已打开');
  } catch (error) {
    return createErrorCard('打开比较差异对话框失败: ' + error.toString());
  }
}

/**
 * 从卡片UI调用合并表格功能
 */
function cardShowMergeDialog(e) {
  try {
    showMergeDialog();
    return createSuccessCard('合并表格对话框已打开');
  } catch (error) {
    return createErrorCard('打开合并表格对话框失败: ' + error.toString());
  }
}

/**
 * 从卡片UI调用清除标记功能
 */
function cardClearAllMarks(e) {
  try {
    const result = clearAllMarks(false);
    if (result.success) {
      return createSuccessCard(result.message);
    } else {
      return createErrorCard(result.message);
    }
  } catch (error) {
    return createErrorCard('清除标记失败: ' + error.toString());
  }
}

/**
 * 配置表检查 - 调用外部接口
 */
function checkConfigTable() {
  try {
    const sheetName = SpreadsheetApp.getActiveSheet().getName();
    const userEmail = Session.getActiveUser().getEmail();
    
    // 构建请求URL，包含必要的查询参数
    const url = `https://script.google.com/a/macros/nibirutech.com/s/AKfycbzEsIxDkszo5CLZ4X1tewaDIC-udhCktaYSyBX6OmeiLKMmzn5rgnM0IwyKb9W9syI/exec?api=v1&action=runWorkflow&workflowId=WF1752049455621&apiKey=test-api-key-123&sheetName=${encodeURIComponent(sheetName)}`;
    
    // 准备POST请求数据
    const payload = {
      sheetName: sheetName,
      userEmail: userEmail
    };
    
    // 获取OAuth token
    let token;
    try {
      token = ScriptApp.getOAuthToken();
      console.log('OAuth token obtained successfully');
    } catch (tokenError) {
      console.error('Failed to get OAuth token:', tokenError);
      SpreadsheetApp.getActive().toast(
        `获取认证令牌失败: ${tokenError.toString()}`,
        '认证错误',
        5
      );
      return;
    }
    
    // 发送POST请求
    const response = UrlFetchApp.fetch(url, {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
      headers: {
        'Accept': 'application/json',
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
      }
    });
    
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    console.log('配置检查响应:', responseCode, responseText);
    console.log('请求URL:', url);
    console.log('请求头:', response.getHeaders());
    
    if (responseCode === 200) {
      SpreadsheetApp.getActive().toast(
        `配置检查请求已发送\n页签: ${sheetName}\n用户: ${userEmail}`,
        '检查请求成功',
        5
      );
    } else {
      SpreadsheetApp.getActive().toast(
        `配置检查请求失败\n状态码: ${responseCode}\n响应: ${responseText}`,
        '检查请求失败',
        8
      );
    }
    
  } catch (error) {
    console.error('配置检查失败:', error);
    SpreadsheetApp.getActive().toast(
      `配置检查失败: ${error.toString()}`,
      '错误',
      5
    );
  }
}

/**
 * 创建成功提示卡片
 */
function createSuccessCard(message) {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('操作成功'))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph().setText(message))
      .addWidget(CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText('返回主页')
          .setOnClickAction(CardService.newAction().setFunctionName('onHomepage')))))
    .build();
}

/**
 * 创建错误提示卡片
 */
function createErrorCard(message) {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('操作失败'))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph().setText(message))
      .addWidget(CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText('返回主页')
          .setOnClickAction(CardService.newAction().setFunctionName('onHomepage')))))
    .build();
}

/**
 * 当获得文件权限时触发
 * @param {Object} e 事件对象
 * @return {CardService.Card} 主页卡片
 */
function onFileScopeGranted(e) {
  return onHomepage(e);
}

function testMenuCreation() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu("测试菜单")
      .addItem('测试项', 'showAlert')
      .addToUi();
    
    // 创建一个简单的提示框，确认函数运行
    SpreadsheetApp.getActive().toast('菜单创建测试成功', '测试', 3);
  } catch (error) {
    console.error('菜单创建测试失败:', error);
  }
}

function showAlert() {
  SpreadsheetApp.getUi().alert('测试成功!');
}