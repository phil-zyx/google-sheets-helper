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
      if (trigger.getHandlerFunction() === 'onEdit') {
        hasEditTrigger = true;
      }
    });

    if (!hasEditTrigger) {
      const ss = SpreadsheetApp.getActive();
      ScriptApp.newTrigger('onEdit')
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
      .addItem('比较差异原生UI', 'nativeCompare')  // 使用原生UI的比较功能
      .addItem('合并表格原生UI', 'smartNativeMerge')  // 使用智能合并函数
      .addItem('比较差异', 'showCompareDialog')
      .addItem('合并表格', 'showMergeDialog')
      .addItem('配置表检查', 'checkConfigTable')  // 新增配置检查功能
      .addItem('清除所有标记', 'clearAllMarks')
      .addItem('刷新触发器', 'createEditTrigger')
      .addToUi();
  } catch (error) {
    console.error('Error creating menu: ' + error.toString());
  }
}

/**
 * 当编辑表格时的触发器
 * @param {Object} e 编辑事件对象
 */
function onEdit(e) {
  try {
    // 检查表格是否包含ID列
    const sheet = e.range.getSheet();
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const hasIdColumn = headerRow.some(header => 
      header && header.toString().endsWith(ID_CHECKER_CONFIG.ID_COLUMN_SUFFIX)
    );
    
    // 如果没有ID列，直接返回
    if (!hasIdColumn) return;

    // 1. 处理基准值记录和颜色标记
    const range = e.range;
    const oldValue = e.oldValue;
    const newValue = range.getValue();
    
    // 获取当前单元格的背景色和注释
    const currentBg = range.getBackground();
    let note = range.getNote();
    
    // 如果当前单元格已经是新增状态（绿色），则不做任何改变
    if (currentBg !== SHEET_CONSTANTS.COLORS.ADDED) {
    // 检查是否已经有修改记录（通过背景色判断）
    const isAlreadyModified = currentBg === SHEET_CONSTANTS.COLORS.MODIFIED;

    if (oldValue !== undefined) {  // 是修改操作
      // 设置为修改颜色（浅蓝色）
      range.setBackground(SHEET_CONSTANTS.COLORS.MODIFIED);
      
      // 只在首次修改时记录基准值
      if (!isAlreadyModified) {
        // 保持原始值的格式
        let baseValue = oldValue;
        // 如果是整数，确保以整数形式存储
        if (Number.isInteger(Number(oldValue))) {
          baseValue = parseInt(oldValue, 10);
        }
        
        // 添加基准值到系统注释，保留用户原有注释
        const newNote = NoteManager.addSystemNote(
          note,
          NOTE_CONSTANTS.TYPES.BASE_VALUE,
          baseValue.toString()
        );
        range.setNote(newNote);
      }
    } else if (newValue && newValue.toString().trim() !== '') {
      // 如果是新增值，设置为新增颜色（淡绿色）
      range.setBackground(SHEET_CONSTANTS.COLORS.ADDED);
    }
    }

    // 2. 处理ID检查 - 无论是否有oldValue都需要检查
    // 查找所有ID列
    const idColumns = [];
    for (let col = 1; col <= headerRow.length; col++) {
      const header = headerRow[col-1];
      if (header && header.toString().endsWith(ID_CHECKER_CONFIG.ID_COLUMN_SUFFIX)) {
        idColumns.push(col);
      }
    }
    
    // 如果存在ID列，并且编辑的单元格在ID列中，才检查冲突
    const editedColStart = range.getColumn();
    const editedColEnd = editedColStart + range.getNumColumns() - 1;
    const editedColumns = Array.from({ length: editedColEnd - editedColStart + 1 }, (_, i) => editedColStart + i);

    // 检查编辑的列是否与任何ID列重叠
    const isEditingIdColumn = editedColumns.some(col => idColumns.includes(col));

    if (idColumns.length > 0 && isEditingIdColumn) {
      // 设置一个短暂的延迟，确保值已经更新
      Utilities.sleep(100);
      
      // 筛选出被编辑的ID列
      const relevantIdColumns = idColumns.filter(idCol => editedColumns.includes(idCol));

      // 对每个相关的ID列进行检查
      for (const idCol of relevantIdColumns) {
        // 检查该行的ID列单元格
        const idRange = sheet.getRange(range.getRow(), idCol, range.getNumRows(), 1);
        checkIdConflicts({
          sheet: sheet,
          range: idRange
        });
      }
    }
  } catch (error) {
    console.error('onEdit触发器出错:', error);
  }
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