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
      .addItem('比较差异', 'showCompareDialog')
      .addItem('合并表格', 'showMergeDialog')
      .addItem('清除所有标记', 'clearAllMarks')
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
    if (currentBg === SHEET_CONSTANTS.COLORS.ADDED) {
      return;
    }

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

    // 2. 处理ID检查 - 无论是否有oldValue都需要检查
    const column = range.getColumn();
    const headerRange = sheet.getRange(1, column);
    const headerValue = headerRange.getValue();
    
    // 检查是否编辑的是 ID 列
    if (headerValue && headerValue.toString().endsWith(ID_CHECKER_CONFIG.ID_COLUMN_SUFFIX)) {
      // 设置一个短暂的延迟，确保值已经更新
      Utilities.sleep(100);
      // 只检查 ID 列的单元格
      const idRange = sheet.getRange(range.getRow(), column, range.getNumRows(), 1);
      checkIdConflicts({
        sheet: sheet,
        range: idRange
      });
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
    const [backgrounds, notes, values] = [
      range.getBackgrounds(),
      range.getNotes(),
      range.getValues()
    ];
    
    // 添加保存水平对齐方式
    const horizontalAlignments = range.getHorizontalAlignments();

    const newBackgrounds = [];
    const newNotes = [];
    const newValues = [];
    const newHorizontalAlignments = []; // 添加新的对齐方式数组
    const rowsToKeep = [];

    // 检查每一行，标记需要保留的行
    for (let i = 0; i < backgrounds.length; i++) {
      let isDeletedRow = true;
      let hasHighlight = false;

      // 检查这一行是否是比较时新增的行
      for (let j = 0; j < backgrounds[i].length; j++) {
        const currentBg = backgrounds[i][j];
        const currentNote = notes[i][j];

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

        const backgroundRow = [];
        const noteRow = [];
        const alignmentRow = []; // 添加对齐方式行

        for (let j = 0; j < backgrounds[i].length; j++) {
          const currentBg = backgrounds[i][j];
          let currentNote = notes[i][j];
          const currentAlignment = horizontalAlignments[i][j]; // 获取当前对齐方式

          // 清除所有比较标记的背景色
          if (currentBg === COMPARE_CONSTANTS.COLORS.MODIFIED || 
              currentBg === COMPARE_CONSTANTS.COLORS.ADDED || 
              currentBg === COMPARE_CONSTANTS.COLORS.REMOVED || 
              currentBg === COMPARE_CONSTANTS.COLORS.HEADER_MODIFIED ||
              currentBg === SHEET_CONSTANTS.COLORS.MODIFIED ||
              currentBg === SHEET_CONSTANTS.COLORS.ADDED ||
              currentBg === MERGE_CONSTANTS.COLORS.NEW ||
              currentBg === MERGE_CONSTANTS.COLORS.CONFLICT ||
              currentBg === MERGE_CONSTANTS.COLORS.UPDATED ||
              currentBg === MERGE_CONSTANTS.COLORS.MERGED ||
              currentBg === MERGE_CONSTANTS.COLORS.RESOLVED) {
            backgroundRow.push(null);
          } else {
            backgroundRow.push(currentBg);
          }

          // 清除系统注释
          if (currentNote) {
            currentNote = NoteManager.removeAllSystemNotes(currentNote);
            noteRow.push(currentNote || '');
          } else {
            noteRow.push('');
          }
          
          // 保存对齐方式
          alignmentRow.push(currentAlignment);
        }

        newBackgrounds.push(backgroundRow);
        newNotes.push(noteRow);
        newValues.push(values[i]);
        newHorizontalAlignments.push(alignmentRow); // 添加对齐方式行到新数组
      }
    }

    // 更新表格
    if (rowsToKeep.length < backgrounds.length) {
      // 如果有行被删除，更新表格并删除多余的行
      const newRange = sheet.getRange(1, 1, newBackgrounds.length, backgrounds[0].length);
      newRange.setBackgrounds(newBackgrounds);
      newRange.setNotes(newNotes);
      newRange.setValues(newValues);
      newRange.setHorizontalAlignments(newHorizontalAlignments); // 设置水平对齐方式

      if (backgrounds.length > newBackgrounds.length) {
        sheet.deleteRows(newBackgrounds.length + 1, backgrounds.length - newBackgrounds.length);
      }
    } else {
      // 如果没有行被删除，只更新背景色和注释
      range.setBackgrounds(newBackgrounds);
      range.setNotes(newNotes);
      range.setHorizontalAlignments(newHorizontalAlignments); // 设置水平对齐方式
    }

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