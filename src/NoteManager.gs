// --------------------- NoteManager ------------------------ 

/**
 * 注释管理工具类
 */
class NoteManager {
  /**
   * 提取系统注释，自动合并多个系统注释块
   */
  static extractSystemNotes(note) {
    if (!note) return {};
    
    const systemNotes = {};
    let currentPosition = 0;
    let hasMultipleBlocks = false;
    
    // 查找所有系统注释块并合并
    while (true) {
      const start = note.indexOf(NOTE_CONSTANTS.SYSTEM_NOTE_START, currentPosition);
      if (start === -1) break;
      
      const end = note.indexOf(NOTE_CONSTANTS.SYSTEM_NOTE_END, start);
      if (end === -1) break;
      
      // 如果不是第一个块，标记存在多个块
      if (currentPosition > 0) {
        hasMultipleBlocks = true;
      }
      
      const notesSection = note.substring(
        start + NOTE_CONSTANTS.SYSTEM_NOTE_START.length,
        end
      );

      const noteRegex = /^(.+?):\s*\n([\s\S]*?)(?=\n\w+:|$)/gm;
      let match;
      
      while ((match = noteRegex.exec(notesSection)) !== null) {
        const [, key, value] = match;
        systemNotes[key.trim()] = value.trim();
      }
      
      currentPosition = end + NOTE_CONSTANTS.SYSTEM_NOTE_END.length;
    }
    
    // 如果发现多个块，自动清理并重写注释
    if (hasMultipleBlocks) {
      const cleanNote = this.removeAllSystemNotes(note);
      const systemPart = this.formatSystemNotes(systemNotes);
      const newNote = this.appendSystemNote(cleanNote, systemPart);
      
      // 如果是在单元格上下文中，尝试更新单元格注释
      try {
        const cell = SpreadsheetApp.getActiveRange();
        if (cell) {
          cell.setNote(newNote);
        }
      } catch (e) {
        // 忽略错误，因为可能不在单元格上下文中
      }
    }
    
    return systemNotes;
  }

  /**
   * 添加系统注释
   */
  static addSystemNote(originalNote, type, content) {
    // 添加参数验证
    if (!type || content === undefined) {
      throw new Error('Type and content are required');
    }
    
    // 直接使用 extractSystemNotes 进行合并处理
    const systemNotes = this.extractSystemNotes(originalNote);
    systemNotes[type] = content;
    
    const systemPart = this.formatSystemNotes(systemNotes);
    return this.appendSystemNote(this.removeAllSystemNotes(originalNote), systemPart);
  }
  
  /**
   * 获取系统注释内容
   * @param {string} note 完整注释
   * @param {string} type 注释类型
   * @returns {string|null} 系统注释内容
   */
  static getSystemNote(note, type) {
    const systemNotes = this.extractSystemNotes(note);
    return systemNotes[type] || null;
  }
  
  /**
   * 移除指定类型的系统注释
   * @param {string} note 完整注释
   * @param {string} type 注释类型
   * @returns {string} 清理后的注释
   */
  static removeSystemNote(note, type) {
    const systemNotes = this.extractSystemNotes(note);
    delete systemNotes[type];
    
    // 如果没有剩余的系统注释，返回清理后的原始注释
    if (Object.keys(systemNotes).length === 0) {
      return this.removeAllSystemNotes(note);
    }
    
    const systemPart = this.formatSystemNotes(systemNotes);
    return this.appendSystemNote(this.removeAllSystemNotes(note), systemPart);
  }
  
  /**
   * 格式化系统注释
   * @private
   * @param {Object} systemNotes 系统注释对象
   * @returns {string} 格式化后的系统注释
   */
  static formatSystemNotes(systemNotes) {
    if (Object.keys(systemNotes).length === 0) return '';
    
    const formattedNotes = Object.entries(systemNotes)
      .map(([key, value]) => `${key}${NOTE_CONSTANTS.KEY_VALUE_SEPARATOR}\n${value}`)
      .join(NOTE_CONSTANTS.LINE_SEPARATOR);
    
    return `${NOTE_CONSTANTS.SYSTEM_NOTE_START}${formattedNotes}${NOTE_CONSTANTS.SYSTEM_NOTE_END}`;
  }
  
  /**
   * 移除所有系统注释
   * @private
   * @param {string} note 完整注释
   * @returns {string} 移除系统注释后的原始注释
   */
  static removeAllSystemNotes(note) {
    if (!note) return '';
    
    const start = note.indexOf(NOTE_CONSTANTS.SYSTEM_NOTE_START);
    if (start === -1) return note;
    
    const end = note.indexOf(NOTE_CONSTANTS.SYSTEM_NOTE_END);
    if (end === -1) return note;
    
    return note.substring(0, start) + note.substring(end + NOTE_CONSTANTS.SYSTEM_NOTE_END.length);
  }
  
  /**
   * 在原始注释后追加系统注释
   * @private
   * @param {string} originalNote 原始注释
   * @param {string} systemNote 系统注释部分
   * @returns {string} 组合后的完整注释
   */
  static appendSystemNote(originalNote, systemNote) {
    if (!systemNote) return originalNote || '';
    if (!originalNote) return systemNote;
    
    return `${originalNote.trim()}\n${systemNote}`;
  }

  /**
   * 移除单元格中指定类型的标记
   * @param {Range} cell 目标单元格
   * @param {string} type 要移除的标记类型
   * @returns {boolean} 是否成功移除标记
   */
  static removeMarkFromCell(cell, type) {
    if (!cell) return false;
    
    const note = cell.getNote();
    if (!note) return false;
    
    const newNote = this.removeSystemNote(note, type);
    
    // 如果注释内容没有变化，说明没有找到对应类型的标记
    if (newNote === note) return false;
    
    // 如果新注释为空，则完全清除注释
    if (newNote.trim() === '') {
      cell.clearNote();
    } else {
      cell.setNote(newNote);
    }
    
    return true;
  }
} 

// --------------------- NoteManager ------------------------