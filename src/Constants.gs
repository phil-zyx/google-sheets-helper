// --------------------- 常量 --------------------------------

// 表格相关常量
const SHEET_CONSTANTS = {
  COLORS: {
    MODIFIED: "#b3e5fc",  // 修改 - 浅蓝色
    ADDED: "#dcedc8",    // 新增 - 淡绿色
    HEADER_MODIFIED: "#fff9c4",  // 表头修改 - 浅黄色
    CONFLICT: "#f8bbd0"  // 冲突 - 粉色
  }
};

// 合并相关常量
const MERGE_CONSTANTS = {
  ID_SUFFIX: '_INT_id',
  CONFLICT_PREFIX: '冲突: ',
  PREVIEW_SUFFIX: '_预览',
  COLORS: {
    NEW: '#dcedc8',      // 浅绿色 - 新行
    CONFLICT: '#ffdce0', // 浅红色 - 冲突
    UPDATED: '#b3e5fc',  // 浅蓝色 - 已更新
    RESOLVED: "#e8f5e9", // 已解决 - 更浅的绿色
    MERGED: "#dfcd4d"    // 合并入表 - 橘色
  }
};

// 比较相关常量
const COMPARE_CONSTANTS = {
  COLORS: {
    MODIFIED: "#ffcdd2",  // 修改 - 浅红色
    ADDED: "#dcedc8",    // 新增 - 浅绿色
    REMOVED: "#ffdce0",  // 删除 - Git风格浅红色
    HEADER_MODIFIED: "#fff9c4"  // 表头修改 - 浅黄色
  }
};

// ID检查器相关常量
const ID_CHECKER_CONFIG = {
  COLORS: {
    CONFLICT: '#ff0000',  // 冲突标记颜色 - 红色
  },
  ID_COLUMN_SUFFIX: '_INT_id',   // ID列的后缀
};

// 注释相关常量
const NOTE_CONSTANTS = {
  // 系统注释使用键值对格式
  SYSTEM_NOTE_START: '===== 系统信息开始 =====\n',
  SYSTEM_NOTE_END: '\n===== 系统信息结束 =====',
  
  TYPES: {
    BASE_VALUE: 'BASE',     // 基准值
    CONFLICT: 'CONFLICT',   // 冲突信息
    MERGE_INFO: 'MERGE',    // 合并信息
    VERSION: 'VERSION',     // 版本信息
    SHEET_CREATION: 'CREATION'  // 页签创建信息
  },

  // 添加分隔符常量
  KEY_VALUE_SEPARATOR: ': ',  // 键值分隔符
  LINE_SEPARATOR: '\n'        // 行分隔符
};

// 日志相关常量
const LOG_CONSTANTS = {
  SHEET_NAME: "配置表工具操作日志表",
  HEADERS: [
    "时间",
    "操作类型",
    "操作人",
    "操作表名",
    "操作内容",
    "详细信息"
  ],
  TYPES: {
    MERGE: "合并操作",
    COMPARE: "比较操作",
    CONFLICT_RESOLVE: "冲突解决",
    SHEET_CREATE: "创建表格",
    SHEET_UPDATE: "更新表格",
    SHEET_DELETE: "删除表格"
  },
  // 日志保留配置
  RETENTION: {
    MAX_ROWS: 10000,        // 最大保留行数
    CLEANUP_THRESHOLD: 0.9,  // 清理阈值（当达到最大行数的90%时触发清理）
    CLEANUP_TARGET: 0.7,     // 清理目标（清理后保留最大行数的70%）
    MIN_DAYS: 30            // 最小保留天数（无论行数多少，30天内的日志都保留）
  }
};

// --------------------- 常量定义 -------------------------------- 