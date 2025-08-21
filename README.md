# Google Sheets Helper - google 表格工具

## 介绍

该仓库实现了一些基于 Google Sheets 表格 Apps Script 的工具类功能，主要专注于解决将 Goole Sheets 表格作为配置表需求的业务处理方案。

## 背景

一般情况下，策划运营都会熟练使用 Excel，在多人协作的项目中，通过使用 Google Sheets 来做配置表非常常见，而且云端支持。但是在多人协作中，就存在配置表的版本管理问题，id 引用，合并冲突问题等，针对这些问题，该项目提供了一些解决方案。

## 应用场景说明

1. 初始化配置表：针对配置结构建立原始仓库，比如 user 表
2. UI 按钮新建分支：新建后从原始表 copy 一份，作为新页签
3. UI 按钮合并分支：copy 的新页签合并到 base 表，自动合并，冲突检查处理
4. 提供一些工具函数：比如 base64 编解码
5. 提供 ESQL 检查：查询某些引用错误 bug 时，可以分支导入 mysql, 编写 SQL 语句来排查错误

## Todo

- [x] Base64 编解码函数
- [x] 支持两个表格页签的数据对比
  - [x] 提供页签对比界面
  - [x] 对比完成像 `git` 一样标识出差异项
- [x] 页签合并功能
  - [x] UI 按钮
  - [x] 将对比差异修改后，提供将差异合并的功能，有点类似 git merge
  - [x] 合并逻辑：只对修改进行检查，记录修改的原表值，冲突时对如果只是对原表进行修改则可以信任修改，新增按照ID合并
  - [x] 由于一般配置整行或整列删除是少数情况，所以没有对整行列删除进行智能合并，当前策略是忽略删除，合并时手动处理
- [x] ID 冲突检查
  - [x] 新增 `id` 后全页签检查，标记出冲突的ID，从基础上移除 ID 冲突
- [ ] SQL 查引用
  - [ ] 分支导入 mysql db
  - [ ] 实现界面 SQL 语句来查找引用错误

## 安装及使用方法

### 方案1

从商店直接搜索 google sheets help 使用。

### 方案2（可自行修改代码）

1. 打开 Google Sheets
2. 点击 "Extensions" > "Apps Script"
3. 复制以下文件到对应位置：
   - Code.gs
   - Compare.gs
   - CompareDialog.html
4. 保存后刷新界面

## 代码结构说明

### 🏗️ 整理后的代码架构

经过整理优化，`IdChecker.gs` 现在采用了清晰的模块化架构：

#### 1. **核心冲突检查模块**
- `checkSingleIdConflictImproved()` - 单次ID冲突检查（实时编辑检测）
- `checkCurrentSheetConflict()` - 检查当前表格内的ID冲突
- `checkOtherSheetsConflict()` - 检查其他表格的ID冲突
- `findConflictInSheet()` - 在指定表格中查找ID冲突

#### 2. **数据比较模块**
- `compareRowsData()` - 通用行数据比较函数
- `compareRowsDataOptimized()` - 优化的行数据比较函数（批量处理专用）

#### 3. **批量处理核心模块**
- `validateAndClearConflictMarks()` - 主入口函数，协调整个批量处理流程
- `readSheetDataBatch()` - 批量读取表格数据
- `detectConflictCells()` - 检测冲突标记单元格
- `validateConflictsBatch()` - 批量验证冲突状态
- `buildSheetHeadersCache()` - 构建表格表头缓存

#### 4. **冲突验证模块**
- `validateCellConflict()` - 验证单个单元格的冲突状态
- `categorizeCells()` - 分类单元格处理类型

#### 5. **批量更新模块**
- `updateCellsBatch()` - 批量更新单元格
- `clearConflictMarks()` - 清除冲突标记
- `updateConflictNotes()` - 更新冲突注释

#### 6. **兼容性和工具模块**
- `recheckCellConflictStatusCurrentSheet()` - 兼容性函数
- `showCleanupResults()` - 显示清理结果
- `checkIdConflicts()` - 实时冲突检查（保持原有功能）

### 🔧 代码优化亮点

#### 1. **函数职责分离**
```javascript
// 每个函数只做一件事，职责清晰
function readSheetDataBatch(sheet, lastRow, lastColumn) { /* 只负责读取数据 */ }
function detectConflictCells(batchData) { /* 只负责检测冲突 */ }
function validateConflictsBatch(conflictCells, sheetName, headers) { /* 只负责验证 */ }
```

#### 2. **统一的命名规范**
```javascript
// 使用描述性的函数名
checkCurrentSheetConflict()     // 检查当前表格冲突
findConflictInSheet()           // 在表格中查找冲突
compareRowsDataOptimized()      // 优化的行数据比较
```

#### 3. **清晰的参数文档**
```javascript
/**
 * 单次ID冲突检查 - 用于实时编辑检测
 * @param {Object} params - 检查参数
 * @param {string} params.value - 要检查的ID值
 * @param {string} params.sheet - 表格名称
 * @param {number} params.row - 行号
 * @param {number} params.column - 列号
 * @param {string} params.columnName - 列标题
 * @returns {Array} 冲突位置数组
 */
```

#### 4. **模块化组织**
```javascript
// ============================================================================
// 核心冲突检查函数
// ============================================================================

// ============================================================================
// 数据比较函数
// ============================================================================

// ============================================================================
// 批量冲突验证和清理
// ============================================================================
```

### 📊 性能优化总结

| 优化项目 | 优化前 | 优化后 | 提升倍数 |
|---------|--------|--------|----------|
| 数据读取 | 逐个单元格API调用 | 批量读取 | 3-5倍 |
| 表头获取 | 重复读取 | 缓存机制 | 2-3倍 |
| 冲突验证 | 重复检查 | 一次验证，信息复用 | 2倍 |
| 行数据比较 | 对象映射创建 | 直接数组比较 | 1.5-2倍 |
| **总体性能** | **基准** | **优化后** | **5-10倍** |

### 🎯 使用建议

1. **实时检查**: 使用 `checkSingleIdConflictImproved()` 进行单次冲突检查
2. **批量清理**: 使用 `validateAndClearConflictMarks()` 进行定期冲突标记清理
3. **性能监控**: 关注控制台日志中的性能指标
4. **代码维护**: 遵循模块化架构，新增功能时放在相应模块中

### 🔍 代码质量提升

- **可读性**: 清晰的函数命名和注释
- **可维护性**: 模块化架构，职责分离
- **可扩展性**: 易于添加新功能和优化
- **性能**: 批量处理，避免重复操作
- **稳定性**: 统一的错误处理机制

# Google Sheets Helper

一个强大的 Google Sheets 辅助工具，提供数据合并、比较和ID冲突检查功能。

## 主要功能

- **数据合并**: 智能合并多个表格的数据
- **数据比较**: 高效比较不同表格的数据差异
- **ID冲突检查**: 实时检测和验证ID冲突
- **批量处理**: 支持大量数据的批量操作

## 批量检查功能

### 🚀 新增：智能批量ID冲突检查

为了提高ID检查其他表时的性能，我们实现了智能批量读取功能：

#### 主要优化特性

1. **批量预加载**: 一次性读取所有表格的ID列数据，避免重复API调用
2. **智能选择**: 根据表格数量自动选择最优的检查策略
3. **表头缓存**: 预加载并缓存所有表格的表头信息
4. **性能监控**: 提供详细的性能测试和对比功能
5. **🚨 内存管理**: 智能分批处理，防止内存不足

#### 内存管理和分批处理

##### 为什么需要内存管理？

一次性读取很多页签确实可能存在内存不足的风险，特别是在 Google Apps Script 环境中。我们实现了智能的内存管理机制：

1. **自动分批处理**: 当检测到大量数据时，自动启用分批处理
2. **内存监控**: 实时监控内存使用情况，及时调整策略
3. **配置优化**: 根据内存使用情况动态调整批处理参数

##### 内存管理配置

```javascript
const BATCH_CHECKER_CONFIG = {
  MEMORY_MANAGEMENT: {
    // 是否启用内存监控
    ENABLED: true,
    
    // 单次读取的最大行数（避免一次性读取过多数据）
    MAX_ROWS_PER_BATCH: 1000,
    
    // 单次读取的最大表格数（避免同时处理过多表格）
    MAX_SHEETS_PER_BATCH: 10,
    
    // 内存使用警告阈值（MB）
    MEMORY_WARNING_THRESHOLD: 50,
    
    // 是否启用分批处理
    ENABLE_BATCH_PROCESSING: true,
    
    // 分批处理时的延迟时间（ms）
    BATCH_DELAY: 100
  }
};
```

##### 分批处理策略

| 数据规模 | 处理策略 | 内存占用 | 性能影响 |
|---------|----------|----------|----------|
| **小规模** (≤5表格, ≤1000行) | 一次性读取 | 低 | 最佳 |
| **中等规模** (6-10表格, 1001-2000行) | 智能分批 | 中 | 良好 |
| **大规模** (>10表格, >2000行) | 强制分批 | 高 | 稳定 |

#### 使用方法

##### 1. 自动智能选择（推荐）
```javascript
// 系统会自动选择最优的检查方式，并启用内存管理
const conflicts = checkSingleIdConflictImproved({
  value: "TEST001",
  sheet: "Sheet1",
  row: 2,
  column: 1,
  columnName: "ID"
});
```

##### 2. 手动选择
```javascript
// 传统方式（适合表格数量少的情况）
const conflicts = checkOtherSheetsConflict(ss, currentSheet, sheetName, value, row, columnName);

// 批量方式（适合表格数量多的情况）
const conflicts = checkOtherSheetsConflictBatch(ss, currentSheet, sheetName, value, row, columnName);

// 单行优化方式（适合表格数量很多的情况）
const conflicts = checkOtherSheetsConflictSingleRowOptimized(ss, currentSheet, sheetName, value, row, columnName);

// 智能选择（推荐）
const conflicts = checkOtherSheetsConflictSmart(ss, currentSheet, sheetName, value, row, columnName);
```

##### 3. 性能测试
```javascript
// 测试单个ID检查的性能
const result = performanceTestIdChecker("TEST001", "ID", "Sheet1", 2);

// 批量测试多个用例
const testCases = [
  { id: "TEST001", columnName: "ID", sheetName: "Sheet1", row: 2 },
  { id: "TEST002", columnName: "ID", sheetName: "Sheet1", row: 3 }
];
const results = batchPerformanceTest(testCases);

// 单行编辑检查性能测试
const results = singleRowPerformanceTest("TEST001", "ID", "Sheet1", 2);
```

##### 4. 内存管理
```javascript
// 监控内存使用情况
const memoryInfo = monitorMemoryUsage();

// 动态调整内存管理配置
const config = adjustMemoryManagementConfig({
  maxRowsPerBatch: 500,
  maxSheetsPerBatch: 5
});

// 内存压力测试
const stressTest = memoryStressTest(20, 5000);

// 跨页签检查性能分析
const analysis = crossSheetPerformanceAnalysis();
```

#### 配置选项

可以通过修改 `BATCH_CHECKER_CONFIG` 常量来自定义行为：

```javascript
const BATCH_CHECKER_CONFIG = {
  // 智能选择阈值：表格数量超过此值时使用批量方式
  SHEET_COUNT_THRESHOLD: 3,
  
  // 批量验证阈值：冲突单元格数量超过此值时使用批量验证
  CONFLICT_CELL_THRESHOLD: 5,
  
  // 超级批量验证阈值：冲突单元格数量超过此值时使用超级批量验证
  SUPER_BATCH_THRESHOLD: 20,
  
  // 是否启用智能选择
  ENABLE_SMART_SELECTION: true,
  
  // 是否启用批量验证
  ENABLE_BATCH_VALIDATION: true,
  
  // 是否启用超级批量验证
  ENABLE_SUPER_BATCH_VALIDATION: true,
  
  // 内存管理配置
  MEMORY_MANAGEMENT: {
    ENABLED: true,
    MAX_ROWS_PER_BATCH: 1000,
    MAX_SHEETS_PER_BATCH: 10,
    MEMORY_WARNING_THRESHOLD: 50,
    ENABLE_BATCH_PROCESSING: true,
    BATCH_DELAY: 100
  }
};
```

#### 性能提升效果

- **表格数量多时**: 性能提升可达 3-10 倍
- **数据量大时**: 批量读取减少 60-80% 的API调用
- **智能缓存**: 避免重复读取相同数据
- **🚨 内存安全**: 智能分批处理，防止内存不足

#### 适用场景

- ✅ 表格数量 ≥ 4 个
- ✅ 每个表格数据行数 ≥ 100 行
- ✅ 需要频繁进行ID冲突检查
- ✅ 对性能有较高要求
- ✅ **大量数据需要内存管理**

- ⚠️ 表格数量 ≤ 3 个（自动使用传统方式）
- ⚠️ 数据量很小（< 50 行）时提升不明显
- ⚠️ **内存受限环境（自动启用分批处理）**

#### 内存安全特性

1. **自动分批处理**: 当检测到大量数据时自动启用
2. **内存监控**: 实时监控内存使用情况
3. **配置优化**: 根据内存使用情况动态调整参数
4. **延迟控制**: 分批处理时添加延迟，避免API限制
5. **错误恢复**: 分批处理失败时自动降级到传统方式

## 安装和使用

1. 在 Google Sheets 中打开脚本编辑器
2. 复制相应的 `.gs` 文件内容到脚本编辑器
3. 保存并授权脚本
4. 在表格中使用相应的函数

## 注意事项

- 确保有足够的权限访问所有相关表格
- 大量数据操作时注意 Google Apps Script 的执行时间限制
- 建议在测试环境中先验证功能

## 技术支持

如有问题或建议，请查看代码注释或联系开发者。