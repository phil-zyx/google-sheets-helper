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
- [ ] 支持两个表格页签的数据对比
  - [x] 提供页签对比界面
  - [x] 对比完成像 `git` 一样标识出差异项
- [ ] 页签合并功能
  - [x] UI 按钮
  - [ ] 将对比差异修改后，提供将差异合并的功能，有点类似 git merge
- [ ] ID 冲突检查
  - [ ] 新增 `id` 后全页签检查，标记出冲突的ID，从基础上移除 ID 冲突
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
