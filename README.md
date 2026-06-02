# Excel MCP 工具

<p align="center">
  <strong>功能强大的 MCP (Model Context Protocol) Excel 文件操作工具</strong><br>
  <em>让 AI 助手轻松读取、写入、格式化和管理 Excel 文件</em>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/version-1.0.0-blue.svg" alt="版本">
  <img src="https://img.shields.io/badge/node-%3E%3D18.0.0-green.svg" alt="Node.js">
  <img src="https://img.shields.io/badge/license-MIT-yellow.svg" alt="许可证">
  <img src="https://img.shields.io/badge/typescript-5.3%2B-3178c6.svg" alt="TypeScript">
  <img src="https://img.shields.io/badge/mcp-1.0-orange.svg" alt="MCP">
</p>

***

## 📋 目录

- [✨ 功能特性](#-功能特性)
- [🚀 安装指南](#-安装指南)
- [🎯 快速开始](#-快速开始)
- [📖 使用说明](#-使用说明)
  - [1. 读取 Excel 文件](#1-读取-excel-文件)
  - [2. 写入数据到 Excel](#2-写入数据到-excel)
  - [3. 格式化单元格](#3-格式化单元格)
  - [4. 管理工作表](#4-管理工作表)
  - [5. 获取文件信息](#5-获取文件信息)
- [⚙️ 配置选项](#️-配置选项)
- [🔧 API 参考](#-api-参考)
- [🛡️ 安全机制](#️-安全机制)
- [📊 性能指标](#-性能指标)
- [🏗️ 系统架构](#️-系统架构)
- [🤝 贡献指南](#-贡献指南)
- [📄 开源许可](#-开源许可)

## ✨ 功能特性

### 📖 **读取功能**

- ✅ 读取整个工作表或指定范围（如 `A1:C100`）
- ✅ 支持 `.xlsx` 和 `.xls` 格式
- ✅ 可选提取单元格公式
- ✅ 可选提取格式化信息
- ✅ 自动识别表头行
- ✅ 返回结构化的 JSON 数据

### ✍️ **写入功能**

- ✅ 创建新的 Excel 文件或编辑现有文件
- ✅ 多种写入模式：覆盖、追加、更新
- ✅ 支持自定义起始位置
- ✅ 支持配置表头行
- ✅ 高性能批量数据写入

### 🎨 **格式化功能**

- ✅ 字体样式设置（字体族、大小、粗体、斜体、颜色）
- ✅ 背景填充颜色
- ✅ 边框定制（6种线型：细线、中等、粗线、虚线、点线、双线）
- ✅ 文本对齐方式（水平、垂直、自动换行）
- ✅ 数字格式化（货币、百分比、日期等）
- ✅ 批量范围格式化

### 📑 **工作表管理**

- ✅ 列出所有工作表及元数据
- ✅ 创建新工作表
- ✅ 删除现有工作表
- ✅ 重命名工作表
- ✅ 复制工作表（包含数据）
- ✅ 移动工作表位置

### 🛡️ **企业级特性**

- ✅ Zod 模式验证所有参数
- ✅ 完善的错误处理和错误码体系
- ✅ 文件路径安全验证
- ✅ 文件大小限制（可配置）
- ✅ 格式白名单保护
- ✅ 异步非阻塞 I/O 操作

## 🚀 安装指南

### 系统要求

- **Node.js**: >= 18.0.0
- **npm**: >= 9.0.0 （或 yarn/pnpm）

### 通过 npm 安装

```bash
# 全局安装（推荐用于命令行使用）
npm install -g @itkmoon/excel-mcp

# 或在项目中本地安装
npm install @itkmoon/excel-mcp
```

### 从源码构建

```bash
# 克隆仓库
git clone https://github.com/ITKMUnigle/Excel-MCP-Tool.git
cd Excel-MCP-Tool

# 安装依赖
npm install

# 构建项目
npm run build

# 启动服务器
npm run start:stdio
```

## 🎯 快速开始

### 在 Claude Desktop 中使用

在 Claude Desktop 配置文件 (`claude_desktop_config.json`) 中添加：

```json
{
  "mcpServers": {
    "excel": {
      "command": "npx",
      "args": ["@itkmoon/excel-mcp", "--stdio"]
    }
  }
}
```

### 在 Cursor IDE 中使用

在 `.cursor/mcp.json` 中添加：

```json
{
  "mcpServers": {
    "excel": {
      "command": "npx",
      "args": ["@itkmoon/excel-mcp", "--stdio"],
      "env": {}
    }
  }
}
```

### 命令行使用

```bash
# 启动 stdio 传输模式（默认）
excel-mcp --stdio

# 显示帮助信息
excel-mcp --help
```

## 📖 使用说明

### 1. 读取 Excel 文件

**读取整个工作表：**

```typescript
const result = await mcpClient.call('read_excel', {
  filePath: '/path/to/data.xlsx',
  sheetName: '销售数据'
});

// 返回结果:
{
  success: true,
  fileName: 'data.xlsx',
  sheetName: '销售数据',
  headers: ['姓名', '年龄', '部门', '薪资'],
  data: [
    ['张三', 28, '技术部', 85000],
    ['李四', 32, '市场部', 75000],
    ['王五', 25, '销售部', 65000]
  ],
  rowCount: 3,
  columnCount: 4
}
```

**读取指定范围：**

```typescript
const partialData = await mcpClient.call('read_excel', {
  filePath: '/path/to/data.xlsx',
  range: 'A1:D10'
});
```

**读取包含公式和格式的详细信息：**

```typescript
const detailedData = await mcpClient.call('read_excel', {
  filePath: '/path/to/financial.xlsx',
  includeFormulas: true,
  includeFormatting: true
});
```

### 2. 写入数据到 Excel

**创建新文件并写入带表头的数据：**

```typescript
await mcpClient.call('write_excel', {
  filePath: '/path/to/output.xlsx',
  headers: ['员工姓名', '部门', '入职日期', '薪资'],
  data: [
    ['张三', '技术部', '2023-01-15', 90000],
    ['李四', '市场部', '2023-03-20', 78000],
    ['王五', '销售部', '2023-06-01', 70000]
  ]
});

// 返回结果:
{
  success: true,
  filePath: '/path/to/output.xlsx',
  sheetName: 'Sheet1',
  rowsWritten: 3,
  message: '成功写入 3 行数据到工作表 "Sheet1"'
}
```

**向已有文件追加数据：**

```typescript
await mcpClient.call('write_excel', {
  filePath: '/path/to/existing.xlsx',
  mode: 'append',
  data: [
    ['赵六', '人事部', '2023-09-10', 72000]
  ]
});
```

**更新指定位置的单元格：**

```typescript
await mcpClient.call('write_excel', {
  filePath: '/path/to/data.xlsx',
  mode: 'update',
  startCell: 'B5',
  data: [
    ['更新后的值']
  ]
});
```

### 3. 格式化单元格

**格式化标题行（专业报表风格）：**

```typescript
await mcpClient.call('format_excel', {
  filePath: '/path/to/report.xlsx',
  range: 'A1:F1',
  font: {
    name: '微软雅黑',
    size: 14,
    bold: true,
    color: 'FFFFFFFF'  // 白色文字
  },
  fill: {
    type: 'solid',
    color: 'FF4472C4'   // 蓝色背景
  },
  alignment: {
    horizontal: 'center',
    vertical: 'middle'
  },
  border: {
    top: { style: 'thin', color: 'FF000000' },
    bottom: { style: 'thin', color: 'FF000000' },
    left: { style: 'thin', color: 'FF000000' },
    right: { style: 'thin', color: 'FF000000' }
  }
});
```

**应用货币格式（人民币）：**

```typescript
await mcpClient.call('format_excel', {
  filePath: '/path/to/financial.xlsx',
  range: 'B2:B100',
  numberFormat: '¥#,##0.00'
});

// 示例输出: ¥12,345.67
```

**应用百分比格式：**

```typescript
await mcpClient.call('format_excel', {
  filePath: '/path/to/data.xlsx',
  range: 'C2:C50',
  numberFormat: '0.00%'
});

// 示例输出: 85.50%
```

**应用日期格式：**

```typescript
await mcpClient.call('format_excel', {
  filePath: '/path/to/report.xlsx',
  range: 'D2:D100',
  numberFormat: 'yyyy-mm-dd'
});

// 示例输出: 2024-01-15
```

### 4. 管理工作表

**列出所有工作表：**

```typescript
const sheets = await mcpClient.call('manage_sheets', {
  filePath: '/path/to/workbook.xlsx',
  action: 'list'
});

// 返回结果:
{
  success: true,
  action: 'list',
  sheets: [
    { name: 'Sheet1', id: 1, rowCount: 100, columnCount: 10 },
    { name: '数据', id: 2, rowCount: 500, columnCount: 8 },
    { name: '汇总', id: 3, rowCount: 20, columnCount: 5 }
  ],
  message: '共找到 3 个工作表'
}
```

**创建新工作表：**

```typescript
await mcpClient.call('manage_sheets', {
  filePath: '/path/to/workbook.xlsx',
  action: 'create',
  sheetName: '新数据表'
});
```

**重命名工作表：**

```typescript
await mcpClient.call('manage_sheets', {
  filePath: '/path/to/workbook.xlsx',
  action: 'rename',
  sheetName: '旧名称',
  newName: '新名称'
});
```

**复制工作表（包含所有数据）：**

```typescript
await mcpClient.call('manage_sheets', {
  filePath: '/path/to/workbook.xlsx',
  action: 'copy',
  sheetName: '模板'
});
// 将创建名为 "模板_copy" 的副本
```

**删除工作表：**

```typescript
await mcpClient.call('manage_sheets', {
  filePath: '/path/to/workbook.xlsx',
  action: 'delete',
  sheetName: '废弃的工作表'
});
```

### 5. 获取文件信息

```typescript
const info = await mcpClient.call('get_file_info', {
  filePath: '/path/to/data.xlsx'
});

// 返回结果:
{
  success: true,
  fileName: 'data.xlsx',
  filePath: '/绝对路径/to/data.xlsx',
  fileSize: 245760,              // 字节数
  lastModified: '2024-01-15T10:30:00.000Z',
  sheetCount: 3,
  sheets: ['Sheet1', '数据', '汇总'],
  isValid: true
}
```

## ⚙️ 配置选项

### 环境变量

在项目根目录创建 `.env` 文件：

```bash
# 文件操作的工作目录（绝对路径）
WORK_DIR=./data

# 最大文件大小限制（字节，默认：100MB）
MAX_FILE_SIZE=104857600

# 是否允许写操作。开启：true/yes/y；关闭：false/no/n（默认）
EXCEL_ALLOW_WRITE=false

# 是否启用调试模式（true/false）
DEBUG=false

# 服务器端口（HTTP 模式使用，如已实现）
SERVER_PORT=3000

# 传输类型：stdio 或 http
TRANSPORT_TYPE=stdio
```

`EXCEL_ALLOW_WRITE` 支持以下值（大小写不敏感）：

| 值 | 含义 |
|----|------|
| `true` / `yes` / `y` | 启用写入、格式化、创建/删除/重命名/复制/移动工作表 |
| `false` / `no` / `n` | 禁用写操作，仅允许读取、获取文件信息和列出工作表 |

### 数字格式参考表

| 格式类型  | 格式模式                  | 示例                  |
| ----- | --------------------- | ------------------- |
| 美元货币  | `$#,##0.00`           | $1,234.56           |
| 人民币货币 | `¥#,##0.00`           | ¥1,234.56           |
| 百分比   | `0.00%`               | 85.50%              |
| 日期    | `yyyy-mm-dd`          | 2024-01-15          |
| 日期时间  | `yyyy-mm-dd hh:mm:ss` | 2024-01-15 14:30:00 |
| 数值    | `#,##0.00`            | 1,234.56            |
| 整数    | `#,##0`               | 1,235               |
| 科学计数法 | `0.00E+00`            | 1.23E+03            |

### 颜色格式说明（ARGB）

颜色使用 8 位十六进制 ARGB 格式：

- `FF` = Alpha 通道（透明度，FF 表示完全不透明）
- `RRGGBB` = 红色、绿色、蓝色值

常用颜色示例：

- `FFFFFFFF` - 白色
- `FF000000` - 黑色
- `FFFF0000` - 红色
- `FF00FF00` - 绿色
- `FF0000FF` - 蓝色
- `FF4472C4` - Excel 默认蓝色
- `FFED7D31` - Excel 默认橙色
- `FFFFC000` - Excel 默认黄色
- `FF70AD47` - Excel 默认绿色

## 🔧 API 参考

### 工具：`read_excel`

**功能描述**：读取 Excel 文件内容并返回结构化 JSON 数据

**参数说明**：

| 参数名                 | 类型      | 必填  | 默认值    | 描述                             |
| ------------------- | ------- | --- | ------ | ------------------------------ |
| `filePath`          | string  | ✅ 是 | -      | Excel 文件的完整路径（支持 .xlsx 或 .xls） |
| `sheetName`         | string  | ❌ 否 | 第一个工作表 | 目标工作表名称                        |
| `range`             | string  | ❌ 否 | 全部数据   | 单元格范围（例如："A1:C100"）            |
| `includeFormulas`   | boolean | ❌ 否 | false  | 是否包含公式信息                       |
| `includeFormatting` | boolean | ❌ 否 | false  | 是否包含格式化详情                      |

**返回值**：`ReadExcelResult` 或 `ErrorResponse`

***

### 工具：`write_excel`

**功能描述**：以多种模式向 Excel 文件写入数据

**参数说明**：

| 参数名                 | 类型        | 必填  | 默认值         | 描述                                              |
| ------------------- | --------- | --- | ----------- | ----------------------------------------------- |
| `filePath`          | string    | ✅ 是 | -           | 文件路径（不存在时会自动创建）                                 |
| `sheetName`         | string    | ❌ 否 | "Sheet1"    | 工作表名称                                           |
| `data`              | array     | ✅ 是 | -           | 要写入的二维数组数据                                      |
| `startCell`         | string    | ❌ 否 | "A1"        | 起始单元格位置                                         |
| `mode`              | 枚举        | ❌ 否 | "overwrite" | "overwrite"(覆盖) \| "append"(追加) \| "update"(更新) |
| `createIfNotExists` | boolean   | ❌ 否 | true        | 文件不存在时是否自动创建                                    |
| `headers`           | string\[] | ❌ 否 | -           | 表头数组（仅在覆盖模式下生效）                                 |

**返回值**：`WriteExcelResult` 或 `ErrorResponse`

***

### 工具：`format_excel`

**功能描述**：对单元格应用格式化样式

**参数说明**：

| 参数名            | 类型     | 必填  | 默认值    | 描述                |
| -------------- | ------ | --- | ------ | ----------------- |
| `filePath`     | string | ✅ 是 | -      | 文件路径              |
| `sheetName`    | string | ❌ 否 | 第一个工作表 | 工作表名称             |
| `range`        | string | ✅ 是 | -      | 单元格范围（例如："A1:F1"） |
| `font`         | object | ❌ 否 | -      | 字体设置对象            |
| `fill`         | object | ❌ 否 | -      | 背景填充设置对象          |
| `border`       | object | ❌ 否 | -      | 边框设置对象            |
| `alignment`    | object | ❌ 否 | -      | 对齐方式设置对象          |
| `numberFormat` | string | ❌ 否 | -      | 数字格式模式            |

**字体对象属性**：

| 属性名      | 类型      | 描述                            |
| -------- | ------- | ----------------------------- |
| `name`   | string  | 字体族名称（例如："Arial"、"微软雅黑"、"宋体"） |
| `size`   | number  | 字体大小（例如：12、14、16）             |
| `bold`   | boolean | 是否加粗                          |
| `italic` | boolean | 是否斜体                          |
| `color`  | string  | ARGB 颜色代码                     |

**边框对象属性**：

| 属性名                                 | 类型     | 描述                                                                                        |
| ----------------------------------- | ------ | ----------------------------------------------------------------------------------------- |
| `top` / `bottom` / `left` / `right` | object | 各边框的设置                                                                                    |
| `style`                             | string | "thin"(细线) \| "medium"(中等) \| "thick"(粗线) \| "dashed"(虚线) \| "dotted"(点线) \| "double"(双线) |
| `color`                             | string | ARGB 颜色代码                                                                                 |

**返回值**：`FormatCellResult` 或 `ErrorResponse`

***

### 工具：`manage_sheets`

**功能描述**：管理工作簿中的工作表

**参数说明**：

| 参数名         | 类型     | 必填     | 描述                                                                                      |
| ----------- | ------ | ------ | --------------------------------------------------------------------------------------- |
| `filePath`  | string | ✅ 是    | 文件路径                                                                                    |
| `action`    | 枚举     | ✅ 是    | "list"(列出) \| "create"(创建) \| "delete"(删除) \| "rename"(重命名) \| "copy"(复制) \| "move"(移动) |
| `sheetName` | string | 条件必填\* | 目标工作表名称（\*大多数操作必需）                                                                      |
| `newName`   | string | 条件必填\* | 新名称（\*重命名操作时必需）                                                                         |
| `index`     | number | 条件必填\* | 位置索引（\*移动操作时必需，从 0 开始）                                                                  |

**返回值**：`SheetManagementResult` 或 `ErrorResponse`

***

### 工具：`get_file_info`

**功能描述**：获取文件元数据和工作表列表

**参数说明**：

| 参数名        | 类型     | 必填  | 描述   |
| ---------- | ------ | --- | ---- |
| `filePath` | string | ✅ 是 | 文件路径 |

**返回值**：`FileInfo` 或 `ErrorResponse`

***

### 错误响应格式

所有错误均遵循以下统一结构：

```typescript
interface ErrorResponse {
  success: false;
  error: {
    code: ErrorCode;       // 错误码枚举
    message: string;        // 人类可读的错误消息
    details?: any;          // 额外的上下文信息
  };
}
```

**错误码列表**：

| 错误码                 | 描述         |
| ------------------- | ---------- |
| `FILE_NOT_FOUND`    | 文件不存在      |
| `INVALID_FORMAT`    | 不支持的文件格式   |
| `SHEET_NOT_FOUND`   | 工作表不存在     |
| `INVALID_RANGE`     | 无效的单元格范围格式 |
| `PERMISSION_DENIED` | 权限不足       |
| `WRITE_ERROR`       | 文件写入失败     |
| `VALIDATION_ERROR`  | 参数验证失败     |
| `UNKNOWN_ERROR`     | 未知错误       |

## 🛡️ 安全机制

### 内置安全保护

1. **路径遍历防护**
   - 验证文件路径以防止目录遍历攻击
   - 可配置为限制在工作目录内访问
2. **文件大小限制**
   - 默认最大值：每个文件 100MB
   - 可通过 `MAX_FILE_SIZE` 环境变量配置
   - 防止内存耗尽攻击
3. **格式白名单**
   - 仅接受 `.xlsx` 和 `.xls` 扩展名
   - 拒绝潜在的危险文件类型
4. **输入验证**
   - 所有参数使用 Zod 模式进行严格验证
   - 自动类型转换和数据清洗
   - 防止注入攻击
5. **错误信息安全控制**
   - 安全敏感失败时返回通用错误消息
   - 详细错误信息仅在调试模式下显示
   - 不暴露内部堆栈跟踪信息
6. **写操作权限控制**
   - 默认禁用写入、格式化和会修改工作簿的工作表管理操作
   - 可通过 `EXCEL_ALLOW_WRITE=true`、`yes` 或 `y` 启用
   - 关闭值支持 `false`、`no`、`n`

### 用户最佳实践

✅ **推荐做法：**

- 尽可能使用绝对路径
- 修改重要文件前先备份
- 验证来自用户输入的文件路径
- 设置合理的文件大小限制
- 使用环境变量存储敏感配置

❌ **避免做法：**

- 不经验证就处理不受信任的文件路径
- 在没有身份认证的情况下将服务器暴露到公网
- 在生产环境中禁用文件大小限制
- 忽略工具返回的错误响应

## 📊 性能指标

### 基准测试

测试环境：Windows 11, Node.js 20.x, 16GB 内存

| 操作类型  | 数据规模            | 平均耗时    | 内存占用    |
| ----- | --------------- | ------- | ------- |
| 读取    | 1,000 行 × 20 列  | < 200ms | \~50MB  |
| 读取    | 10,000 行 × 20 列 | < 800ms | \~120MB |
| 写入    | 1,000 行 × 20 列  | < 300ms | \~80MB  |
| 写入    | 10,000 行 × 20 列 | < 1.5s  | \~200MB |
| 格式化   | 1,000 个单元格      | < 150ms | \~60MB  |
| 列出工作表 | 10 个工作表         | < 50ms  | \~30MB  |

### 性能优化建议

1. **使用范围限制**：当只需要部分数据时，请指定 `range` 参数
2. **避免不必要的格式提取**：只在需要时请求 `includeFormatting`
3. **批量操作**：一次写入多行数据，而不是多次单独调用
4. **内存管理**：对于超大文件（>50MB），考虑分块处理

## 🏗️ 系统架构

### 整体架构图

```mermaid
flowchart TB
    subgraph Client["🖥️ 客户端层 Client Layer"]
        C1["Claude Desktop"]
        C2["Cursor IDE"]
        C3["Trae AI"]
        C4["自定义应用<br/>Custom App"]
    end

    subgraph MCP["⚙️ MCP 服务器层 Server Layer"]
        direction LR
        subgraph Tools["📦 工具集 Tools"]
            T1["📖 read_excel<br/>读取工具"]
            T2["✍️ write_excel<br/>写入工具"]
            T3["🎨 format_excel<br/>格式化工具"]
            T4["📑 manage_sheets<br/>工作表管理"]
            T5["ℹ️ get_file_info<br/>文件信息"]
        end
    end

    subgraph Service["🔧 服务层 Service Layer"]
        S1["ExcelService<br/>核心业务服务"]
        S2["ValidationService<br/>参数验证服务"]
    end

    subgraph Data["💾 数据层 Data Layer"]
        D1["ExcelJS 库<br/>Excel 引擎"]
        D2["File System<br/>文件系统"]
    end

    subgraph Storage["📁 存储层 Storage Layer"]
        ST1[".xlsx 文件"]
        ST2[".xls 文件"]
    end

    %% 连接关系
    Client -->|"MCP Protocol<br/>JSON-RPC"| MCP
    Tools -->|"调用"| Service
    S1 -->|"读写操作"| Data
    S1 -->|"文件 I/O"| D2
    Data -->|"解析/生成"| Storage
    D2 -->|"访问"| Storage

    %% 样式定义
    classDef clientStyle fill:#e1f5fe,stroke:#0288d1,stroke-width:2px,color:#01579b
    classDef mcpStyle fill:#fff3e0,stroke:#ef6c00,stroke-width:2px,color:#e65100
    classDef toolStyle fill:#f3e5f5,stroke:#7b1fa2,stroke-width:2px,color:#4a148c
    classDef serviceStyle fill:#e8f5e9,stroke:#2e7d32,stroke-width:2px,color:#1b5e20
    classDef dataStyle fill:#fce4ec,stroke:#c62828,stroke-width:2px,color:#b71c1c
    classDef storageStyle fill:#fff9c4,stroke:#f9a825,stroke-width:2px,color:#f57f17

    class C1,C2,C3,C4 clientStyle
    class MCP mcpStyle
    class T1,T2,T3,T4,T5 toolStyle
    class S1,S2 serviceStyle
    class D1,D2 dataStyle
    class ST1,ST2 storageStyle
```

### 项目结构

```
excel-mcp-tool/
├── src/
│   ├── index.ts                 # 应用程序入口
│   ├── server.ts                # MCP 服务器配置
│   ├── types/
│   │   └── index.ts             # TypeScript 类型和 Zod 模式定义
│   ├── tools/
│   │   ├── read.ts              # 读取工具实现
│   │   ├── write.ts             # 写入工具实现
│   │   ├── format.ts            # 格式化工具实现
│   │   └── sheet-management.ts  # 工作表管理工具实现
│   ├── services/
│   │   └── excel.service.ts     # 核心 Excel 操作服务
│   └── utils/
│       ├── constants.ts         # 常量和配置定义
│       └── helpers.ts           # 辅助函数库
├── dist/                        # 编译后的 JavaScript 输出目录
├── package.json
├── tsconfig.json
├── .env.example
└── .gitignore
```

### 技术栈

| 组件       | 技术                        | 版本要求      |
| -------- | ------------------------- | --------- |
| 运行环境     | Node.js                   | >= 18.0.0 |
| 编程语言     | TypeScript                | 5.3+      |
| MCP 框架   | @modelcontextprotocol/sdk | ^1.0.4    |
| Excel 引擎 | exceljs                   | ^4.4.0    |
| 参数验证     | zod                       | ^3.22.4   |
| 包管理器     | npm/yarn/pnpm             | 最新版本      |

## 🤝 贡献指南

我们欢迎社区贡献！请按照以下步骤参与：

### 开发环境搭建

```bash
# Fork 并克隆仓库
git clone https://github.com/YOUR_USERNAME/Excel-MCP-Tool.git
cd Excel-MCP-Tool

# 安装依赖
npm install

# 启动开发模式（启用文件监视）
npm run dev

# 运行类型检查
npm run typecheck

# 代码检查
npm run lint

# 生产环境构建
npm run build
```

### 代码规范

- 使用 TypeScript 严格模式
- 遵循现有的代码约定
- 为公共 API 添加 JSDoc 注释
- 为新功能编写单元测试
- 提交前确保无 TypeScript 错误

### 提交更改

1. 创建功能分支：`git checkout -b feature/amazing-feature`
2. 进行修改并充分测试
3. 提交清晰的提交信息：`git commit -m 'Add amazing feature'`
4. 推送到您的 fork：`git push origin feature/amazing-feature`
5. 创建 Pull Request 并详细描述更改内容

### 问题反馈

如果您发现 bug 或有功能需求：

1. 先检查是否已有相关问题
2. 创建新 issue 并提供：
   - 清晰的标题和描述
   - 复现步骤（针对 bug）
   - 期望行为 vs 实际行为
   - 环境详情（操作系统、Node 版本等）
   - 相关截图或日志（如有）

## 📄 开源许可

本项目基于 MIT 许可证开源 - 详见 [LICENSE](LICENSE) 文件。

```
MIT License

Copyright (c) 2026 ITKMUnigle

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

## 🙏 致谢

- [Model Context Protocol (MCP)](https://modelcontextprotocol.io/) - 协议规范
- [ExcelJS](https://github.com/exceljs/exceljs) - 优秀的 Node.js Excel 处理库
- [Zod](https://zod.dev/) - 优先 TypeScript 的模式验证方案
- 开源社区提供的灵感和工具支持

## 📞 技术支持

- 📧 邮箱：524682687@qq.com
- 💬 问题反馈：[GitHub Issues](https://github.com/ITKMUnigle/Excel-MCP-Tool/issues)
- 📖 完整文档：[Wiki](https://github.com/ITKMUnigle/Excel-MCP-Tool/wiki)

***

<div align="center">

**⭐ 如果这个项目对你有帮助，请给它一个 star！⭐**

由 [ITKMUnigle](https://github.com/ITKMUnigle) 用 ❤️ 制作

</div>

***

## 🌐 其他语言版本

- [English Version](./README.en.md) | **中文版本（当前）**
