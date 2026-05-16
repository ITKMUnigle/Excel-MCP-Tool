# Excel MCP Tool

<p align="center">
  <strong>A powerful MCP (Model Context Protocol) tool for Excel file operations</strong><br>
  <em>Read, write, format, and manage Excel files with AI assistance</em>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/version-1.0.0-blue.svg" alt="Version">
  <img src="https://img.shields.io/badge/node-%3E%3D18.0.0-green.svg" alt="Node.js">
  <img src="https://img.shields.io/badge/license-MIT-yellow.svg" alt="License">
  <img src="https://img.shields.io/badge/typescript-5.3%2B-3178c6.svg" alt="TypeScript">
  <img src="https://img.shields.io/badge/mcp-1.0-orange.svg" alt="MCP">
</p>

---

## 📋 Table of Contents

- [✨ Features](#-features)
- [🚀 Installation](#-installation)
- [🎯 Quick Start](#-quick-start)
- [📖 Usage Guide](#-usage-guide)
  - [1. Read Excel Files](#1-read-excel-files)
  - [2. Write Data to Excel](#2-write-data-to-excel)
  - [3. Format Cells](#3-format-cells)
  - [4. Manage Worksheets](#4-manage-worksheets)
  - [5. Get File Information](#5-get-file-information)
- [⚙️ Configuration](#️-configuration)
- [🔧 API Reference](#-api-reference)
- [🛡️ Security](#️-security)
- [📊 Performance](#-performance)
- [🏗️ Architecture](#️-architecture)
- [🤝 Contributing](#-contributing)
- [📄 License](#-license)

## ✨ Features

### 📖 **Read Operations**
- ✅ Read entire worksheets or specific ranges (e.g., `A1:C100`)
- ✅ Support for `.xlsx` and `.xls` formats
- ✅ Extract cell formulas (optional)
- ✅ Extract formatting information (optional)
- ✅ Auto-detect header rows
- ✅ Return structured JSON data

### ✍️ **Write Operations**
- ✅ Create new Excel files or edit existing ones
- ✅ Multiple write modes: overwrite, append, update
- ✅ Custom start position support
- ✅ Header row configuration
- ✅ Batch data writing with high performance

### 🎨 **Formatting Capabilities**
- ✅ Font styling (family, size, bold, italic, color)
- ✅ Background fill colors
- ✅ Border customization (6 line styles)
- ✅ Text alignment (horizontal, vertical, wrap text)
- ✅ Number formatting (currency, percentage, date, etc.)
- ✅ Batch range formatting

### 📑 **Worksheet Management**
- ✅ List all worksheets with metadata
- ✅ Create new worksheets
- ✅ Delete existing worksheets
- ✅ Rename worksheets
- ✅ Copy worksheets (with data)
- ✅ Move worksheet positions

### 🛡️ **Enterprise Features**
- ✅ Zod schema validation for all parameters
- ✅ Comprehensive error handling with error codes
- ✅ File path security validation
- ✅ File size limits (configurable)
- ✅ Format whitelist protection
- ✅ Async/await non-blocking I/O

## 🚀 Installation

### Prerequisites

- **Node.js**: >= 18.0.0
- **npm**: >= 9.0.0 (or yarn/pnpm)

### Install from npm

```bash
# Install globally (recommended for CLI usage)
npm install -g @itkmoon/excel-mcp

# Or install locally in your project
npm install @itkmoon/excel-mcp
```

### Build from Source

```bash
# Clone the repository
git clone https://github.com/ITKMUnigle/Excel-MCP-Tool.git
cd Excel-MCP-Tool

# Install dependencies
npm install

# Build the project
npm run build

# Start the server
npm run start:stdio
```

## 🎯 Quick Start

### Basic Usage with Claude Desktop

Add to your Claude Desktop configuration file (`claude_desktop_config.json`):

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

### Basic Usage with Cursor IDE

Add to your `.cursor/mcp.json`:

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

### Command Line Usage

```bash
# Start with stdio transport (default)
excel-mcp --stdio

# Show help
excel-mcp --help
```

## 📖 Usage Guide

### 1. Read Excel Files

**Read entire worksheet:**

```typescript
const result = await mcpClient.call('read_excel', {
  filePath: '/path/to/data.xlsx',
  sheetName: 'SalesData'
});

// Response:
{
  success: true,
  fileName: 'data.xlsx',
  sheetName: 'SalesData',
  headers: ['Name', 'Age', 'Department', 'Salary'],
  data: [
    ['Alice', 28, 'Engineering', 85000],
    ['Bob', 32, 'Marketing', 75000],
    ['Charlie', 25, 'Sales', 65000]
  ],
  rowCount: 3,
  columnCount: 4
}
```

**Read specific range:**

```typescript
const partialData = await mcpClient.call('read_excel', {
  filePath: '/path/to/data.xlsx',
  range: 'A1:D10'
});
```

**Read with formulas and formatting:**

```typescript
const detailedData = await mcpClient.call('read_excel', {
  filePath: '/path/to/financial.xlsx',
  includeFormulas: true,
  includeFormatting: true
});
```

### 2. Write Data to Excel

**Create a new file with headers:**

```typescript
await mcpClient.call('write_excel', {
  filePath: '/path/to/output.xlsx',
  headers: ['Employee Name', 'Department', 'Start Date', 'Salary'],
  data: [
    ['John Doe', 'Engineering', '2023-01-15', 90000],
    ['Jane Smith', 'Marketing', '2023-03-20', 78000],
    ['Mike Johnson', 'Sales', '2023-06-01', 70000]
  ]
});

// Response:
{
  success: true,
  filePath: '/path/to/output.xlsx',
  sheetName: 'Sheet1',
  rowsWritten: 3,
  message: 'Successfully wrote 3 rows to worksheet "Sheet1"'
}
```

**Append data to existing file:**

```typescript
await mcpClient.call('write_excel', {
  filePath: '/path/to/existing.xlsx',
  mode: 'append',
  data: [
    ['Sarah Wilson', 'HR', '2023-09-10', 72000]
  ]
});
```

**Update specific cells:**

```typescript
await mcpClient.call('write_excel', {
  filePath: '/path/to/data.xlsx',
  mode: 'update',
  startCell: 'B5',
  data: [
    ['Updated Value']
  ]
});
```

### 3. Format Cells

**Format header row:**

```typescript
await mcpClient.call('format_excel', {
  filePath: '/path/to/report.xlsx',
  range: 'A1:F1',
  font: {
    name: 'Arial',
    size: 14,
    bold: true,
    color: 'FFFFFFFF'  // White text
  },
  fill: {
    type: 'solid',
    color: 'FF4472C4'   // Blue background
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

**Apply currency format:**

```typescript
await mcpClient.call('format_excel', {
  filePath: '/path/to/financial.xlsx',
  range: 'B2:B100',
  numberFormat: '$#,##0.00'
});
```

**Apply percentage format:**

```typescript
await mcpClient.call('format_excel', {
  filePath: '/path/to/data.xlsx',
  range: 'C2:C50',
  numberFormat: '0.00%'
});
```

### 4. Manage Worksheets

**List all worksheets:**

```typescript
const sheets = await mcpClient.call('manage_sheets', {
  filePath: '/path/to/workbook.xlsx',
  action: 'list'
});

// Response:
{
  success: true,
  action: 'list',
  sheets: [
    { name: 'Sheet1', id: 1, rowCount: 100, columnCount: 10 },
    { name: 'Data', id: 2, rowCount: 500, columnCount: 8 },
    { name: 'Summary', id: 3, rowCount: 20, columnCount: 5 }
  ],
  message: 'Found 3 worksheets'
}
```

**Create a new worksheet:**

```typescript
await mcpClient.call('manage_sheets', {
  filePath: '/path/to/workbook.xlsx',
  action: 'create',
  sheetName: 'NewDataSheet'
});
```

**Rename a worksheet:**

```typescript
await mcpClient.call('manage_sheets', {
  filePath: '/path/to/workbook.xlsx',
  action: 'rename',
  sheetName: 'OldName',
  newName: 'NewName'
});
```

**Copy a worksheet:**

```typescript
await mcpClient.call('manage_sheets', {
  filePath: '/path/to/workbook.xlsx',
  action: 'copy',
  sheetName: 'Template'
});
// Creates "Template_copy"
```

**Delete a worksheet:**

```typescript
await mcpClient.call('manage_sheets', {
  filePath: '/path/to/workbook.xlsx',
  action: 'delete',
  sheetName: 'UnusedSheet'
});
```

### 5. Get File Information

```typescript
const info = await mcpClient.call('get_file_info', {
  filePath: '/path/to/data.xlsx'
});

// Response:
{
  success: true,
  fileName: 'data.xlsx',
  filePath: '/absolute/path/to/data.xlsx',
  fileSize: 245760,           // bytes
  lastModified: '2024-01-15T10:30:00.000Z',
  sheetCount: 3,
  sheets: ['Sheet1', 'Data', 'Summary'],
  isValid: true
}
```

## ⚙️ Configuration

### Environment Variables

Create a `.env` file in the project root:

```bash
# Working directory for file operations (absolute path)
WORK_DIR=./data

# Maximum file size in bytes (default: 100MB)
MAX_FILE_SIZE=104857600

# Enable debug mode (true/false)
DEBUG=false

# Server port (for HTTP mode, if implemented)
SERVER_PORT=3000

# Transport type: stdio or http
TRANSPORT_TYPE=stdio
```

### Number Formats Reference

| Format | Pattern | Example |
|--------|---------|---------|
| Currency (USD) | `$#,##0.00` | $1,234.56 |
| Currency (CNY) | `¥#,##0.00` | ¥1,234.56 |
| Percentage | `0.00%` | 85.50% |
| Date | `yyyy-mm-dd` | 2024-01-15 |
| DateTime | `yyyy-mm-dd hh:mm:ss` | 2024-01-15 14:30:00 |
| Number | `#,##0.00` | 1,234.56 |
| Integer | `#,##0` | 1,235 |
| Scientific | `0.00E+00` | 1.23E+03 |

### Color Format (ARGB)

Colors use 8-character hexadecimal ARGB format:
- `FF` = Alpha (opacity, FF = fully opaque)
- `RRGGBB` = Red, Green, Blue values

Examples:
- `FFFFFFFF` - White
- `FF000000` - Black
- `FFFF0000` - Red
- `FF00FF00` - Green
- `FF0000FF` - Blue
- `FF4472C4` - Excel default blue
- `FFED7D31` - Excel default orange

## 🔧 API Reference

### Tool: `read_excel`

**Description**: Read Excel file content and return structured JSON data

**Parameters**:

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `filePath` | string | ✅ Yes | - | Full path to Excel file (.xlsx or .xls) |
| `sheetName` | string | ❌ No | First sheet | Target worksheet name |
| `range` | string | ❌ No | All data | Cell range (e.g., "A1:C100") |
| `includeFormulas` | boolean | ❌ No | false | Include formula information |
| `includeFormatting` | boolean | ❌ No | false | Include formatting details |

**Returns**: `ReadExcelResult` or `ErrorResponse`

---

### Tool: `write_excel`

**Description**: Write data to Excel files with multiple modes

**Parameters**:

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `filePath` | string | ✅ Yes | - | File path (created if not exists) |
| `sheetName` | string | ❌ No | "Sheet1" | Worksheet name |
| `data` | array | ✅ Yes | - | 2D array of data to write |
| `startCell` | string | ❌ No | "A1" | Starting cell position |
| `mode` | enum | ❌ No | "overwrite" | "overwrite" \| "append" \| "update" |
| `createIfNotExists` | boolean | ❌ No | true | Create file if missing |
| `headers` | string[] | ❌ No | - | Header row (overwrite mode only) |

**Returns**: `WriteExcelResult` or `ErrorResponse`

---

### Tool: `format_excel`

**Description**: Apply formatting styles to cells

**Parameters**:

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `filePath` | string | ✅ Yes | - | File path |
| `sheetName` | string | ❌ No | First sheet | Worksheet name |
| `range` | string | ✅ Yes | - | Cell range (e.g., "A1:F1") |
| `font` | object | ❌ No | - | Font settings |
| `fill` | object | ❌ No | - | Background fill settings |
| `border` | object | ❌ No | - | Border settings |
| `alignment` | object | ❌ No | - | Alignment settings |
| `numberFormat` | string | ❌ No | - | Number format pattern |

**Font Object**:

| Property | Type | Description |
|----------|------|-------------|
| `name` | string | Font family (e.g., "Arial", "微软雅黑") |
| `size` | number | Font size (e.g., 12, 14, 16) |
| `bold` | boolean | Bold text |
| `italic` | boolean | Italic text |
| `color` | string | ARGB color code |

**Border Object**:

| Property | Type | Description |
|----------|------|-------------|
| `top` / `bottom` / `left` / `right` | object | Border side settings |
| `style` | string | "thin" \| "medium" \| "thick" \| "dashed" \| "dotted" \| "double" |
| `color` | string | ARGB color code |

**Returns**: `FormatCellResult` or `ErrorResponse`

---

### Tool: `manage_sheets`

**Description**: Manage worksheets in a workbook

**Parameters**:

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `filePath` | string | ✅ Yes | File path |
| `action` | enum | ✅ Yes | "list" \| "create" \| "delete" \| "rename" \| "copy" \| "move" |
| `sheetName` | string | Conditional* | Target worksheet name (*required for most actions) |
| `newName` | string | Conditional* | New name (*required for rename) |
| `index` | number | Conditional* | Position index (*required for move) |

**Returns**: `SheetManagementResult` or `ErrorResponse`

---

### Tool: `get_file_info`

**Description**: Get file metadata and worksheet list

**Parameters**:

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `filePath` | string | ✅ Yes | File path |

**Returns**: `FileInfo` or `ErrorResponse`

---

### Error Response Format

All errors follow this structure:

```typescript
interface ErrorResponse {
  success: false;
  error: {
    code: ErrorCode;       // Error code enum
    message: string;        // Human-readable message
    details?: any;          // Additional context
  };
}
```

**Error Codes**:

| Code | Description |
|------|-------------|
| `FILE_NOT_FOUND` | File does not exist |
| `INVALID_FORMAT` | Unsupported file format |
| `SHEET_NOT_FOUND` | Worksheet not found |
| `INVALID_RANGE` | Invalid cell range format |
| `PERMISSION_DENIED` | Insufficient permissions |
| `WRITE_ERROR` | File write failure |
| `VALIDATION_ERROR` | Parameter validation failed |
| `UNKNOWN_ERROR` | Unexpected error |

## 🛡️ Security

### Built-in Protections

1. **Path Traversal Prevention**
   - Validates file paths to prevent directory traversal attacks
   - Restricts access to working directory (if configured)

2. **File Size Limits**
   - Default maximum: 100MB per file
   - Configurable via `MAX_FILE_SIZE` environment variable
   - Prevents memory exhaustion attacks

3. **Format Whitelist**
   - Only accepts `.xlsx` and `.xls` extensions
   - Rejects potentially dangerous file types

4. **Input Validation**
   - All parameters validated using Zod schemas
   - Type coercion and sanitization
   - Prevents injection attacks

5. **Error Information Control**
   - Generic error messages for security-sensitive failures
   - Detailed errors only in debug mode
   - No internal stack traces exposed

### Best Practices for Users

✅ **DO:**
- Use absolute file paths when possible
- Backup important files before modification
- Validate file paths from user input
- Set appropriate file size limits
- Use environment variables for sensitive config

❌ **DON'T:**
- Process untrusted file paths without validation
- Expose the server on public networks without authentication
- Disable file size limits in production
- Ignore error responses from the tool

## 📊 Performance

### Benchmarks

Tested on: Windows 11, Node.js 20.x, 16GB RAM

| Operation | Data Scale | Avg Time | Memory Usage |
|-----------|-----------|----------|--------------|
| Read | 1,000 rows × 20 cols | < 200ms | ~50MB |
| Read | 10,000 rows × 20 cols | < 800ms | ~120MB |
| Write | 1,000 rows × 20 cols | < 300ms | ~80MB |
| Write | 10,000 rows × 20 cols | < 1.5s | ~200MB |
| Format | 1,000 cells | < 150ms | ~60MB |
| List Sheets | 10 worksheets | < 50ms | ~30MB |

### Optimization Tips

1. **Use Range Limiting**: Specify `range` parameter when you only need partial data
2. **Avoid Unnecessary Formatting**: Only request `includeFormatting` when needed
3. **Batch Operations**: Write multiple rows at once instead of individual calls
4. **Memory Management**: For very large files (>50MB), consider processing in chunks

## 🏗️ Architecture

### System Overview

```
┌─────────────────────────────────────────────────────┐
│                   Client Layer                       │
│         (Claude / Cursor / Trae / Custom App)        │
└──────────────────────┬──────────────────────────────┘
                       │ MCP Protocol
                       ▼
┌─────────────────────────────────────────────────────┐
│                  Server Layer                        │
│  ┌─────────┐ ┌──────────┐ ┌──────────┐ ┌─────────┐ │
│  │read_excel│write_excel│format_   │manage_  │ │
│  │  Tool   │  Tool     │excel Tool│sheets   │ │
│  └────┬────┘ └────┬─────┘ └────┬─────┘ └────┬────┘ │
│       └────────────┼────────────┼────────────┘      │
│                    ▼            ▼                    │
│  ┌─────────────────────────────────────────────┐    │
│  │              Service Layer                   │    │
│  │  ┌─────────────┐  ┌──────────────────────┐  │    │
│  │  │ ExcelService│  │ ValidationService    │  │    │
│  │  └─────────────┘  └──────────────────────┘  │    │
│  └─────────────────────────────────────────────┘    │
│                    │                                │
│                    ▼                                │
│  ┌─────────────────────────────────────────────┐    │
│  │              Data Layer                     │    │
│  │  ┌──────────┐  ┌─────────────────────────┐  │    │
│  │  │ ExcelJS  │  │ File System             │  │    │
│  │  │ Library  │  │                         │  │    │
│  │  └──────────┘  └─────────────────────────┘  │    │
│  └─────────────────────────────────────────────┘    │
└─────────────────────────────────────────────────────┘
                       │
                       ▼
┌─────────────────────────────────────────────────────┐
│                Storage Layer                         │
│              Excel Files (.xlsx/.xls)               │
└─────────────────────────────────────────────────────┘
```

### Project Structure

```
excel-mcp-tool/
├── src/
│   ├── index.ts                 # Application entry point
│   ├── server.ts                # MCP server configuration
│   ├── types/
│   │   └── index.ts             # TypeScript types & Zod schemas
│   ├── tools/
│   │   ├── read.ts              # Read tool implementation
│   │   ├── write.ts             # Write tool implementation
│   │   ├── format.ts            # Formatting tool implementation
│   │   └── sheet-management.ts  # Worksheet management tools
│   ├── services/
│   │   └── excel.service.ts     # Core Excel operations service
│   └── utils/
│       ├── constants.ts         # Constants & configurations
│       └── helpers.ts           # Utility functions
├── dist/                        # Compiled JavaScript output
├── package.json
├── tsconfig.json
├── .env.example
└── .gitignore
```

### Technology Stack

| Component | Technology | Version |
|-----------|------------|---------|
| Runtime | Node.js | >= 18.0.0 |
| Language | TypeScript | 5.3+ |
| MCP Framework | @modelcontextprotocol/sdk | ^1.0.4 |
| Excel Engine | exceljs | ^4.4.0 |
| Validation | zod | ^3.22.4 |
| Package Manager | npm/yarn/pnpm | Latest |

## 🤝 Contributing

We welcome contributions! Please follow these steps:

### Development Setup

```bash
# Fork and clone the repository
git clone https://github.com/YOUR_USERNAME/Excel-MCP-Tool.git
cd Excel-MCP-Tool

# Install dependencies
npm install

# Start development mode (with watch)
npm run dev

# Run type checking
npm run typecheck

# Lint code
npm run lint

# Build for production
npm run build
```

### Code Style

- Use TypeScript strict mode
- Follow existing code conventions
- Add JSDoc comments for public APIs
- Write unit tests for new features
- Ensure no TypeScript errors before committing

### Submitting Changes

1. Create a feature branch: `git checkout -b feature/amazing-feature`
2. Make your changes and test thoroughly
3. Commit with clear messages: `git commit -m 'Add amazing feature'`
4. Push to your fork: `git push origin feature/amazing-feature`
5. Open a Pull Request with description of changes

### Reporting Issues

If you find a bug or have a feature request:

1. Check existing issues first
2. Create a new issue with:
   - Clear title and description
   - Steps to reproduce (for bugs)
   - Expected vs actual behavior
   - Environment details (OS, Node version, etc.)
   - Screenshots/logs if applicable

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

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

## 🙏 Acknowledgments

- [Model Context Protocol (MCP)](https://modelcontextprotocol.io/) - The protocol specification
- [ExcelJS](https://github.com/exceljs/exceljs) - Excellent Excel library for Node.js
- [Zod](https://zod.dev/) - TypeScript-first schema validation
- The open-source community for inspiration and tools

## 📞 Support

- 📧 Email: support@example.com
- 💬 Issues: [GitHub Issues](https://github.com/ITKMUnigle/Excel-MCP-Tool/issues)
- 📖 Documentation: [Wiki](https://github.com/ITKMUnigle/Excel-MCP-Tool/wiki)

---

<div align="center">

**⭐ If this project helped you, please give it a star! ⭐**

Made with ❤️ by [ITKMUnigle](https://github.com/ITKMUnigle)

</div>

---

## 🌐 Other Languages

- **English Version (Current)** | [中文版本](./README.md)
