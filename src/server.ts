import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';

import { ExcelService } from './services/excel.service.js';
import {
  ReadExcelSchema,
  WriteExcelSchema,
  FormatCellSchema,
  SheetManagementSchema,
  GetFileInfoSchema
} from './types/index.js';

const excelService = new ExcelService();

export function createServer(): McpServer {
  const server = new McpServer({
    name: 'excel-mcp',
    version: '1.0.0',
  });

  // 注册读取工具
  server.tool(
    'read_excel',
    '读取Excel文件内容，支持指定工作表、范围、是否包含公式和格式信息',
    ReadExcelSchema.shape,
    async (args: any) => {
      try {
        const result = await excelService.readFile(args);
        return {
          content: [{
            type: 'text' as const,
            text: JSON.stringify(result, null, 2)
          }]
        };
      } catch (error) {
        return {
          content: [{
            type: 'text' as const,
            text: JSON.stringify({
              success: false,
              error: {
                code: 'UNKNOWN_ERROR',
                message: error instanceof Error ? error.message : String(error)
              }
            }, null, 2)
          }],
          isError: true
        };
      }
    }
  );

  // 注册写入工具
  server.tool(
    'write_excel',
    '向Excel文件写入数据，支持创建新文件、覆盖、追加或更新模式',
    WriteExcelSchema.shape,
    async (args: any) => {
      try {
        const result = await excelService.writeFile(args);
        return {
          content: [{
            type: 'text' as const,
            text: JSON.stringify(result, null, 2)
          }]
        };
      } catch (error) {
        return {
          content: [{
            type: 'text' as const,
            text: JSON.stringify({
              success: false,
              error: {
                code: 'WRITE_ERROR',
                message: error instanceof Error ? error.message : String(error)
              }
            }, null, 2)
          }],
          isError: true
        };
      }
    }
  );

  // 注册格式化工具
  server.tool(
    'format_excel',
    '格式化Excel单元格，包括字体样式、背景色、边框、对齐方式和数字格式等',
    FormatCellSchema.shape,
    async (args: any) => {
      try {
        const result = await excelService.formatCells(args);
        return {
          content: [{
            type: 'text' as const,
            text: JSON.stringify(result, null, 2)
          }]
        };
      } catch (error) {
        return {
          content: [{
            type: 'text' as const,
            text: JSON.stringify({
              success: false,
              error: {
                code: 'WRITE_ERROR',
                message: error instanceof Error ? error.message : String(error)
              }
            }, null, 2)
          }],
          isError: true
        };
      }
    }
  );

  // 注册工作表管理工具
  server.tool(
    'manage_sheets',
    '管理Excel工作簿中的工作表，包括列出、创建、删除、重命名、复制和移动操作',
    SheetManagementSchema.shape,
    async (args: any) => {
      try {
        const result = await excelService.manageSheets(args);
        return {
          content: [{
            type: 'text' as const,
            text: JSON.stringify(result, null, 2)
          }]
        };
      } catch (error) {
        return {
          content: [{
            type: 'text' as const,
            text: JSON.stringify({
              success: false,
              error: {
                code: 'UNKNOWN_ERROR',
                message: error instanceof Error ? error.message : String(error)
              }
            }, null, 2)
          }],
          isError: true
        };
      }
    }
  );

  // 注册获取文件信息工具
  server.tool(
    'get_file_info',
    '获取Excel文件的基本信息，包括文件大小、修改时间、工作表列表等',
    GetFileInfoSchema.shape,
    async (args: any) => {
      try {
        const result = await excelService.getFileInfo(args.filePath);
        return {
          content: [{
            type: 'text' as const,
            text: JSON.stringify(result, null, 2)
          }]
        };
      } catch (error) {
        return {
          content: [{
            type: 'text' as const,
            text: JSON.stringify({
              success: false,
              error: {
                code: 'UNKNOWN_ERROR',
                message: error instanceof Error ? error.message : String(error)
              }
            }, null, 2)
          }],
          isError: true
        };
      }
    }
  );

  return server;
}

export async function startServer(transportType: 'stdio' | 'http' = 'stdio'): Promise<void> {
  const server = createServer();

  if (transportType === 'stdio') {
    const transport = new StdioServerTransport();
    await server.connect(transport);
    console.error('Excel MCP Server running on stdio');
  } else {
    console.error('HTTP transport not yet implemented. Use stdio mode.');
    process.exit(1);
  }
}
