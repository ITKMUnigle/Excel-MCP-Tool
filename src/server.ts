import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';

import {
  readExcelTool,
  handleReadExcel,
  writeExcelTool,
  handleWriteExcel,
  formatExcelTool,
  handleFormatExcel,
  manageSheetsTool,
  handleManageSheets,
  getFileInfoTool,
  handleGetFileInfo
} from './tools/index.js';

export function createServer(): McpServer {
  const server = new McpServer({
    name: 'excel-mcp',
    version: '1.0.0',
  });

  // 注册读取工具
  server.tool(
    readExcelTool.name,
    readExcelTool.description || '读取Excel文件内容',
    readExcelTool.inputSchema.properties as any,
    async (args: any) => {
      const result = await handleReadExcel(args as any);
      return {
        content: [{
          type: 'text' as const,
          text: JSON.stringify(result, null, 2)
        }]
      };
    }
  );

  // 注册写入工具
  server.tool(
    writeExcelTool.name,
    writeExcelTool.description || '向Excel文件写入数据',
    writeExcelTool.inputSchema.properties as any,
    async (args: any) => {
      const result = await handleWriteExcel(args as any);
      return {
        content: [{
          type: 'text' as const,
          text: JSON.stringify(result, null, 2)
        }]
      };
    }
  );

  // 注册格式化工具
  server.tool(
    formatExcelTool.name,
    formatExcelTool.description || '格式化Excel单元格',
    formatExcelTool.inputSchema.properties as any,
    async (args: any) => {
      const result = await handleFormatExcel(args as any);
      return {
        content: [{
          type: 'text' as const,
          text: JSON.stringify(result, null, 2)
        }]
      };
    }
  );

  // 注册工作表管理工具
  server.tool(
    manageSheetsTool.name,
    manageSheetsTool.description || '管理工作表',
    manageSheetsTool.inputSchema.properties as any,
    async (args: any) => {
      const result = await handleManageSheets(args as any);
      return {
        content: [{
          type: 'text' as const,
          text: JSON.stringify(result, null, 2)
        }]
      };
    }
  );

  // 注册获取文件信息工具
  server.tool(
    getFileInfoTool.name,
    getFileInfoTool.description || '获取文件信息',
    getFileInfoTool.inputSchema.properties as any,
    async (args: any) => {
      const result = await handleGetFileInfo((args as any).filePath);
      return {
        content: [{
          type: 'text' as const,
          text: JSON.stringify(result, null, 2)
        }]
      };
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
