import { Tool } from '@modelcontextprotocol/sdk/types.js';
import { ExcelService } from '../services/excel.service.js';
import { SheetManagementSchema, SheetManagementParams, SheetManagementResult, ErrorResponse } from '../types/index.js';

const excelService = new ExcelService();

export const manageSheetsTool: Tool = {
  name: 'manage_sheets',
  description: '管理Excel工作簿中的工作表，包括列出、创建、删除、重命名、复制和移动操作',
  inputSchema: {
    type: 'object' as const,
    properties: {
      filePath: {
        type: 'string',
        description: 'Excel文件的完整路径'
      },
      action: {
        type: 'string',
        enum: ['list', 'create', 'delete', 'rename', 'copy', 'move'],
        description: '操作类型：list-列出所有工作表，create-创建新工作表，delete-删除工作表，rename-重命名工作表，copy-复制工作表，move-移动工作表位置'
      },
      sheetName: {
        type: 'string',
        description: '目标工作表名称（create/rename/delete/copy/move操作时需要）'
      },
      newName: {
        type: 'string',
        description: '新名称（rename操作时需要）'
      },
      index: {
        type: 'number',
        description: '目标位置索引（move操作时使用，从0开始）'
      }
    },
    required: ['filePath', 'action']
  }
};

export async function handleManageSheets(args: SheetManagementParams): Promise<SheetManagementResult | ErrorResponse> {
  try {
    const validatedParams = SheetManagementSchema.parse(args);
    const result = await excelService.manageSheets(validatedParams);
    return result;
  } catch (error) {
    if (error instanceof Error && 'issues' in error) {
      return {
        success: false,
        error: {
          code: 'VALIDATION_ERROR' as any,
          message: '参数验证失败',
          details: error
        }
      };
    }

    return {
      success: false,
      error: {
        code: 'UNKNOWN_ERROR' as any,
        message: `管理工作表时发生错误: ${error instanceof Error ? error.message : '未知错误'}`
      }
    };
  }
}

export const getFileInfoTool: Tool = {
  name: 'get_file_info',
  description: '获取Excel文件的基本信息，包括文件大小、修改时间、工作表列表等',
  inputSchema: {
    type: 'object' as const,
    properties: {
      filePath: {
        type: 'string',
        description: 'Excel文件的完整路径'
      }
    },
    required: ['filePath']
  }
};

export async function handleGetFileInfo(filePath: string): Promise<any> {
  try {
    const result = await excelService.getFileInfo(filePath);
    return result;
  } catch (error) {
    return {
      success: false,
      error: {
        code: 'UNKNOWN_ERROR' as any,
        message: `获取文件信息时发生错误: ${error instanceof Error ? error.message : '未知错误'}`
      }
    };
  }
}
