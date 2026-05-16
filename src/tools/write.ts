import { Tool } from '@modelcontextprotocol/sdk/types.js';
import { ExcelService } from '../services/excel.service.js';
import { WriteExcelSchema, WriteExcelParams, WriteExcelResult, ErrorResponse } from '../types/index.js';

const excelService = new ExcelService();

export const writeExcelTool: Tool = {
  name: 'write_excel',
  description: '向Excel文件写入数据，支持创建新文件、覆盖、追加或更新模式',
  inputSchema: {
    type: 'object' as const,
    properties: {
      filePath: {
        type: 'string',
        description: 'Excel文件的完整路径（如果不存在会自动创建）'
      },
      sheetName: {
        type: 'string',
        description: '目标工作表名称（可选，默认为"Sheet1"）'
      },
      data: {
        type: 'array',
        description: '要写入的数据（二维数组格式）',
        items: {
          type: 'array',
          items: {}
        }
      },
      startCell: {
        type: 'string',
        description: '起始单元格位置（可选，默认"A1"）'
      },
      mode: {
        type: 'string',
        enum: ['overwrite', 'append', 'update'],
        description: '写入模式：overwrite-覆盖现有数据，append-追加到末尾，update-从指定位置更新（可选，默认"overwrite"）'
      },
      createIfNotExists: {
        type: 'boolean',
        description: '文件不存在时是否自动创建（可选，默认true）'
      },
      headers: {
        type: 'array',
        description: '表头数组（可选，仅在overwrite模式下有效）',
        items: {
          type: 'string'
        }
      }
    },
    required: ['filePath', 'data']
  }
};

export async function handleWriteExcel(args: WriteExcelParams): Promise<WriteExcelResult | ErrorResponse> {
  try {
    const validatedParams = WriteExcelSchema.parse(args);
    const result = await excelService.writeFile(validatedParams);
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
        code: 'WRITE_ERROR' as any,
        message: `写入Excel时发生错误: ${error instanceof Error ? error.message : '未知错误'}`
      }
    };
  }
}
