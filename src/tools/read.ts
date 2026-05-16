import { Tool } from '@modelcontextprotocol/sdk/types.js';
import { ExcelService } from '../services/excel.service.js';
import { ReadExcelSchema, ReadExcelParams, ReadExcelResult, ErrorResponse } from '../types/index.js';

const excelService = new ExcelService();

export const readExcelTool: Tool = {
  name: 'read_excel',
  description: '读取Excel文件内容，支持指定工作表、范围、是否包含公式和格式信息',
  inputSchema: {
    type: 'object' as const,
    properties: {
      filePath: {
        type: 'string',
        description: 'Excel文件的完整路径（支持.xlsx和.xls格式）'
      },
      sheetName: {
        type: 'string',
        description: '要读取的工作表名称（可选，默认读取第一个工作表）'
      },
      range: {
        type: 'string',
        description: '要读取的单元格范围（可选，如"A1:C10"，默认读取全部数据）'
      },
      includeFormulas: {
        type: 'boolean',
        description: '是否包含单元格公式（可选，默认false）'
      },
      includeFormatting: {
        type: 'boolean',
        description: '是否包含格式化信息（可选，默认false）'
      }
    },
    required: ['filePath']
  }
};

export async function handleReadExcel(args: ReadExcelParams): Promise<ReadExcelResult | ErrorResponse> {
  try {
    const validatedParams = ReadExcelSchema.parse(args);
    const result = await excelService.readFile(validatedParams);
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
        message: `读取Excel时发生错误: ${error instanceof Error ? error.message : '未知错误'}`
      }
    };
  }
}
