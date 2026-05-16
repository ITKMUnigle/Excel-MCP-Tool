import { Tool } from '@modelcontextprotocol/sdk/types.js';
import { ExcelService } from '../services/excel.service.js';
import { FormatCellSchema, FormatCellParams, FormatCellResult, ErrorResponse } from '../types/index.js';

const excelService = new ExcelService();

export const formatExcelTool: Tool = {
  name: 'format_excel',
  description: '格式化Excel单元格，包括字体样式、背景色、边框、对齐方式和数字格式等',
  inputSchema: {
    type: 'object' as const,
    properties: {
      filePath: {
        type: 'string',
        description: 'Excel文件的完整路径'
      },
      sheetName: {
        type: 'string',
        description: '目标工作表名称（可选，默认第一个工作表）'
      },
      range: {
        type: 'string',
        description: '要格式化的单元格范围（如"A1:F1"、"B2:B100"）'
      },
      font: {
        type: 'object',
        description: '字体设置',
        properties: {
          name: {
            type: 'string',
            description: '字体名称（如"微软雅黑"、"Arial"）'
          },
          size: {
            type: 'number',
            description: '字体大小（如12、14、16）'
          },
          bold: {
            type: 'boolean',
            description: '是否加粗'
          },
          italic: {
            type: 'boolean',
            description: '是否斜体'
          },
          color: {
            type: 'string',
            description: '字体颜色（ARGB格式，如"FF000000"表示黑色）'
          }
        }
      },
      fill: {
        type: 'object',
        description: '背景填充设置',
        properties: {
          type: {
            type: 'string',
            enum: ['solid', 'gradient'],
            description: '填充类型（目前仅支持solid）'
          },
          color: {
            type: 'string',
            description: '背景颜色（ARGB格式，如"FF4472C4"表示蓝色）'
          }
        }
      },
      border: {
        type: 'object',
        description: '边框设置',
        properties: {
          top: {
            type: 'object',
            description: '上边框',
            properties: {
              style: {
                type: 'string',
                enum: ['thin', 'medium', 'thick', 'dashed', 'dotted', 'double'],
                description: '边框样式'
              },
              color: {
                type: 'string',
                description: '边框颜色（ARGB格式）'
              }
            }
          },
          bottom: {
            type: 'object',
            description: '下边框',
            properties: {
              style: {
                type: 'string',
                enum: ['thin', 'medium', 'thick', 'dashed', 'dotted', 'double'],
                description: '边框样式'
              },
              color: {
                type: 'string',
                description: '边框颜色（ARGB格式）'
              }
            }
          },
          left: {
            type: 'object',
            description: '左边框',
            properties: {
              style: {
                type: 'string',
                enum: ['thin', 'medium', 'thick', 'dashed', 'dotted', 'double'],
                description: '边框样式'
              },
              color: {
                type: 'string',
                description: '边框颜色（ARGB格式）'
              }
            }
          },
          right: {
            type: 'object',
            description: '右边框',
            properties: {
              style: {
                type: 'string',
                enum: ['thin', 'medium', 'thick', 'dashed', 'dotted', 'double'],
                description: '边框样式'
              },
              color: {
                type: 'string',
                description: '边框颜色（ARGB格式）'
              }
            }
          }
        }
      },
      alignment: {
        type: 'object',
        description: '对齐方式设置',
        properties: {
          horizontal: {
            type: 'string',
            enum: ['left', 'center', 'right'],
            description: '水平对齐方式'
          },
          vertical: {
            type: 'string',
            enum: ['top', 'middle', 'bottom'],
            description: '垂直对齐方式'
          },
          wrapText: {
            type: 'boolean',
            description: '是否自动换行'
          }
        }
      },
      numberFormat: {
        type: 'string',
        description: '数字格式（如"¥#,##0.00"表示货币、"0.00%"表示百分比）'
      }
    },
    required: ['filePath', 'range']
  }
};

export async function handleFormatExcel(args: FormatCellParams): Promise<FormatCellResult | ErrorResponse> {
  try {
    const validatedParams = FormatCellSchema.parse(args);
    const result = await excelService.formatCells(validatedParams);
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
        message: `格式化Excel时发生错误: ${error instanceof Error ? error.message : '未知错误'}`
      }
    };
  }
}
