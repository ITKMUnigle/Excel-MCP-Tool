export const ERROR_MESSAGES: Record<string, string> = {
  FILE_NOT_FOUND: '文件不存在',
  INVALID_FORMAT: '无效的Excel格式',
  SHEET_NOT_FOUND: '工作表不存在',
  INVALID_RANGE: '无效的单元格范围',
  PERMISSION_DENIED: '权限不足，无法访问该文件',
  WRITE_ERROR: '写入文件失败',
  VALIDATION_ERROR: '参数验证失败',
  UNKNOWN_ERROR: '未知错误'
};

export const SUPPORTED_EXTENSIONS = ['.xlsx', '.xls'];
export const MAX_FILE_SIZE = 100 * 1024 * 1024; // 100MB
export const DEFAULT_SHEET_NAME = 'Sheet1';

export const NUMBER_FORMATS = {
  CURRENCY_CN: '¥#,##0.00',
  CURRENCY_US: '$#,##0.00',
  PERCENTAGE: '0.00%',
  DATE: 'yyyy-mm-dd',
  DATETIME: 'yyyy-mm-dd hh:mm:ss',
  NUMBER: '#,##0.00',
  INTEGER: '#,##0',
  SCIENTIFIC: '0.00E+00'
} as const;

export const FONT_NAMES = [
  '微软雅黑',
  '宋体',
  '黑体',
  'Arial',
  'Times New Roman',
  'Calibri',
  'Helvetica'
] as const;

export const BORDER_STYLES = ['thin', 'medium', 'thick', 'dashed', 'dotted', 'double'] as const;
