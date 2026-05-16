import { z } from 'zod';

// ==================== 读取参数 ====================
export const ReadExcelSchema = z.object({
  filePath: z.string().min(1, '文件路径不能为空'),
  sheetName: z.string().optional(),
  range: z.string().optional(),
  includeFormulas: z.boolean().default(false),
  includeFormatting: z.boolean().default(false)
});

export type ReadExcelParams = z.infer<typeof ReadExcelSchema>;

export interface ReadExcelResult {
  success: boolean;
  fileName: string;
  sheetName: string;
  data: Array<Array<any>>;
  headers?: string[];
  rowCount: number;
  columnCount: number;
  formulas?: Record<string, string>;
  formatting?: CellFormatting[];
}

// ==================== 写入参数 ====================
export const WriteExcelSchema = z.object({
  filePath: z.string().min(1, '文件路径不能为空'),
  sheetName: z.string().optional(),
  data: z.array(z.array(z.any())).min(1, '数据不能为空'),
  startCell: z.string().default('A1'),
  mode: z.enum(['overwrite', 'append', 'update']).default('overwrite'),
  createIfNotExists: z.boolean().default(true),
  headers: z.array(z.string()).optional()
});

export type WriteExcelParams = z.infer<typeof WriteExcelSchema>;

export interface WriteExcelResult {
  success: boolean;
  filePath: string;
  sheetName: string;
  rowsWritten: number;
  message: string;
}

// ==================== 格式化参数 ====================
export const BorderStyleSchema = z.object({
  style: z.enum(['thin', 'medium', 'thick', 'dashed', 'dotted', 'double']).default('thin'),
  color: z.string().optional()
});

export const FontSchema = z.object({
  name: z.string().optional(),
  size: z.number().positive().optional(),
  bold: z.boolean().optional(),
  italic: z.boolean().optional(),
  color: z.string().optional()
});

export const FillSchema = z.object({
  type: z.enum(['solid', 'gradient']).default('solid'),
  color: z.string().optional()
});

export const AlignmentSchema = z.object({
  horizontal: z.enum(['left', 'center', 'right']).optional(),
  vertical: z.enum(['top', 'middle', 'bottom']).optional(),
  wrapText: z.boolean().optional()
});

export const FormatCellSchema = z.object({
  filePath: z.string().min(1, '文件路径不能为空'),
  sheetName: z.string().optional(),
  range: z.string().min(1, '单元格范围不能为空'),
  font: FontSchema.optional(),
  fill: FillSchema.optional(),
  border: z.object({
    top: BorderStyleSchema.optional(),
    bottom: BorderStyleSchema.optional(),
    left: BorderStyleSchema.optional(),
    right: BorderStyleSchema.optional()
  }).optional(),
  alignment: AlignmentSchema.optional(),
  numberFormat: z.string().optional()
});

export type FormatCellParams = z.infer<typeof FormatCellSchema>;
export type BorderStyle = z.infer<typeof BorderStyleSchema>;
export type FontStyle = z.infer<typeof FontSchema>;
export type FillStyle = z.infer<typeof FillSchema>;
export type AlignmentStyle = z.infer<typeof AlignmentSchema>;

export interface FormatCellResult {
  success: boolean;
  filePath: string;
  range: string;
  cellsFormatted: number;
  message: string;
}

// ==================== 工作表管理参数 ====================
export const SheetManagementSchema = z.object({
  filePath: z.string().min(1, '文件路径不能为空'),
  action: z.enum(['list', 'create', 'delete', 'rename', 'copy', 'move']),
  sheetName: z.string().optional(),
  newName: z.string().optional(),
  index: z.number().int().min(0).optional()
});

export type SheetManagementParams = z.infer<typeof SheetManagementSchema>;
export type SheetAction = SheetManagementParams['action'];

export interface SheetInfo {
  name: string;
  id: number;
  rowCount: number;
  columnCount: number;
}

export interface SheetManagementResult {
  success: boolean;
  action: SheetAction;
  sheets?: SheetInfo[];
  message: string;
}

// ==================== 单元格格式信息 ====================
export interface CellFormatting {
  cell: string;
  font?: FontStyle;
  fill?: FillStyle;
  border?: {
    top?: BorderStyle;
    bottom?: BorderStyle;
    left?: BorderStyle;
    right?: BorderStyle;
  };
  alignment?: AlignmentStyle;
  numberFormat?: string;
}

// ==================== 文件信息 ====================
export const GetFileInfoSchema = z.object({
  filePath: z.string().min(1, '文件路径不能为空')
});

export type GetFileInfoParams = z.infer<typeof GetFileInfoSchema>;

export interface FileInfo {
  success: boolean;
  fileName: string;
  filePath: string;
  fileSize: number;
  lastModified: Date;
  sheetCount: number;
  sheets: string[];
  isValid: boolean;
}

// ==================== 错误类型 ====================
export enum ErrorCode {
  FILE_NOT_FOUND = 'FILE_NOT_FOUND',
  INVALID_FORMAT = 'INVALID_FORMAT',
  SHEET_NOT_FOUND = 'SHEET_NOT_FOUND',
  INVALID_RANGE = 'INVALID_RANGE',
  PERMISSION_DENIED = 'PERMISSION_DENIED',
  WRITE_ERROR = 'WRITE_ERROR',
  VALIDATION_ERROR = 'VALIDATION_ERROR',
  UNKNOWN_ERROR = 'UNKNOWN_ERROR'
}

export interface ErrorDetails {
  code: ErrorCode;
  message: string;
  details?: any;
}

export interface ErrorResponse {
  success: false;
  error: ErrorDetails;
}
