import * as fs from 'fs';
import * as path from 'path';
import { ErrorCode, ErrorDetails } from '../types/index.js';
import { ERROR_MESSAGES, SUPPORTED_EXTENSIONS, MAX_FILE_SIZE } from './constants.js';

export function createError(code: ErrorCode, message?: string, details?: any): ErrorDetails {
  return {
    code,
    message: message || ERROR_MESSAGES[code] || '未知错误',
    details
  };
}

export function validateFilePath(filePath: string): { valid: boolean; error?: ErrorDetails } {
  if (!filePath || typeof filePath !== 'string') {
    return { valid: false, error: createError(ErrorCode.VALIDATION_ERROR, '文件路径不能为空') };
  }

  const ext = path.extname(filePath).toLowerCase();
  if (!SUPPORTED_EXTENSIONS.includes(ext)) {
    return {
      valid: false,
      error: createError(ErrorCode.INVALID_FORMAT, `不支持的文件格式: ${ext}，仅支持 ${SUPPORTED_EXTENSIONS.join(', ')}`)
    };
  }

  return { valid: true };
}

export function validateFileExists(filePath: string): { valid: boolean; error?: ErrorDetails } {
  if (!fs.existsSync(filePath)) {
    return { valid: false, error: createError(ErrorCode.FILE_NOT_FOUND, `文件不存在: ${filePath}`) };
  }

  try {
    const stats = fs.statSync(filePath);
    if (stats.size > MAX_FILE_SIZE) {
      return {
        valid: false,
        error: createError(ErrorCode.VALIDATION_ERROR, `文件过大，最大支持 ${MAX_FILE_SIZE / 1024 / 1024}MB`)
      };
    }
  } catch (error) {
    return { valid: false, error: createError(ErrorCode.PERMISSION_DENIED, '无法读取文件信息') };
  }

  return { valid: true };
}

export function validateCellRange(range: string): { valid: boolean; error?: ErrorDetails } {
  if (!range || typeof range !== 'string') {
    return { valid: false, error: createError(ErrorCode.VALIDATION_ERROR, '单元格范围不能为空') };
  }

  const rangePattern = /^[A-Z]+\d+:[A-Z]+\d+$/;
  if (!rangePattern.test(range)) {
    return { valid: false, error: createError(ErrorCode.INVALID_RANGE, `无效的单元格范围格式: ${range}，正确格式如 A1:C10`) };
  }

  return { valid: true };
}

export function parseRange(range: string): { startRow: number; startCol: number; endRow: number; endCol: number } | null {
  const match = range.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
  if (!match) return null;

  const colToIndex = (col: string) => {
    let index = 0;
    for (let i = 0; i < col.length; i++) {
      index = index * 26 + (col.charCodeAt(i) - 64);
    }
    return index - 1;
  };

  return {
    startRow: parseInt(match[2]) - 1,
    startCol: colToIndex(match[1]),
    endRow: parseInt(match[4]) - 1,
    endCol: colToIndex(match[3])
  };
}

export function indexToCol(index: number): string {
  let col = '';
  index++;
  while (index > 0) {
    const remainder = (index - 1) % 26;
    col = String.fromCharCode(65 + remainder) + col;
    index = Math.floor((index - 1) / 26);
  }
  return col;
}

export function getFileName(filePath: string): string {
  return path.basename(filePath);
}

export function getFileSize(bytes: number): string {
  const units = ['B', 'KB', 'MB', 'GB'];
  let unitIndex = 0;
  let size = bytes;

  while (size >= 1024 && unitIndex < units.length - 1) {
    size /= 1024;
    unitIndex++;
  }

  return `${size.toFixed(2)} ${units[unitIndex]}`;
}
