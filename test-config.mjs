import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { ExcelService } from './dist/services/excel.service.js';

const truthyValues = ['true', 'TRUE', 'yes', 'YES', 'y', 'Y'];
const falsyValues = ['false', 'FALSE', 'no', 'NO', 'n', 'N'];

const tempDir = await mkdtemp(join(tmpdir(), 'excel-mcp-'));

try {
  for (const value of falsyValues) {
    process.env.EXCEL_ALLOW_WRITE = value;
    const service = new ExcelService();
    const result = await service.writeFile({
      filePath: join(tempDir, `blocked-${value}.xlsx`),
      data: [['blocked']],
      createIfNotExists: true,
      mode: 'overwrite'
    });

    assert.equal(result.success, false, `${value} should disable writes`);
    assert.equal(result.error?.code, 'PERMISSION_DENIED');
  }

  delete process.env.EXCEL_ALLOW_WRITE;
  const defaultService = new ExcelService();
  const defaultResult = await defaultService.writeFile({
    filePath: join(tempDir, 'blocked-default.xlsx'),
    data: [['blocked']],
    createIfNotExists: true,
    mode: 'overwrite'
  });

  assert.equal(defaultResult.success, false, 'missing EXCEL_ALLOW_WRITE should disable writes');
  assert.equal(defaultResult.error?.code, 'PERMISSION_DENIED');

  for (const value of truthyValues) {
    process.env.EXCEL_ALLOW_WRITE = value;
    const service = new ExcelService();
    const result = await service.writeFile({
      filePath: join(tempDir, `allowed-${value}.xlsx`),
      data: [['allowed']],
      createIfNotExists: true,
      mode: 'overwrite'
    });

    assert.equal(result.success, true, `${value} should enable writes`);
  }

  const entrypoint = await readFile('./dist/index.js', 'utf8');
  assert.equal(entrypoint.split('\n')[0], '#!/usr/bin/env node', 'bin entrypoint should have a node shebang');
} finally {
  await rm(tempDir, { recursive: true, force: true });
}
