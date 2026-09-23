import fs from 'node:fs/promises';
import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';

const inputPath = 'E:/融辉物流/网点返利/99-产品原型prototype/政策管理/政策管理导入模板.xlsx';
const previewPath = 'E:/融辉物流/网点返利/.artifact_work/政策管理导入模板_修改前.png';
const input = await FileBlob.load(inputPath);
const workbook = await SpreadsheetFile.importXlsx(input);
const summary = await workbook.inspect({
  kind: 'workbook,sheet,table',
  maxChars: 10000,
  tableMaxRows: 20,
  tableMaxCols: 24,
  tableMaxCellChars: 120,
});
console.log(summary.ndjson);
const sheet = workbook.worksheets.getItemAt(0);
console.log('SHEET', sheet.name);
console.log('USED_VALUES', JSON.stringify(sheet.getUsedRange().values));
const enumSheet = workbook.worksheets.getItem('枚举值');
console.log('ENUM_VALUES', JSON.stringify(enumSheet.getRange('A1:E140').values));
const noteSheet = workbook.worksheets.getItem('填写说明');
console.log('NOTES_VALUES', JSON.stringify(noteSheet.getRange('A1:I21').values));
const validationInfo = await workbook.inspect({
  kind: 'region',
  sheetId: sheet.name,
  range: 'A1:R12',
  include: 'values,formulas',
  maxChars: 8000,
});
console.log(validationInfo.ndjson);
const preview = await workbook.render({ sheetName: sheet.name, range: 'A1:R12', scale: 1.5, format: 'png' });
await fs.writeFile(previewPath, new Uint8Array(await preview.arrayBuffer()));
console.log('PREVIEW', previewPath);
