import fs from 'node:fs/promises';
import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';

const inputPath = 'E:/融辉物流/网点返利/99-产品原型prototype/政策管理/政策管理导入模板.xlsx';
const outputPath = 'E:/融辉物流/网点返利/99-产品原型prototype/政策管理/政策管理导入模板_四档公斤段.xlsx';
const previewPath = 'E:/融辉物流/网点返利/.artifact_work/政策管理导入模板_修改后.png';
const input = await FileBlob.load(inputPath);
const workbook = await SpreadsheetFile.importXlsx(input);
const main = workbook.worksheets.getItem('政策导入模板');
const notes = workbook.worksheets.getItem('填写说明');
const enums = workbook.worksheets.getItem('枚举值');

const commaJoin = (value) => typeof value === 'string' ? value.replaceAll('|', ',') : value;
const normalizeCargo = (value) => typeof value === 'string' ? commaJoin(value).replaceAll('轻抛', '轻泡').replaceAll('重抛', '重泡') : value;

// 主表样例：所有多选值使用英文逗号。
main.getRange('E2:G2').values = [[
  '融汇达,精准零担',
  '轻泡,重泡,重货',
  '0kg-100kg,100kg-300kg,300kg-800kg,800kg-1500kg',
]];

// 保留既有产品、货物类型组合，只替换多选分隔符。
const productValues = enums.getRange('A2:A16').values.map(([value]) => [commaJoin(value)]);
const cargoValues = enums.getRange('B2:B8').values.map(([value]) => [normalizeCargo(value)]);
enums.getRange('A2:A16').values = productValues;
enums.getRange('B2:B8').values = cargoValues;

// 公斤段固定为四档；“全部”之外提供所有单选与组合。
const bands = ['0kg-100kg', '100kg-300kg', '300kg-800kg', '800kg-1500kg'];
const bandOptions = ['全部'];
for (let size = 1; size <= bands.length; size += 1) {
  const choose = (start, picked) => {
    if (picked.length === size) {
      bandOptions.push(picked.join(','));
      return;
    }
    for (let index = start; index < bands.length; index += 1) choose(index + 1, [...picked, bands[index]]);
  };
  choose(0, []);
}
enums.getRange('C2:C128').clear({ applyTo: 'contents' });
enums.getRange(`C2:C${bandOptions.length + 1}`).values = bandOptions.map((value) => [value]);
enums.getRange('E2').values = [['“全部”表示当前列全部选项；多选组合使用英文逗号（,）分隔。']];

// 填写说明与拆分示例同步。
notes.getRange('G8:G10').values = [[
  '融汇达,精准零担',
], [
  '轻泡,重泡,重货',
], [
  '0kg-100kg,100kg-300kg,300kg-800kg,800kg-1500kg',
]];
notes.getRange('D12').values = [['仅可选择0kg-100kg、100kg-300kg、300kg-800kg、800kg-1500kg；支持全部、单选和预组合多选']];
notes.getRange('D11').values = [['下拉可选择全部、单个或预组合的多个货物类型：轻泡、重泡、重货']];
notes.getRange('F15').values = [['2. 选择字段必须使用下拉选项；“全部”由系统展开，多选组合使用英文逗号（,）分隔']];

// 公斤段下拉缩减为新的四档组合列表。
main.getRange('G2:G200000').dataValidation = {
  rule: { type: 'list', formula1: "'枚举值'!$C$2:$C$17" },
};

workbook.recalculate();

const checkMain = await workbook.inspect({
  kind: 'table',
  sheetId: '政策导入模板',
  range: 'A1:P2',
  include: 'values,formulas',
  tableMaxRows: 5,
  tableMaxCols: 16,
  maxChars: 8000,
});
console.log(checkMain.ndjson);
const checkEnums = await workbook.inspect({
  kind: 'table',
  sheetId: '枚举值',
  range: 'A1:E18',
  include: 'values,formulas',
  tableMaxRows: 20,
  tableMaxCols: 5,
  maxChars: 10000,
});
console.log(checkEnums.ndjson);
const errors = await workbook.inspect({
  kind: 'match',
  searchTerm: '#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',
  options: { useRegex: true, maxResults: 100 },
  summary: 'final formula error scan',
});
console.log(errors.ndjson);

const preview = await workbook.render({ sheetName: '政策导入模板', range: 'A1:P12', scale: 1.5, format: 'png' });
await fs.writeFile(previewPath, new Uint8Array(await preview.arrayBuffer()));
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log('SAVED', outputPath);
console.log('PREVIEW', previewPath);
