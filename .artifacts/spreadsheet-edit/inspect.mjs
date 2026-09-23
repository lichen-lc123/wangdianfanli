import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = "E:/融辉物流/网点返利/99-产品原型prototype/政策管理/政策管理导入模板.xlsx";
const outputDir = "E:/融辉物流/网点返利/.artifacts/spreadsheet-edit";
const input = await FileBlob.load(inputPath);
const workbook = await SpreadsheetFile.importXlsx(input);
const summary = await workbook.inspect({
  kind: "workbook,sheet,table,definedName",
  maxChars: 10000,
  tableMaxRows: 12,
  tableMaxCols: 24,
  tableMaxCellChars: 120,
});
console.log(summary.ndjson);
for (const sheet of workbook.worksheets.items) {
  const used = sheet.getUsedRange();
  if (!used) continue;
  const region = await workbook.inspect({
    kind: "region",
    sheetId: sheet.name,
    range: used.address,
    maxChars: 12000,
    tableMaxRows: 20,
    tableMaxCols: 24,
  });
  console.log(region.ndjson);
  const preview = await workbook.render({ sheetName: sheet.name, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(`${outputDir}/before-${sheet.name}.png`, new Uint8Array(await preview.arrayBuffer()));
}
