import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";
const path = "E:/融辉物流/网点返利/99-产品原型prototype/政策管理/政策管理导入模板.xlsx";
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(path));
const sheet = workbook.worksheets.getItem("政策导入模板");
for (const address of ["A2", "B2", "E2", "F2", "G2", "H2"]) {
  console.log(address, JSON.stringify(sheet.getRange(address).dataValidation));
}
const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!",
  options: { useRegex: true, maxResults: 100 },
  summary: "final formula error scan",
});
console.log(errors.ndjson);
