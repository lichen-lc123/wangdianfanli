import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";
const path = "E:/融辉物流/网点返利/99-产品原型prototype/政策管理/政策管理导入模板.xlsx";
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(path));
const sheet = workbook.worksheets.getItem("政策导入模板");
console.log(Array.isArray(sheet.dataValidations.items), sheet.dataValidations.items?.length ?? "no-items");
