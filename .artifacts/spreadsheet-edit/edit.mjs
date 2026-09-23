import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputPath = "E:/融辉物流/网点返利/99-产品原型prototype/政策管理/政策管理导入模板.xlsx";
const tempPath = "E:/融辉物流/网点返利/.artifacts/spreadsheet-edit/政策管理导入模板-待验证.xlsx";
const previewDir = "E:/融辉物流/网点返利/.artifacts/spreadsheet-edit";
const orange = "#C2410C";
const orange2 = "#EA580C";
const pale = "#FFF7ED";
const line = "#E7E5E4";
const text = "#292524";
const font = "Arial";

function choices(values) {
  const result = ["全部"];
  const n = values.length;
  for (let size = 1; size < n; size += 1) {
    for (let mask = 1; mask < (1 << n); mask += 1) {
      if (mask.toString(2).replaceAll("0", "").length !== size) continue;
      result.push(values.filter((_, i) => mask & (1 << i)).join("|"));
    }
  }
  return result;
}

const products = choices(["融汇达", "精准零担", "融安达", "融速达"]);
const cargos = choices(["轻抛", "重抛", "重货"]);
const bands = choices(["0kg-100kg", "100kg-300kg", "300kg-800kg", "800kg-1500kg", "1500kg-3000kg", "3T-10T", "10T以上"]);

const workbook = Workbook.create();
const main = workbook.worksheets.add("政策导入模板");
const guide = workbook.worksheets.add("填写说明");
const enums = workbook.worksheets.add("枚举值");

const headers = ["生效开始日期", "生效结束日期", "寄件网点", "目的流向", "产品类型", "货物类型", "公斤段", "考核周期", "日基准量", "日返利单价", "工作日基准量", "工作日返利单价", "周日基准量", "周日返利单价", "节假日基准量", "节假日返利单价"];
main.getRange("A1:P2").values = [
  headers,
  [new Date("2026-09-15T00:00:00"), new Date("2026-10-31T00:00:00"), "杭州临平网点", "杭州分拨中心", "融汇达|精准零担", "轻抛|重抛", "0kg-100kg|100kg-300kg|300kg-800kg", "周", 500, 0.01, 650, 0.012, 600, 0.016, 700, 0.022],
];
main.getRange("A1:P1").format = {
  fill: orange,
  font: { name: font, size: 10, bold: true, color: "#FFFFFF" },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  wrapText: true,
  borders: { preset: "all", style: "thin", color: "#FFFFFF" },
  rowHeight: 32,
};
main.getRange("A2:P2").format = {
  font: { name: font, size: 10, color: text },
  verticalAlignment: "center",
  borders: { preset: "all", style: "thin", color: line },
  rowHeight: 24,
};
main.getRange("A2:B2").format.numberFormat = "yyyy-mm-dd";
main.getRange("I2:P2").format.numberFormat = "0.0000";
for (const c of ["A", "B"]) main.getRange(`${c}:${c}`).format.columnWidth = 16;
for (const c of ["C", "D"]) main.getRange(`${c}:${c}`).format.columnWidth = 20;
for (const c of ["E", "F", "H"]) main.getRange(`${c}:${c}`).format.columnWidth = 18;
main.getRange("G:G").format.columnWidth = 42;
for (const c of ["I", "J", "K", "L", "M", "N", "O", "P"]) main.getRange(`${c}:${c}`).format.columnWidth = 18;
main.freezePanes.freezeRows(1);

const maxRows = 200000;
const dateValidation = {
  allowBlank: false,
  rule: { type: "date", operator: "between", formula1: "=DATE(2020,1,1)", formula2: "=DATE(2099,12,31)" },
  prompt: { title: "选择日期", message: "请选择或输入有效日期，格式为yyyy-mm-dd。" },
  errorAlert: { style: "stop", title: "日期无效", message: "请输入2020-01-01至2099-12-31之间的有效日期。" },
};
main.dataValidations.add({ range: `A2:A${maxRows}`, ...dateValidation });
main.dataValidations.add({ range: `B2:B${maxRows}`, ...dateValidation });
function addList(address, formula1, title, message) {
  main.dataValidations.add({
    range: address,
    allowBlank: false,
    rule: { type: "list", formula1 },
    prompt: { title, message },
    errorAlert: { style: "stop", title: "选项无效", message: "请从下拉列表选择，不要自行输入未定义值。" },
  });
}
addList(`E2:E${maxRows}`, `='枚举值'!$A$2:$A$${products.length + 1}`, "选择产品类型", "可选择全部、单个或预组合的多个产品类型。");
addList(`F2:F${maxRows}`, `='枚举值'!$B$2:$B$${cargos.length + 1}`, "选择货物类型", "可选择全部、单个或预组合的多个货物类型。");
addList(`G2:G${maxRows}`, `='枚举值'!$C$2:$C$${bands.length + 1}`, "选择公斤段", "可选择全部、单个或预组合的多个公斤段。");
addList(`H2:H${maxRows}`, `='枚举值'!$D$2:$D$3`, "选择考核周期", "考核周期只能单选周或月。");

guide.showGridLines = false;
guide.getRange("A2:D2").merge();
guide.getRange("A2").values = [["网点返利政策导入填写说明"]];
guide.getRange("A2:D2").format = { font: { name: font, size: 15, bold: true, color: text }, verticalAlignment: "center", rowHeight: 30 };
guide.getRange("A3:I3").merge();
guide.getRange("A3").values = [["日期字段使用日期校验；产品类型、货物类型和公斤段使用下拉列表，可选择全部、单个或预组合的多个值；考核周期只能单选。"]];
guide.getRange("A3:I3").format = { font: { name: font, size: 10, italic: true, color: "#78716C" }, rowHeight: 24 };

const rules = [
  ["字段名称", "是否必填", "填写格式", "填写规则"],
  ["生效开始日期", "是", "日期选择", "选择或输入政策生效起始日期，不得晚于结束日期"],
  ["生效结束日期", "是", "日期选择", "选择或输入政策生效结束日期；允许跨月并由系统按自然月拆分"],
  ["寄件网点", "是", "文本", "填写系统中的标准网点名称"],
  ["目的流向", "是", "文本", "填写分拨或营运区名称；营运区覆盖导入时快照内的分拨"],
  ["产品类型", "是", "列表选择", "下拉可选择全部、单个或预组合的多个产品类型"],
  ["货物类型", "是", "列表选择", "下拉可选择全部、单个或预组合的多个货物类型：轻抛、重抛、重货"],
  ["公斤段", "是", "列表选择", "下拉可选择全部、单个或预组合的多个公斤段"],
  ["考核周期", "是", "单选列表", "只能单选周或月"],
  ["日基准量", "是", "数值（kg）", "大于或等于0"],
  ["日返利单价", "是", "数值（元/kg）", "大于或等于0，建议保留4位小数"],
  ["工作日基准量", "否", "数值（kg）", "为空时继承日基准量"],
  ["工作日返利单价", "否", "数值（元/kg）", "为空时继承日返利单价"],
  ["周日基准量", "否", "数值（kg）", "为空时继承日基准量"],
  ["周日返利单价", "否", "数值（元/kg）", "为空时继承日返利单价"],
  ["节假日基准量", "否", "数值（kg）", "为空时按当天星期继承工作日或周日基准量"],
  ["节假日返利单价", "否", "数值（元/kg）", "为空时按当天星期继承工作日或周日返利单价"],
];
guide.getRange("A5").write(rules);
guide.getRange("A5:D5").format = { fill: orange, font: { name: font, size: 10, bold: true, color: "#FFFFFF" }, horizontalAlignment: "center", verticalAlignment: "center", rowHeight: 28 };
guide.getRange("A6:D21").format = { font: { name: font, size: 10, color: text }, verticalAlignment: "center", borders: { preset: "inside", style: "thin", color: line } };
guide.getRange("A6:D15").format.fill = pale;

guide.getRange("F5:I5").merge();
guide.getRange("F5").values = [["系统拆分规则示例"]];
guide.getRange("F5:I5").format = { fill: orange2, font: { name: font, size: 10, bold: true, color: "#FFFFFF" }, horizontalAlignment: "left", verticalAlignment: "center", rowHeight: 28 };
guide.getRange("F6:I11").values = [
  ["输入项目", "示例", "数量", "说明"],
  ["日期范围", "2026-09-15 至 2026-10-31", 2, "跨2个自然月"],
  ["产品类型", "融汇达|精准零担", 2, "选择2个产品类型"],
  ["货物类型", "轻抛|重抛", 2, "选择2个货物类型"],
  ["公斤段", "0kg-100kg|100kg-300kg|300kg-800kg", 3, "选择3个公斤段"],
  ["拆分结果", "2 × 2 × 2 × 3", 24, "导入系统后形成24条月度配置"],
];
guide.getRange("F6:I6").format = { fill: "#FFEDD5", font: { name: font, size: 10, bold: true, color: orange }, horizontalAlignment: "center" };
guide.getRange("F7:I11").format = { font: { name: font, size: 10, color: text }, borders: { preset: "inside", style: "thin", color: line } };

guide.getRange("F13:I13").merge();
guide.getRange("F13").values = [["导入校验"]];
guide.getRange("F13:I13").format = { fill: orange2, font: { name: font, size: 10, bold: true, color: "#FFFFFF" }, horizontalAlignment: "left", rowHeight: 28 };
const checks = [
  "1. 日期、网点、流向和必填字段不能为空",
  "2. 选择字段必须使用下拉选项；“全部”由系统展开，多选组合使用英文竖线|分隔",
  "3. 考核周期只能单选周或月",
  "4. 同一范围组合的生效日期重叠时阻止导入",
  "5. 跨月数据由系统拆分，Excel中无需按月重复填写",
  "6. 导入失败数据进入待处理信息池，业务修改后可继续导入",
];
checks.forEach((value, i) => {
  guide.getRange(`F${14 + i}:I${14 + i}`).merge();
  guide.getRange(`F${14 + i}`).values = [[value]];
});
guide.getRange("F14:I19").format = { fill: pale, font: { name: font, size: 10, color: text }, verticalAlignment: "center", rowHeight: 26 };
guide.getRange("A:A").format.columnWidth = 20;
guide.getRange("B:B").format.columnWidth = 12;
guide.getRange("C:C").format.columnWidth = 16;
guide.getRange("D:D").format.columnWidth = 48;
guide.getRange("E:E").format.columnWidth = 6;
guide.getRange("F:F").format.columnWidth = 20;
guide.getRange("G:G").format.columnWidth = 38;
guide.getRange("H:H").format.columnWidth = 10;
guide.getRange("I:I").format.columnWidth = 36;

enums.showGridLines = false;
enums.getRange("A1:E1").values = [["产品类型选择项", "货物类型选择项", "公斤段选择项", "考核周期（单选）", "使用说明"]];
const enumRows = Math.max(products.length, cargos.length, bands.length, 2);
enums.getRange("A2").write(Array.from({ length: enumRows }, (_, i) => [
  products[i] ?? null,
  cargos[i] ?? null,
  bands[i] ?? null,
  ["周", "月"][i] ?? null,
  i === 0 ? "“全部”表示当前列全部选项；多选组合使用英文竖线|分隔。" : null,
]));
enums.getRange("A1:E1").format = { fill: orange, font: { name: font, size: 10, bold: true, color: "#FFFFFF" }, horizontalAlignment: "center", verticalAlignment: "center", rowHeight: 28 };
enums.getRange(`A2:E${enumRows + 1}`).format = { font: { name: font, size: 10, color: text }, verticalAlignment: "center", borders: { preset: "inside", style: "thin", color: line } };
enums.getRange("A:A").format.columnWidth = 30;
enums.getRange("B:B").format.columnWidth = 24;
enums.getRange("C:C").format.columnWidth = 72;
enums.getRange("D:D").format.columnWidth = 18;
enums.getRange("E:E").format.columnWidth = 52;
enums.freezePanes.freezeRows(1);

workbook.recalculate();
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(tempPath);

const verifyWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(tempPath));
const errors = await verifyWorkbook.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!", options: { useRegex: true, maxResults: 100 }, summary: "final formula error scan" });
console.log(errors.ndjson);
for (const sheetName of ["政策导入模板", "填写说明", "枚举值"]) {
  const preview = await verifyWorkbook.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(`${previewDir}/final-${sheetName}.png`, new Uint8Array(await preview.arrayBuffer()));
}
await fs.copyFile(tempPath, outputPath);
await fs.unlink(tempPath);
