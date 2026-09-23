import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputPath = "E:/融辉物流/网点返利/99-产品原型prototype/政策管理/政策管理导入模板.xlsx";
const previewDir = "E:/融辉物流/网点返利/.artifacts/policy_import_template/previews";
const font = "Arial";
const orange = "#C2410C";
const orange2 = "#EA580C";
const cream = "#FFF7ED";
const peach = "#FFEDD5";
const line = "#E7E5E4";
const text = "#292524";
const sub = "#78716C";

const workbook = Workbook.create();
const input = workbook.worksheets.add("政策导入模板");
const guide = workbook.worksheets.add("填写说明");
const enums = workbook.worksheets.add("枚举值");

input.showGridLines = false;
guide.showGridLines = false;
enums.showGridLines = false;
input.tabColor = orange2;
guide.tabColor = "#FB923C";
enums.tabColor = "#FDBA74";

const headers = [
  "生效开始日期","生效结束日期","寄件网点","目的流向","产品类型","货物类型","公斤段","考核周期",
  "日基准量","日返利单价","工作日基准量","工作日返利单价","周日基准量","周日返利单价","节假日基准量","节假日返利单价"
];
const sample = [
  new Date(2026,8,15),new Date(2026,9,31),"杭州临平网点","杭州分拨中心",
  "融汇达|精准零担","轻抛|重货","0kg-100kg|100kg-300kg|300kg-800kg","周",
  500,0.01,650,0.012,600,0.016,700,0.022
];

input.getRange("A1:P2").values = [headers,sample];
const table = input.tables.add("A1:P2", true, "PolicyImportTable");
table.style = "TableStyleMedium2";
table.showFilterButton = true;
input.freezePanes.freezeRows(1);

input.getRange("A1:P1").format = {
  fill: orange,
  font: { name: font, size: 10, bold: true, color: "#FFFFFF" },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  wrapText: true,
  borders: { preset: "all", style: "thin", color: "#FFFFFF" },
  rowHeight: 38
};
input.getRange("A2:P2").format = {
  fill: cream,
  font: { name: font, size: 10, color: text },
  verticalAlignment: "center",
  borders: { preset: "inside", style: "thin", color: line },
  rowHeight: 28
};
input.getRange("A2:B10000").setNumberFormat("yyyy-mm-dd");
input.getRange("I2:I10000").setNumberFormat("#,##0.00");
input.getRange("K2:K10000").setNumberFormat("#,##0.00");
input.getRange("M2:M10000").setNumberFormat("#,##0.00");
input.getRange("O2:O10000").setNumberFormat("#,##0.00");
input.getRange("J2:J10000").setNumberFormat("0.0000");
input.getRange("L2:L10000").setNumberFormat("0.0000");
input.getRange("N2:N10000").setNumberFormat("0.0000");
input.getRange("P2:P10000").setNumberFormat("0.0000");
input.getRange("H2:H10000").dataValidation = { rule: { type: "list", values: ["周","月"] } };

const widths = [15,15,20,22,25,20,34,12,14,15,16,17,16,17,17,18];
for (let i=0;i<widths.length;i++) input.getRangeByIndexes(0,i,2,1).format.columnWidth = widths[i];

guide.getRange("A2:F2").merge();
guide.getRange("A2").values = [["网点返利政策导入填写说明"]];
guide.getRange("A2:F2").format = { font:{name:font,size:15,bold:true,color:text}, verticalAlignment:"center", rowHeight:28 };
guide.getRange("A3:F3").merge();
guide.getRange("A3").values = [["多选字段使用英文竖线“|”分隔；系统按自然月和多选组合自动拆分。"]];
guide.getRange("A3:F3").format = { font:{name:font,size:10,italic:true,color:sub}, rowHeight:24 };

guide.getRange("A5:D5").values = [["字段名称","是否必填","填写格式","填写规则"]];
const fieldRows = [
  ["生效开始日期","是","yyyy-mm-dd","政策生效起始日期，不得晚于结束日期"],
  ["生效结束日期","是","yyyy-mm-dd","允许跨月；系统按自然月拆分"],
  ["寄件网点","是","文本","填写系统中的标准网点名称"],
  ["目的流向","是","文本","填写分拨或营运区名称；营运区覆盖其下所有分拨"],
  ["产品类型","是","多选文本","多个值用|分隔；可选融汇达、精准零担、融安达、融速达"],
  ["货物类型","是","多选文本","多个值用|分隔；当前示例为轻抛、重货"],
  ["公斤段","是","多选文本","多个值用|分隔；必须使用枚举值中的完整名称"],
  ["考核周期","是","枚举","只能填写周或月"],
  ["日基准量","是","数值（kg）","大于或等于0"],
  ["日返利单价","是","数值（元/kg）","大于或等于0，建议保留4位小数"],
  ["工作日基准量","否","数值（kg）","为空时继承日基准量"],
  ["工作日返利单价","否","数值（元/kg）","为空时继承日返利单价"],
  ["周日基准量","否","数值（kg）","为空时继承日基准量"],
  ["周日返利单价","否","数值（元/kg）","为空时继承日返利单价"],
  ["节假日基准量","否","数值（kg）","为空时按当天星期继承工作日或周日基准量"],
  ["节假日返利单价","否","数值（元/kg）","为空时按当天星期继承工作日或周日返利单价"]
];
guide.getRange("A6:D21").values = fieldRows;
guide.getRange("A5:D5").format = { fill:orange,font:{name:font,size:10,bold:true,color:"#FFFFFF"},horizontalAlignment:"center",verticalAlignment:"center",rowHeight:28,borders:{preset:"all",style:"thin",color:"#FFFFFF"} };
guide.getRange("A6:D21").format = { font:{name:font,size:10,color:text},verticalAlignment:"center",wrapText:true,borders:{preset:"inside",style:"thin",color:line} };
guide.getRange("A6:D21").format.rowHeight = 29;
guide.getRange("A6:D21").conditionalFormats.add("Custom", { formula:'=$B6="是"', format:{ fill:"#FFF7ED" } });
guide.getRange("A:D").format.font = { name:font,size:10,color:text };
guide.getRange("A:A").format.columnWidth = 20;
guide.getRange("B:B").format.columnWidth = 12;
guide.getRange("C:C").format.columnWidth = 20;
guide.getRange("D:D").format.columnWidth = 58;
guide.freezePanes.freezeRows(5);

guide.getRange("F5:J5").merge();
guide.getRange("F5").values = [["系统拆分规则示例"]];
guide.getRange("F5:J5").format = { fill:orange2,font:{name:font,size:10,bold:true,color:"#FFFFFF"},horizontalAlignment:"left",verticalAlignment:"center",rowHeight:28 };
const splitRows = [
  ["输入项目","示例","数量","说明",""],
  ["日期范围","2026-09-15 至 2026-10-31",2,"跨2个自然月",""],
  ["产品类型","融汇达|精准零担",2,"选择2个产品类型",""],
  ["货物类型","轻抛|重货",2,"选择2个货物类型",""],
  ["公斤段","0kg-100kg|100kg-300kg|300kg-800kg",3,"选择3个公斤段",""],
  ["拆分结果","2 × 2 × 2 × 3",24,"导入系统后形成24条月度配置",""],
];
guide.getRange("F6:J11").values = splitRows;
guide.getRange("F6:J6").format = { fill:peach,font:{name:font,size:10,bold:true,color:orange},horizontalAlignment:"center",verticalAlignment:"center",rowHeight:26 };
guide.getRange("F7:J11").format = { font:{name:font,size:10,color:text},verticalAlignment:"center",wrapText:true,borders:{preset:"inside",style:"thin",color:line},rowHeight:31 };
guide.getRange("F:F").format.columnWidth = 18;
guide.getRange("G:G").format.columnWidth = 36;
guide.getRange("H:H").format.columnWidth = 10;
guide.getRange("I:I").format.columnWidth = 35;
guide.getRange("J:J").format.columnWidth = 4;

guide.getRange("F13:J13").merge();
guide.getRange("F13").values = [["导入校验"]];
guide.getRange("F13:J13").format = { fill:orange2,font:{name:font,size:10,bold:true,color:"#FFFFFF"},rowHeight:27 };
guide.getRange("F14:J18").merge(true);
guide.getRange("F14:F18").values = [["1. 日期、网点、流向和必填字段不能为空"],["2. 多选值必须来自枚举值，使用英文竖线|分隔"],["3. 同一组合的生效日期不得重叠"],["4. 跨月数据由系统拆分，Excel中无需按月重复填写"],["5. 导入几十万条数据时由后台异步处理并生成任务记录"]];
guide.getRange("F14:J18").format = { fill:cream,font:{name:font,size:10,color:text},verticalAlignment:"center",wrapText:true,rowHeight:29 };

const enumMatrix = [
  ["产品类型","货物类型","公斤段","考核周期"],
  ["融汇达","轻抛","0kg-100kg","周"],
  ["精准零担","重货","100kg-300kg","月"],
  ["融安达","","300kg-800kg",""],
  ["融速达","","800kg-1500kg",""],
  ["","","1500kg-3000kg",""],
  ["","","3T-10T",""],
  ["","","10T以上",""]
];
enums.getRange("A1:D8").values = enumMatrix;
enums.getRange("A1:D1").format = { fill:orange,font:{name:font,size:10,bold:true,color:"#FFFFFF"},horizontalAlignment:"center",verticalAlignment:"center",rowHeight:28,borders:{preset:"all",style:"thin",color:"#FFFFFF"} };
enums.getRange("A2:D8").format = { font:{name:font,size:10,color:text},verticalAlignment:"center",borders:{preset:"inside",style:"thin",color:line},rowHeight:26 };
enums.getRange("A:D").format.columnWidth = 22;
enums.getRange("F1:J1").merge();
enums.getRange("F1").values = [["公斤段边界规则"]];
enums.getRange("F1:J1").format = { fill:orange2,font:{name:font,size:10,bold:true,color:"#FFFFFF"},rowHeight:28 };
enums.getRange("F2:J4").merge(true);
enums.getRange("F2:F4").values = [["100kg归入0kg-100kg"],["300kg、800kg、1500kg、3000kg等边界值均归入前一档"],["多选示例：0kg-100kg|100kg-300kg|300kg-800kg"]];
enums.getRange("F2:J4").format = { fill:cream,font:{name:font,size:10,color:text},verticalAlignment:"center",wrapText:true,rowHeight:30 };
enums.getRange("F:J").format.columnWidth = 18;
enums.freezePanes.freezeRows(1);

workbook.recalculate();
await fs.mkdir(previewDir,{recursive:true});
for (const [sheetName,range] of [["政策导入模板","A1:P5"],["填写说明","A1:J21"],["枚举值","A1:J8"]]) {
  const image = await workbook.render({sheetName,range,scale:1.4,format:"png"});
  const bytes = new Uint8Array(await image.arrayBuffer());
  await fs.writeFile(`${previewDir}/${sheetName}.png`,bytes);
}

const inspect = await workbook.inspect({kind:"table",range:"政策导入模板!A1:P2",include:"values,formulas",tableMaxRows:5,tableMaxCols:20,maxChars:8000});
console.log(inspect.ndjson);
const errors = await workbook.inspect({kind:"match",searchTerm:"#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!",options:{useRegex:true,maxResults:100},summary:"final formula error scan"});
console.log(errors.ndjson);

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(JSON.stringify({outputPath,previewDir}));
