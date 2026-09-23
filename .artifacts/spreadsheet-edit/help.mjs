import { Workbook } from "@oai/artifact-tool";
const workbook = Workbook.create();
console.log(workbook.help("range.dataValidation", { include: "index,examples,notes", maxChars: 5000 }).ndjson);
