import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const projectDir = "D:\\brushcustom-email-imap";
const excelPath = "D:\\brushcustom-email-imap\\brushcustom询盘记录_updated.xlsx";

const row = [
  46223,
  "Jim Farina",
  "jfarina@nichiha.com",
  "按图片定制长条刷/尼龙条刷询盘：客户询问能否制造附件图片中的刷子，并表示可继续发送规格。表单来源页面为 Industrial Nylon Rigid Rubber Strip Brush；图片显示灰/白色长条底座，密集白色尼龙刷丝，疑似生产线用长条刷/密封或清扫刷。当前缺少尺寸、材质、数量、应用场景等详细规格，需回复索取图纸/尺寸/刷丝材质与硬度/数量/使用环境。",
  "Nichiha USA, Inc.",
  "https://www.nichiha.com/",
  "建筑材料制造；纤维水泥外墙板/建筑墙板",
  "待确认",
  "待确认",
  "美国 Georgia, Mableton",
];

await fs.mkdir(path.join(projectDir, "work"), { recursive: true });
const timestamp = new Date().toISOString().replace(/[-:T]/g, "").slice(0, 14);
const backupPath = path.join(projectDir, "work", `brushcustom询盘记录_backup_before_nichiha_${timestamp}.xlsx`);
await fs.copyFile(excelPath, backupPath);

const input = await FileBlob.load(excelPath);
const workbook = await SpreadsheetFile.importXlsx(input);
const sheet = workbook.worksheets.getItem("Sheet1");
const values = sheet.getUsedRange(true).values ?? [];
const existingEmails = new Set(values.slice(1).map((item) => String(item[2] ?? "").trim().toLowerCase()).filter(Boolean));

let added = false;
let targetRowNumber = values.length + 1;

if (!existingEmails.has(String(row[2]).toLowerCase())) {
  const table = sheet.tables.items?.[0];
  if (table?.rows?.add) {
    table.rows.add(null, [row]);
  }
  sheet.getRange(`A${values.length}:J${values.length}`).copyTo(sheet.getRange(`A${targetRowNumber}:J${targetRowNumber}`), "formats");
  sheet.getRange(`A${targetRowNumber}:J${targetRowNumber}`).values = [row];
  sheet.getRange(`A${targetRowNumber}:A${targetRowNumber}`).format.numberFormat = "m/d/yy";
  sheet.getRange(`A${targetRowNumber}:J${targetRowNumber}`).format.wrapText = true;
  sheet.getRange(`A${targetRowNumber}:J${targetRowNumber}`).format.verticalAlignment = "center";
  sheet.getRange(`A${targetRowNumber}:J${targetRowNumber}`).format.rowHeight = 105;
  added = true;
}

const finalValues = sheet.getUsedRange(true).values ?? [];
const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  maxChars: 3000,
});
console.log(errors.ndjson);

const preview = await workbook.render({
  sheetName: "Sheet1",
  range: `A1:J${finalValues.length}`,
  scale: 1,
  format: "png",
});
const previewPath = path.join(projectDir, "work", "preview_after_nichiha_20260720.png");
await fs.writeFile(previewPath, new Uint8Array(await preview.arrayBuffer()));

const exported = await SpreadsheetFile.exportXlsx(workbook);
let savedPath = excelPath;
let saveWarning = null;
try {
  await exported.save(excelPath);
} catch (error) {
  savedPath = path.join(projectDir, `brushcustom询盘记录_updated_pending_nichiha_${timestamp}.xlsx`);
  await exported.save(savedPath);
  saveWarning = `Could not overwrite main workbook: ${error.code ?? error.message}`;
}

console.log(
  JSON.stringify(
    {
      added,
      finalRows: finalValues.length,
      targetRowNumber: added ? targetRowNumber : null,
      backupPath,
      excelPath,
      savedPath,
      saveWarning,
      previewPath,
      lastRow: finalValues[finalValues.length - 1],
    },
    null,
    2,
  ),
);

process.exit(0);
