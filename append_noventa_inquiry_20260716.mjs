import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const projectDir = "D:\\brushcustom-email-imap";
const excelPath = "D:\\brushcustom-email-imap\\brushcustom询盘记录_updated.xlsx";

const row = [
  46219,
  "Kodchaphan Sae-chao (LookMee)",
  "kodchaphan@noventa.com",
  "ROUND BRUSH 报价请求；附件含 STEP 图纸。材料：塑料本体 PA-GF30 或 PBT-GF30 Pocan 等同材料，需耐连续蒸汽压力和最高 150°C 不变形/软化；刷丝 100% PBT，黑色，实心圆丝 0.20 mm，中等硬度，适合湿蒸汽环境；金属钉/锚固植毛，高拉拔力；内部需精密成型 bayonet twist-lock 蒸汽连接通道；可见自由刷毛长度 26 mm。报价数量：1,000 / 2,000 / 5,000 / 10,000 pcs；年需求约 40K pcs，可按此设计模穴；交货到泰国。",
  "Noventa (Thailand) Co., Ltd.",
  "https://www.noventa.com/en",
  "塑料注塑、OEM 制造、产品组装、工程开发及供应链管理",
  "201-500 人（泰国公司）",
  "待确认",
  "泰国 Prachinburi, 304 Industrial Park",
];

await fs.mkdir(path.join(projectDir, "work"), { recursive: true });
const timestamp = new Date().toISOString().replace(/[-:T]/g, "").slice(0, 14);
const backupPath = path.join(projectDir, "work", `brushcustom询盘记录_backup_before_noventa_${timestamp}.xlsx`);
await fs.copyFile(excelPath, backupPath);

const input = await FileBlob.load(excelPath);
const workbook = await SpreadsheetFile.importXlsx(input);
const sheet = workbook.worksheets.getItem("Sheet1");
const used = sheet.getUsedRange(true);
const values = used.values ?? [];
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
  sheet.getRange(`A${targetRowNumber}:J${targetRowNumber}`).format.rowHeight = 118;
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
const previewPath = path.join(projectDir, "work", "preview_after_noventa_20260716.png");
await fs.writeFile(previewPath, new Uint8Array(await preview.arrayBuffer()));

const exported = await SpreadsheetFile.exportXlsx(workbook);
let savedPath = excelPath;
let saveWarning = null;
try {
  await exported.save(excelPath);
} catch (error) {
  savedPath = path.join(projectDir, `brushcustom询盘记录_updated_pending_noventa_${timestamp}.xlsx`);
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
