import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const projectDir = "D:\\brushcustom-email-imap";
const excelPath = "D:\\brushcustom-email-imap\\brushcustom询盘记录_updated.xlsx";

const row = [
  46217,
  "Rita Finale",
  "rfinale@cisa.net",
  "按样品照片定制刷子；需确认可否按定制规格生产，重点包括刷毛长度、刷毛粗细/直径、硬度、颜色；手柄需为塑料或全包塑；刷毛必须为合成材料，不接受动物来源刷毛；客户要求提供基于 MOQ 的大致单价、MOQ、样品费、预计生产交期及到目的地运输时间。初始订单量较小，合格后预计重复订单。",
  "CISA, CEDACERIA INDUSTRIAL, S.L.",
  "https://www.cisa.net/",
  "实验室/工业筛分设备、认证筛网及粒度分析设备制造",
  "待确认",
  "待确认",
  "西班牙 Barcelona, Llissa de Vall",
];

await fs.mkdir(path.join(projectDir, "work"), { recursive: true });
const timestamp = new Date().toISOString().replace(/[-:T]/g, "").slice(0, 14);
const backupPath = path.join(projectDir, "work", `brushcustom询盘记录_backup_before_cisa_${timestamp}.xlsx`);
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
  sheet.getRange(`A${targetRowNumber}:J${targetRowNumber}`).format.rowHeight = 88;
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
const previewPath = path.join(projectDir, "work", "preview_after_cisa_20260714.png");
await fs.writeFile(previewPath, new Uint8Array(await preview.arrayBuffer()));

const exported = await SpreadsheetFile.exportXlsx(workbook);
let savedPath = excelPath;
let saveWarning = null;
try {
  await exported.save(excelPath);
} catch (error) {
  savedPath = path.join(projectDir, `brushcustom询盘记录_updated_pending_cisa_${timestamp}.xlsx`);
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
      companies: finalValues.slice(1).map((item) => item[4]),
    },
    null,
    2,
  ),
);

process.exit(0);
