import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const projectDir = "D:\\brushcustom-email-imap";
const excelPath = "D:\\brushcustom-email-imap\\brushcustom询盘记录_updated.xlsx";

const rows = [
  [
    46226,
    "Zory",
    "zmielke@americanflexible.com",
    "定制尼龙滚刷询价：用于松动并清除 foam parts 上的小矩形废料；当前工艺使用切开的发刷，客户希望改为一整支带轴滚刷。规格：OD 76.2 mm，轴径/ID 12.7 mm，总长 508 mm，刷毛覆盖长度 381 mm；底座材料塑料；希望高密度尼龙刷毛，可接受建议。表单来源页为 customized industrial cylinder nylon roller brush，附件图片显示发刷片段固定在轴上。",
    "American Flexible Products",
    "https://americanflexible.com/",
    "密封件、垫片、绝缘材料、泡棉/橡胶/塑料转换加工制造",
    "待确认",
    "待确认",
    "美国 Illinois, Northlake",
    null,
  ],
  [
    46226,
    "Henry Niggli",
    "henry.niggli@nauticatechnologies.com",
    "正式产品询盘/现有清洁机器人滚刷替代：客户正在寻找 current brush 的替代品。现有 cylinder brush 规格：0.2 PA6 bristles/filler，外径 OD 110 mm，芯管 OD 57 mm，总长 1000 mm；询问是否有类似产品。该邮件是 7/17 来访请求后的具体规格补充，需求方向为船体清洁机器人用圆柱滚刷/耐水清洁滚刷。",
    "Nautica Technologies AG",
    "https://nauticatechnologies.com/",
    "海洋机器人/船体自主清洁与检测",
    "2-10 人（公开资料）",
    "融资约 400 万美元（Seed，公开报道）",
    "新加坡（提交地点）；公司位于瑞士 Zurich / Dietlikon",
    null,
  ],
];

await fs.mkdir(path.join(projectDir, "work"), { recursive: true });
const timestamp = new Date().toISOString().replace(/[-:T]/g, "").slice(0, 14);
const backupPath = path.join(projectDir, "work", `brushcustom询盘记录_backup_before_20260723_valid_${timestamp}.xlsx`);
await fs.copyFile(excelPath, backupPath);

const input = await FileBlob.load(excelPath);
const workbook = await SpreadsheetFile.importXlsx(input);
const sheet = workbook.worksheets.getItem("Sheet1");
const values = sheet.getUsedRange(true).values ?? [];

const beforeRows = values.length;
const startRow = beforeRows + 1;
const endRow = beforeRows + rows.length;
const table = sheet.tables.items?.[0];
if (table?.rows?.add) {
  table.rows.add(null, rows);
}
sheet.getRange(`A${beforeRows}:K${beforeRows}`).copyTo(sheet.getRange(`A${startRow}:K${endRow}`), "formats");
sheet.getRange(`A${startRow}:K${endRow}`).values = rows;
sheet.getRange(`A${startRow}:A${endRow}`).format.numberFormat = "m/d/yy";
sheet.getRange(`A${startRow}:K${endRow}`).format.wrapText = true;
sheet.getRange(`A${startRow}:K${endRow}`).format.verticalAlignment = "center";
sheet.getRange(`A${startRow}:K${endRow}`).format.rowHeight = 118;

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
  range: `A1:K${finalValues.length}`,
  scale: 1,
  format: "png",
});
const previewPath = path.join(projectDir, "work", "preview_after_valid_inquiries_20260723.png");
await fs.writeFile(previewPath, new Uint8Array(await preview.arrayBuffer()));

const exported = await SpreadsheetFile.exportXlsx(workbook);
let savedPath = excelPath;
let saveWarning = null;
try {
  await exported.save(excelPath);
} catch (error) {
  savedPath = path.join(projectDir, `brushcustom询盘记录_updated_pending_20260723_valid_${timestamp}.xlsx`);
  await exported.save(savedPath);
  saveWarning = `Could not overwrite main workbook: ${error.code ?? error.message}`;
}

console.log(
  JSON.stringify(
    {
      added: rows.length,
      finalRows: finalValues.length,
      backupPath,
      excelPath,
      savedPath,
      saveWarning,
      previewPath,
      lastRows: finalValues.slice(-rows.length),
    },
    null,
    2,
  ),
);

process.exit(0);
