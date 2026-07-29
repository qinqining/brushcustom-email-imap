import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const projectDir = "D:\\brushcustom-email-imap";
const excelPath = "D:\\brushcustom-email-imap\\brushcustom询盘记录_updated.xlsx";

const rows = [
  [
    46220,
    "Henry Niggli",
    "henry.niggli@nauticatechnologies.com",
    "访问请求/潜在线索：Nautica Technologies 是做船体自主清洁机器人的瑞士海洋机器人创业公司，目前在深圳寻找零部件并拜访供应商；希望当天下午到办公室了解 BrushCustom 产品。邮件未给具体刷子规格、数量或图纸，可能需求方向为船体清洁机器人用刷子、耐海水/耐磨清洁刷、刷盘或定制刷组件。建议先确认应用场景、目标刷型、尺寸/材质、来访人数和具体时间。",
    "Nautica Technologies AG",
    "https://nauticatechnologies.com/",
    "海洋机器人/船体自主清洁与检测",
    "2-10 人（公开资料）",
    "融资约 400 万美元（Seed，公开报道）",
    "瑞士 Zurich / Dietlikon；当前在深圳寻源",
  ],
  [
    46220,
    "Jasper Poppele",
    "jasper.poppele@assetcool.com",
    "诈骗事件后合作恢复/采购跟进：AssetCool 之前与 VC Zhang 沟通并疑似遭遇中间人邮件诈骗，欺诈者使用假冒 VC 邮箱（@asia.com）和假冒 AssetCool 员工邮箱（@dr.com）夹在双方中间，导致其误向骗子付款。7/17 最新回复称仍想购买此前讨论的 brush 用于生产，希望了解善意折扣比例，并向内部财务说明原 PI 银行信息与 2024 年付款记录匹配，以争取继续合作/重新付款。当前重点不是新规格 RFQ，而是恢复信任、确认安全付款渠道并推进原刷子采购。",
    "AssetCool / Cable Coatings Limited",
    "https://www.assetcool.com/",
    "电网机器人、功能涂层、架空电力线增容与维护",
    "待确认",
    "2025 年完成约 1000 万英镑融资（公开报道）",
    "英国 Batley, West Yorkshire",
  ],
];

await fs.mkdir(path.join(projectDir, "work"), { recursive: true });
const timestamp = new Date().toISOString().replace(/[-:T]/g, "").slice(0, 14);
const backupPath = path.join(projectDir, "work", `brushcustom询盘记录_backup_before_nautica_assetcool_${timestamp}.xlsx`);
await fs.copyFile(excelPath, backupPath);

const input = await FileBlob.load(excelPath);
const workbook = await SpreadsheetFile.importXlsx(input);
const sheet = workbook.worksheets.getItem("Sheet1");
const used = sheet.getUsedRange(true);
const values = used.values ?? [];
const existingEmails = new Set(values.slice(1).map((item) => String(item[2] ?? "").trim().toLowerCase()).filter(Boolean));

const rowsToAdd = rows.filter((row) => !existingEmails.has(String(row[2]).toLowerCase()));
let added = 0;

if (rowsToAdd.length > 0) {
  const beforeRows = values.length;
  const startRow = beforeRows + 1;
  const endRow = beforeRows + rowsToAdd.length;
  const table = sheet.tables.items?.[0];
  if (table?.rows?.add) {
    table.rows.add(null, rowsToAdd);
  }
  sheet.getRange(`A${beforeRows}:J${beforeRows}`).copyTo(sheet.getRange(`A${startRow}:J${endRow}`), "formats");
  sheet.getRange(`A${startRow}:J${endRow}`).values = rowsToAdd;
  sheet.getRange(`A${startRow}:A${endRow}`).format.numberFormat = "m/d/yy";
  sheet.getRange(`A${startRow}:J${endRow}`).format.wrapText = true;
  sheet.getRange(`A${startRow}:J${endRow}`).format.verticalAlignment = "center";
  sheet.getRange(`A${startRow}:J${endRow}`).format.rowHeight = 125;
  added = rowsToAdd.length;
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
const previewPath = path.join(projectDir, "work", "preview_after_nautica_assetcool_20260717.png");
await fs.writeFile(previewPath, new Uint8Array(await preview.arrayBuffer()));

const exported = await SpreadsheetFile.exportXlsx(workbook);
let savedPath = excelPath;
let saveWarning = null;
try {
  await exported.save(excelPath);
} catch (error) {
  savedPath = path.join(projectDir, `brushcustom询盘记录_updated_pending_nautica_assetcool_${timestamp}.xlsx`);
  await exported.save(savedPath);
  saveWarning = `Could not overwrite main workbook: ${error.code ?? error.message}`;
}

console.log(
  JSON.stringify(
    {
      added,
      skipped: rows.length - added,
      finalRows: finalValues.length,
      backupPath,
      excelPath,
      savedPath,
      saveWarning,
      previewPath,
      lastRows: finalValues.slice(-rowsToAdd.length || undefined),
    },
    null,
    2,
  ),
);

process.exit(0);
