import fs from "node:fs/promises";
import path from "node:path";
import { spawnSync } from "node:child_process";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const projectDir = "D:\\brushcustom-email-imap";
const inquiryStem = "brushcustom\u8be2\u76d8\u8bb0\u5f55";
const excelPath = path.join(projectDir, `${inquiryStem}_updated.xlsx`);
const timestamp = new Date().toISOString().replace(/[-:T]/g, "").slice(0, 14);
const backupPath = path.join(projectDir, "work", `${inquiryStem}_backup_before_rhino_20260727_${timestamp}.xlsx`);
const previewPath = path.join(projectDir, "work", "preview_after_rhino_20260727.png");
const attachmentPath = "D:\\brushcustom-email-imap\\attachments\\2026-07-27_Rhino-Roofing-Products_Wasantha";

const row = [
  46230,
  "Wasantha",
  "imports@rhino.lk",
  "Upper Cleaning Brush Fibre 报价请求；数量 02 Nos；规格：fitted ø 350 / 254.75 x 1200 mm；要求 C&F Colombo 报价，并提供品牌/原产国、技术数据表、图纸、交期、质保期和付款条款。附件 image002.jpg 为邮件签名/logo，非产品图纸。",
  "Rhino Roofing Products Ltd.",
  "https://www.rhino.lk/",
  "建筑材料/屋面和天花板材料制造（roofing & ceiling sheets、金属屋面及配套产品）",
  "待确认",
  "待确认",
  "斯里兰卡 Colombo 09",
  "有效询盘；需 C&F Colombo 报价",
  attachmentPath,
  "",
];

function repairInvalidSharedStringFontSize(filePath) {
  const script = `
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$file = @'
${filePath}
'@
$zip = [System.IO.Compression.ZipFile]::Open($file, [System.IO.Compression.ZipArchiveMode]::Update)
try {
  $entry = $zip.GetEntry('xl/sharedStrings.xml')
  if ($null -eq $entry) {
    [pscustomobject]@{Repaired=$false; Reason='sharedStrings.xml not found'} | ConvertTo-Json -Compress
    return
  }
  $reader = New-Object System.IO.StreamReader($entry.Open())
  $xml = $reader.ReadToEnd()
  $reader.Close()
  $fixed = $xml.Replace('sz val="1100"', 'sz val="11"')
  if ($fixed -eq $xml) {
    [pscustomobject]@{Repaired=$false; Reason='no invalid font size found'} | ConvertTo-Json -Compress
    return
  }
  $entry.Delete()
  $newEntry = $zip.CreateEntry('xl/sharedStrings.xml')
  $stream = $newEntry.Open()
  $utf8 = New-Object System.Text.UTF8Encoding($false)
  $writer = New-Object System.IO.StreamWriter($stream, $utf8)
  $writer.Write($fixed)
  $writer.Close()
  [pscustomobject]@{Repaired=$true; Reason='replaced sz val 1100 with 11'} | ConvertTo-Json -Compress
} finally {
  if ($null -ne $zip) { $zip.Dispose() }
}
`;
  const result = spawnSync("powershell.exe", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script], {
    encoding: "utf8",
  });
  if (result.status !== 0) {
    throw new Error(`Font repair failed: ${result.stderr || result.stdout}`);
  }
  return result.stdout.trim();
}

function hasAnyValue(rowValues) {
  return (rowValues ?? []).some((value) => value !== null && value !== undefined && String(value).trim() !== "");
}

await fs.mkdir(path.join(projectDir, "work"), { recursive: true });
await fs.copyFile(excelPath, backupPath);

const input = await FileBlob.load(backupPath);
const workbook = await SpreadsheetFile.importXlsx(input);
const sheet = workbook.worksheets.getItem("Sheet1");
const values = sheet.getUsedRange(true).values ?? [];

const existingIndex = values.findIndex((item, index) => {
  if (index === 0) return false;
  const email = String(item[2] ?? "").trim().toLowerCase();
  const product = String(item[3] ?? "");
  const company = String(item[4] ?? "");
  return email === "imports@rhino.lk" || (company.includes("Rhino") && product.includes("Upper Cleaning Brush"));
});

const lastDataIndex = values.reduce((last, item, index) => (hasAnyValue(item) ? index : last), 0);
const targetRow = existingIndex >= 0 ? existingIndex + 1 : lastDataIndex + 2;
const action = existingIndex >= 0 ? "updated" : "added";

if (action === "added") {
  const sourceRow = Math.max(2, lastDataIndex + 1);
  sheet.getRange(`A${sourceRow}:M${sourceRow}`).copyTo(sheet.getRange(`A${targetRow}:M${targetRow}`), "formats");
}

sheet.getRange(`A${targetRow}:M${targetRow}`).values = [row];
sheet.getRange(`A${targetRow}:A${targetRow}`).format.numberFormat = "m/d/yy";
sheet.getRange(`A${targetRow}:M${targetRow}`).format.wrapText = true;
sheet.getRange(`A${targetRow}:M${targetRow}`).format.verticalAlignment = "center";
sheet.getRange(`A${targetRow}:M${targetRow}`).format.rowHeight = 118;

const finalValues = sheet.getUsedRange(true).values ?? [];
const preview = await workbook.render({
  sheetName: "Sheet1",
  range: `A1:M${Math.max(targetRow, finalValues.length)}`,
  scale: 1,
  format: "png",
});
await fs.writeFile(previewPath, new Uint8Array(await preview.arrayBuffer()));

const exported = await SpreadsheetFile.exportXlsx(workbook);
let savedPath = excelPath;
let saveWarning = null;
try {
  await exported.save(excelPath);
} catch (error) {
  savedPath = path.join(projectDir, `${inquiryStem}_updated_pending_rhino_20260727_${timestamp}.xlsx`);
  await exported.save(savedPath);
  saveWarning = `Could not overwrite main workbook: ${error.code ?? error.message}`;
}

const fontRepair = repairInvalidSharedStringFontSize(savedPath);

console.log(
  JSON.stringify(
    {
      action,
      targetRow,
      backupPath,
      excelPath,
      savedPath,
      saveWarning,
      previewPath,
      fontRepair,
      row,
    },
    null,
    2,
  ),
);
