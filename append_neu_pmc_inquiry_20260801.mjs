import fs from "node:fs/promises";
import path from "node:path";
import { spawnSync } from "node:child_process";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const projectDir = "D:\\brushcustom-email-imap";
const inquiryStem = "brushcustom\u8be2\u76d8\u8bb0\u5f55";
const excelPath = path.join(projectDir, `${inquiryStem}_updated.xlsx`);
const timestamp = new Date().toISOString().replace(/[-:T]/g, "").slice(0, 14);
const backupPath = path.join(projectDir, "work", `${inquiryStem}_backup_before_neu_pmc_20260801_${timestamp}.xlsx`);
const previewPath = path.join(projectDir, "work", "preview_after_neu_pmc_20260801.png");
const attachmentDir = "D:\\brushcustom-email-imap\\attachments\\2026-08-01_NEU-PMC-Pte-Ltd_Cherie-Yeow";

const emailBody = `New Brushcustom RFQ

Customer Information

Name: Cherie Yeow
Company: NEU-PMC Pte Ltd
Email: cherieyeow@neucorporation.com
Phone: 94556189

Project Information

Industry / Application: Semiconductor & ESD
Brush type: panel brushes
Estimated quantity: evaluation 4 pcs, and minimum 20pcs monthly when approved

Project details:
This panel brush is to be mounted on our machines in a semiconductor factory environment.

The requirements are:
- the base needs to be PVC
- the filaments needs to be white ESD and 0.07mm diameter

I have attached our drawing.

For quotation, if you can provide me 4pcs and 20pcs will be great.

Attachments included: BRUSH ASSEMBLY.PDF.pdf

--- Submission Source Tracking ---
Date / Time: 2026-08-01 17:24:31 CST
Page URL: https://brushcustom.com/rfq/
Source Page URL: https://brushcustom.com/industries/semiconductor-esd/
Source Page Title: ESD Brushes for Semiconductor Equipment | Brushcustom
Server Referrer: https://brushcustom.com/rfq/
Client Referrer: https://brushcustom.com/industries/semiconductor-esd/
User Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/150.0.0.0 Safari/537.36
Remote IP: 122.11.243.89
Turnstile: passed
Honeypot: clean
Rate Limit: accepted
Form Elapsed Seconds: 181
Form Source: Brushcustom RFQ Form
Powered by: Brushcustom Website RFQ
Request Endpoint: https://brushcustom.com/api/rfq.php
`;

const drawingNotes = `NEU-PMC Pte Ltd - Cherie Yeow - 2026-08-01

有效询盘判断：
- 客户来自网站 RFQ 表单，Turnstile passed、Honeypot clean、Rate Limit accepted。
- 有明确应用：semiconductor factory environment / ESD。
- 有明确产品：panel brushes / BRUSH ASSEMBLY。
- 有明确材料要求：PVC base，white ESD filaments，single brush diameter 0.07 mm。
- 有数量：evaluation 4 pcs；approved 后 minimum 20 pcs monthly。
- 有 PDF 图纸附件。

图纸信息：
- Title: BRUSH ASSEMBLY
- Part No.: N037-260608-A01
- Dwg No.: N037-260608-A01
- Rev level: A0
- Unit: MM
- Qty: 01 SET
- Material: TBA
- Heat treatment: NA
- Surface treatment: NA
- Main overall size shown on drawing: 168 x 42 x 30 mm
- Inner/brush window shown: 148 x 22 mm
- Section B-B shows base/body 15 mm, visible brush height 15 mm, overall height 30 mm, bottom lip 2 mm
- Cross-section width shown: 42 mm
- Holes/features include 72x diameter 1 THRU, 6x diameter 5 THRU, 72x diameter 5 depth/press-fit brush holes, 4x R1.5, and mounting hole positions per drawing.
- Note: ALL EDGES CHAMFER (0.10 [0.004]) UNLESS OTHERWISE SPECIFIED.

General tolerances:
- Millimeters: X ±0.5; X.X ±0.1; X.XX ±0.05; X.XXX ±0.020
- Inches: .X ±.02; .XX ±.004; .XXX ±.002; .XXXX ±.0008
- Angles: X° ±1°; X.X° ±30'
- General roughness: Ra 0.8 µm

Follow-up:
- Quote both 4 pcs evaluation and 20 pcs monthly quantity.
- Confirm ESD filament standard/resistance range, PVC grade, exact press-fit hole interpretation, tolerances, and delivery address.
`;

const row = [
  46235,
  "Cherie Yeow",
  "cherieyeow@neucorporation.com",
  "Panel brush / BRUSH ASSEMBLY 报价请求；应用：半导体工厂设备/ESD 环境，安装在客户机器上。数量：评估 4 pcs，批准后每月至少 20 pcs/月。客户要求：PVC 底座；白色 ESD 刷丝，单根刷丝直径 0.07 mm。附件图纸信息：Title BRUSH ASSEMBLY；Part No./DWG No. N037-260608-A01；Rev A0；单位 MM；QTY 01 SET；Material TBA；主体尺寸约 168 x 42 x 30 mm；刷区/内框约 148 x 22 mm；Section B-B 显示底座/本体 15 mm、可见刷毛高度 15 mm、总高 30 mm、底部唇边 2 mm；横截面宽 42 mm；孔位/特征含 72x ø1 THRU、6x ø5 THRU、72x ø5 深/压装刷孔、4x R1.5，具体孔位按图纸。图纸备注：所有边倒角 0.10 [0.004]，未特别说明。一般公差：毫米 X ±0.5，X.X ±0.1，X.XX ±0.05，X.XXX ±0.020；英寸 .X ±.02，.XX ±.004，.XXX ±.002，.XXXX ±.0008；角度 X° ±1°，X.X° ±30'；粗糙度 Ra 0.8 µm。需报价 4pcs 和 20pcs，并确认 ESD 指标/PVC 等级/压装孔解释/交付地址。",
  "NEU-PMC Pte Ltd",
  "待确认",
  "半导体装配/测试设备制造与维修；半导体工厂设备/ESD 应用",
  "待确认（新加坡注册公司，2009 年成立；UEN 200907809Z）",
  "待确认",
  "新加坡；公开注册地址为 2 Yishun Industrial Street 1 #03-29, North Point Bizhub, Singapore 768159",
  "有效询盘；网站 RFQ 表单；有 PDF 图纸；需报价 4pcs 评估和 20pcs/月量产；官网/营收待确认",
  attachmentDir,
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
await fs.mkdir(attachmentDir, { recursive: true });
await fs.copyFile(excelPath, backupPath);

const workPdf = path.join(projectDir, "work", "attachments", "NEU-PMC", "03_BRUSH ASSEMBLY.PDF.pdf");
const officialPdf = path.join(attachmentDir, "BRUSH ASSEMBLY.PDF.pdf");
try {
  await fs.copyFile(workPdf, officialPdf);
} catch {
  // The PDF may already be in the official attachment folder.
}

const renderedPage = path.join(projectDir, "work", "neu_pmc_brush_assembly_page-1.png");
const officialRenderedPage = path.join(attachmentDir, "BRUSH ASSEMBLY_page-1.png");
try {
  await fs.copyFile(renderedPage, officialRenderedPage);
} catch {
  // Rendering is a helpful audit artifact, but the PDF is the source attachment.
}

await fs.writeFile(path.join(attachmentDir, "inquiry_email.txt"), emailBody, "utf8");
await fs.writeFile(path.join(attachmentDir, "notes.txt"), drawingNotes, "utf8");
await fs.writeFile(
  path.join(attachmentDir, "matched_mailbox.txt"),
  [
    "Mailbox: INBOX",
    "Fetch window: 2026-07-31 to 2026-08-03 exclusive",
    "IMAP id in fetched JSON: 17",
    "Subject: Brushcustom RFQ - NEU-PMC Pte Ltd - panel brushes",
    "Read-only policy: selected readonly and fetched with BODY.PEEK[].",
    "",
  ].join("\n"),
  "utf8",
);

const input = await FileBlob.load(excelPath);
const workbook = await SpreadsheetFile.importXlsx(input);
const sheet = workbook.worksheets.getItem("Sheet1");
const values = sheet.getUsedRange(true).values ?? [];

const existingIndex = values.findIndex((item, index) => {
  if (index === 0) return false;
  const email = String(item[2] ?? "").trim().toLowerCase();
  const product = String(item[3] ?? "");
  const company = String(item[4] ?? "");
  return (
    email === "cherieyeow@neucorporation.com" ||
    product.includes("N037-260608-A01") ||
    company.toLowerCase().includes("neu-pmc")
  );
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
sheet.getRange(`A${targetRow}:M${targetRow}`).format.rowHeight = 190;

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
  savedPath = path.join(projectDir, `${inquiryStem}_updated_pending_neu_pmc_20260801_${timestamp}.xlsx`);
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
      attachmentDir,
      fontRepair,
      row,
    },
    null,
    2,
  ),
);
