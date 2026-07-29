import fs from "node:fs/promises";
import path from "node:path";
import { spawnSync } from "node:child_process";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const projectDir = "D:\\brushcustom-email-imap";
const inquiryStem = "brushcustom\u8be2\u76d8\u8bb0\u5f55";
const excelPath = path.join(projectDir, `${inquiryStem}_updated.xlsx`);
const timestamp = new Date().toISOString().replace(/[-:T]/g, "").slice(0, 14);
const backupPath = path.join(projectDir, "work", `${inquiryStem}_backup_before_attachment_columns_${timestamp}.xlsx`);

const attachmentMap = {
  "oyerli@emremakinasan.com.tr": {
    path: "D:\\brushcustom-email-imap\\attachments\\2026-06-16_Emre-Makina_Mehmet-Okan-Yerli",
    url: "",
  },
  "glam@tnt-group.com": {
    path: "D:\\brushcustom-email-imap\\attachments\\2026-06-18_TNT-Group_Gigi-Lam",
    url: "",
  },
  "admin02@lienminhine.vn": {
    path: "D:\\brushcustom-email-imap\\attachments\\2026-06-26_Lien-Minh_Trang",
    url: "",
  },
  "general@advanxtech.com.sg": {
    path: "D:\\brushcustom-email-imap\\attachments\\2026-07-09_ADVANX_Rena-Goh",
    url: "",
  },
  "mcipriano@acncutting.com": {
    path: "D:\\brushcustom-email-imap\\attachments\\2026-07-10_ACN-Cutting_Maria-Clara-Cabrita",
    url: "",
  },
  "rfinale@cisa.net": {
    path: "D:\\brushcustom-email-imap\\attachments\\2026-07-14_CISA_Rita-Finale",
    url: "",
  },
  "kodchaphan@noventa.com": {
    path: "D:\\brushcustom-email-imap\\attachments\\2026-07-16_Noventa_Kodchaphan",
    url: "",
  },
  "jasper.poppele@assetcool.com": {
    path: "D:\\brushcustom-email-imap\\attachments\\2026-07-17_AssetCool_Jasper-Poppele",
    url: "",
  },
  "jfarina@nichiha.com": {
    path: "D:\\brushcustom-email-imap\\attachments\\2026-07-20_Nichiha_Jim-Farina",
    url: "https://brushcustom.com/wp-content/uploads/elementor/forms/6a5e2d7a5e430.jpg",
  },
  "zmielke@americanflexible.com": {
    path: "D:\\brushcustom-email-imap\\attachments\\2026-07-23_American-Flexible_Zory",
    url: "https://brushcustom.com/wp-content/uploads/elementor/forms/6a6215f58e30f.jpg",
  },
};

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

await fs.mkdir(path.join(projectDir, "work"), { recursive: true });
await fs.copyFile(excelPath, backupPath);

const input = await FileBlob.load(excelPath);
const workbook = await SpreadsheetFile.importXlsx(input);
const sheet = workbook.worksheets.getItem("Sheet1");
const values = sheet.getUsedRange(true).values ?? [];
const rowCount = values.length;

sheet.getRange("J1:J1").copyTo(sheet.getRange("K1:M1"), "formats");
sheet.getRange("K1:M1").values = [["\u72b6\u6001/\u5907\u6ce8", "\u9644\u4ef6/\u8d44\u6599\u8def\u5f84", "\u539f\u59cb\u56fe\u7247URL"]];
sheet.getRange("K1:M1").format.wrapText = true;

if (rowCount > 1) {
  sheet.getRange(`J2:J${rowCount}`).copyTo(sheet.getRange(`K2:M${rowCount}`), "formats");
  const data = values.slice(1).map((row) => {
    const currentStatus = row[10] ?? null;
    const email = String(row[2] ?? "").trim().toLowerCase();
    const attachment = attachmentMap[email] ?? { path: "", url: "" };
    return [currentStatus, attachment.path, attachment.url];
  });
  sheet.getRange(`K2:M${rowCount}`).values = data;
  sheet.getRange(`K2:M${rowCount}`).format.wrapText = true;
  sheet.getRange(`K2:M${rowCount}`).format.verticalAlignment = "center";
}

sheet.getRange(`K1:K${rowCount}`).format.columnWidth = 16;
sheet.getRange(`L1:L${rowCount}`).format.columnWidth = 52;
sheet.getRange(`M1:M${rowCount}`).format.columnWidth = 58;

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
  range: `A1:M${finalValues.length}`,
  scale: 1,
  format: "png",
});
const previewPath = path.join(projectDir, "work", "preview_after_attachment_columns_20260724.png");
await fs.writeFile(previewPath, new Uint8Array(await preview.arrayBuffer()));

const exported = await SpreadsheetFile.exportXlsx(workbook);
let savedPath = excelPath;
let saveWarning = null;
try {
  await exported.save(excelPath);
} catch (error) {
  savedPath = path.join(projectDir, `${inquiryStem}_updated_pending_attachment_columns_${timestamp}.xlsx`);
  await exported.save(savedPath);
  saveWarning = `Could not overwrite main workbook: ${error.code ?? error.message}`;
}
const fontRepair = repairInvalidSharedStringFontSize(savedPath);

console.log(
  JSON.stringify(
    {
      backupPath,
      excelPath,
      savedPath,
      saveWarning,
      fontRepair,
      previewPath,
      rows: finalValues.length,
      cols: finalValues[0].length,
      header: finalValues[0],
      attachmentRows: finalValues
        .slice(1)
        .filter((row) => row[11] || row[12])
        .map((row) => ({ email: row[2], company: row[4], folder: row[11], url: row[12] })),
    },
    null,
    2,
  ),
);

process.exit(0);
