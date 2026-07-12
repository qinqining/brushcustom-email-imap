import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const projectDir = "D:\\brushcustom-email-imap";
const outputDir = "C:\\Users\\HP\\Documents\\Codex\\2026-07-12\\http-127-0-0-1-4321\\outputs";

function parseEnv(text) {
  const env = {};
  for (const rawLine of text.split(/\r?\n/)) {
    const line = rawLine.trim();
    if (!line || line.startsWith("#") || !line.includes("=")) continue;
    const idx = line.indexOf("=");
    env[line.slice(0, idx).trim()] = line.slice(idx + 1).trim();
  }
  return env;
}

const env = parseEnv(await fs.readFile(path.join(projectDir, ".env"), "utf8"));
const excelPath = env.EXCEL_PATH;
const input = await FileBlob.load(excelPath);
const workbook = await SpreadsheetFile.importXlsx(input);
const sheet = workbook.worksheets.getItem("Sheet1");

sheet.getRange("A1:J9").format.wrapText = true;
sheet.getRange("A1:J9").format.verticalAlignment = "center";
sheet.getRange("A2:A9").format.numberFormat = "m/d/yy";

const widths = [12, 17, 28, 66, 30, 28, 36, 20, 20, 22];
for (let col = 0; col < widths.length; col += 1) {
  sheet.getRangeByIndexes(0, col, 9, 1).format.columnWidth = widths[col];
}

const rowHeights = {
  1: 35,
  5: 96,
  6: 92,
  7: 58,
  8: 84,
  9: 60,
};
for (const [row, height] of Object.entries(rowHeights)) {
  sheet.getRange(`A${row}:J${row}`).format.rowHeight = height;
}

const preview = await workbook.render({
  sheetName: "Sheet1",
  range: "A1:J9",
  scale: 1,
  format: "png",
});
await fs.writeFile(
  path.join(projectDir, "work", "final_workbook_preview.png"),
  new Uint8Array(await preview.arrayBuffer()),
);

const exported = await SpreadsheetFile.exportXlsx(workbook);
await exported.save(excelPath);
await fs.mkdir(outputDir, { recursive: true });
await exported.save(path.join(outputDir, "brushcustom询盘记录_updated.xlsx"));

console.log(JSON.stringify({ excelPath, rows: 9 }, null, 2));
