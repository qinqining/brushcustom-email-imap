import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

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

const projectDir = "D:\\brushcustom-email-imap";
const env = parseEnv(await fs.readFile(path.join(projectDir, ".env"), "utf8"));
const excelPath = env.EXCEL_PATH;
if (!excelPath) throw new Error("EXCEL_PATH is empty in .env");

const input = await FileBlob.load(excelPath);
const workbook = await SpreadsheetFile.importXlsx(input);

const inspect = await workbook.inspect({
  kind: "workbook,sheet,table,region",
  maxChars: 12000,
  tableMaxRows: 12,
  tableMaxCols: 20,
  tableMaxCellChars: 160,
});
console.log(inspect.ndjson);

const sheetInfo = await workbook.inspect({
  kind: "sheet",
  include: "id,name",
  maxChars: 4000,
});
console.error("SHEETS");
console.error(sheetInfo.ndjson);

await fs.mkdir(path.join(projectDir, "work"), { recursive: true });
try {
  const preview = await workbook.render({
    sheetName: workbook.worksheets.getItemAt(0).name,
    autoCrop: "all",
    scale: 1,
    format: "png",
  });
  await fs.writeFile(
    path.join(projectDir, "work", "workbook_preview.png"),
    new Uint8Array(await preview.arrayBuffer()),
  );
  console.error("PREVIEW D:\\brushcustom-email-imap\\work\\workbook_preview.png");
} catch (error) {
  console.error(`PREVIEW_ERROR ${error.message}`);
}
