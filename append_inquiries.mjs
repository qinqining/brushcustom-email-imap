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
if (!excelPath) throw new Error("EXCEL_PATH is empty in .env");

const rows = [
  [
    46189,
    "Mehmet Okan Yerli",
    "oyerli@emremakinasan.com.tr",
    "Backenburste / scraper brush complete assembly，按图 D904918 b 制造；117 pcs；尺寸 205 x 100 x 11 mm；含 S235JRG2 钢板、C100S 硬化钢板、4 层织物橡胶涂层/圆钢丝刷条、AIA/St-H 4 x 8 ISO 15977 抽芯铆钉及粘接要求；需单价/总价、交期、图纸确认和付款/运输条款。",
    "Emre Makina Sanayi",
    "https://www.emremakinasan.com.tr/",
    "机械加工/工业设备及零部件制造或采购",
    "待确认",
    "待确认",
    "土耳其",
  ],
  [
    46191,
    "Gigi Lam",
    "glam@tnt-group.com",
    "Christmas 2027 GWP 项目：木质或醋酸纤维 body brush 与 hand/nail brush；尺寸分别约 17 x 7 cm、8 x 5 cm；方案含天然木纹+米白/白色刷毛或仿珍珠母贝；需 20k/23k/25k 报价、排期和样品，计划 2026 年 11 月底深圳出货。",
    "TNT Group",
    "https://tnt-group.com/",
    "美妆/奢侈品促销赠品、包装及产品开发",
    "中型企业（待确认）",
    "待确认",
    "中国香港（项目联系人）",
  ],
  [
    46191,
    "Paul Tonkinson",
    "Paul.Tonkinson@smcwolverhampton.co.uk",
    "304 Stainless Steel Pipe Filter Brush / Pipe Cleaning Brush，询价 10 pcs；来源产品页为 304 不锈钢管道过滤刷/管道清洁刷。",
    "Specialist Metallic Coatings",
    "https://www.smcwolverhampton.co.uk/",
    "金属表面处理/专业涂层",
    "小型企业（待确认）",
    "待确认",
    "英国 Wolverhampton",
  ],
  [
    46198,
    "Matthew Zhou",
    "extern.matthew.zhou@vw.com",
    "定制 cylinder helical brushes 2 支：长度 1300 mm，芯轴直径 50 mm，刷毛长度 50 mm，螺旋 pitch 260 mm，4 coils；一支右旋、一支左旋；用途为 conveyor skid cleaning；需产品目录、刷丝直径/材质建议、预计交期和 CAD 格式要求。",
    "Volkswagen (VW)",
    "https://www.vw.com/",
    "汽车制造",
    "大型跨国企业",
    "千亿美元级（集团规模）",
    "美国 Georgia, Atlanta",
  ],
  [
    46199,
    "Trang",
    "admin02@lienminhine.vn",
    "Wheel brush：D150 mm，孔径 25.4 mm，厚度 50 mm，flange 92 mm，0.3 mm brass coated wire；客户询问是否有合适型号并要求报价。",
    "Lien Minh Industrial Equipment Joint Stock Company",
    "http://www.lienminhine.vn/",
    "切削工具及工业刷经销/贸易",
    "小型企业 (SME)",
    "待确认",
    "越南 Hanoi",
  ],
];

await fs.mkdir(path.join(projectDir, "work"), { recursive: true });
await fs.mkdir(outputDir, { recursive: true });

const timestamp = new Date().toISOString().replace(/[-:T]/g, "").slice(0, 14);
const backupPath = path.join(projectDir, "work", `brushcustom询盘记录_backup_${timestamp}.xlsx`);
await fs.copyFile(excelPath, backupPath);

const input = await FileBlob.load(excelPath);
const workbook = await SpreadsheetFile.importXlsx(input);
const sheet = workbook.worksheets.getItem("Sheet1");
const used = sheet.getUsedRange(true);
const existingValues = used.values ?? [];
const existingEmails = new Set(
  existingValues
    .slice(1)
    .map((row) => String(row[2] ?? "").trim().toLowerCase())
    .filter(Boolean),
);

const rowsToAdd = rows.filter((row) => !existingEmails.has(String(row[2]).toLowerCase()));
const beforeRows = existingValues.length;
let added = 0;

if (rowsToAdd.length > 0) {
  const startRow = beforeRows + 1;
  const endRow = beforeRows + rowsToAdd.length;
  const table = sheet.tables.items?.[0];
  if (table?.rows?.add) {
    table.rows.add(null, rowsToAdd);
  }
  added = rowsToAdd.length;
  sheet.getRange(`A${beforeRows}:J${beforeRows}`).copyTo(sheet.getRange(`A${startRow}:J${endRow}`), "formats");
  sheet.getRange(`A${startRow}:J${endRow}`).values = rowsToAdd;
  sheet.getRange(`A${startRow}:A${endRow}`).format.numberFormat = "yyyy-mm-dd";
  sheet.getRange(`A${startRow}:J${endRow}`).format.wrapText = true;
  sheet.getRange(`A${startRow}:J${endRow}`).format.autofitRows();
}

const finalUsed = sheet.getUsedRange(true);
const finalValues = finalUsed.values ?? [];
const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  maxChars: 3000,
});
console.log(errors.ndjson);
console.log(
  JSON.stringify(
    {
      previewRows: finalValues.map((row) => row.slice(0, 5)),
    },
    null,
    2,
  ),
);

try {
  const preview = await workbook.render({
    sheetName: "Sheet1",
    range: `A1:J${Math.max(finalValues.length, 1)}`,
    scale: 1,
    format: "png",
  });
  await fs.writeFile(
    path.join(projectDir, "work", "updated_workbook_preview.png"),
    new Uint8Array(await preview.arrayBuffer()),
  );
} catch (error) {
  console.error(`PREVIEW_ERROR ${error.message}`);
}

const exported = await SpreadsheetFile.exportXlsx(workbook);
await exported.save(excelPath);
const outputPath = path.join(outputDir, "brushcustom询盘记录_updated.xlsx");
await exported.save(outputPath);

console.log(
  JSON.stringify(
    {
      added,
      skipped: rows.length - added,
      finalRows: finalValues.length,
      backupPath,
      excelPath,
      outputPath,
    },
    null,
    2,
  ),
);

process.exit(0);
