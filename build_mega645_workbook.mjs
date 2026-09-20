import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const sourcePath = "Vietlott_Mega_645_Full_Results.xlsx";
const outputDir = "outputs/vietlott-mega645";
const outputPath = `${outputDir}/Vietlott_Mega_645_Crawled_Data.xlsx`;

const sourceBlob = await FileBlob.load(sourcePath);
const sourceWorkbook = await SpreadsheetFile.importXlsx(sourceBlob);
const sourceSheet = sourceWorkbook.worksheets.getItemAt(0);
const sourceValues = sourceSheet.getUsedRange(true).values;
const records = sourceValues.slice(1).filter((row) => row[0] && row[1]);

if (records.length === 0) {
  throw new Error("The source workbook does not contain any Mega 6/45 result rows.");
}

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("Mega 6-45");
sheet.showGridLines = false;
sheet.tabColor = "#C00000";

sheet.mergeCells("A1:K1");
sheet.getRange("A1").values = [["Vietlott Mega 6/45 — Crawled Results"]];
sheet.getRange("A1:K1").format = {
  fill: "#C00000",
  font: { name: "Aptos", size: 18, bold: true, color: "#FFFFFF" },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
sheet.getRange("A1:K1").format.rowHeight = 30;

sheet.mergeCells("A2:K2");
sheet.getRange("A2").values = [["Historical draw data from the repository's Mega 6/45 crawler snapshot. Refresh by running vietlott.py, then rebuild this file."]];
sheet.getRange("A2:K2").format = {
  font: { name: "Aptos", size: 10, color: "#666666", italic: true },
  horizontalAlignment: "left",
  verticalAlignment: "center",
};

sheet.getRange("A4:F5").values = [
  ["Draws captured", null, "First draw", null, "Latest draw", null],
  [null, null, null, null, null, null],
];
sheet.getRange("H4:K5").values = [
  ["Latest draw date", null, "Latest jackpot (VND)", null],
  [null, null, null, null],
];
sheet.getRange("A4:K5").format.font = { name: "Aptos", size: 10 };
sheet.getRange("A4:K4").format = {
  fill: "#FCE4D6",
  font: { name: "Aptos", size: 10, bold: true, color: "#9C0006" },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
sheet.getRange("A5:K5").format = {
  fill: "#FFF8F3",
  font: { name: "Aptos", size: 12, bold: true, color: "#1F1F1F" },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
sheet.getRange("A4:K5").format.borders = { preset: "all", style: "thin", color: "#E6B8AF" };

const firstDataRow = 9;
const lastDataRow = firstDataRow + records.length - 1;
sheet.getRange("B5").formulas = [[`=COUNTA(A${firstDataRow}:A${lastDataRow})`]];
sheet.getRange("D5").formulas = [[`=A${firstDataRow}`]];
sheet.getRange("F5").formulas = [[`=A${lastDataRow}`]];
sheet.getRange("I5").formulas = [[`=B${lastDataRow}`]];
sheet.getRange("K5").formulas = [[`=J${lastDataRow}`]];
sheet.getRange("K5").format.numberFormat = "#,##0 \"VND\"";

const headers = [[
  "Draw ID", "Draw date", "Weekday", "Number 1", "Number 2", "Number 3",
  "Number 4", "Number 5", "Number 6", "Jackpot (VND)", "Jackpot winners",
]];
sheet.getRange(`A8:K${lastDataRow}`).values = [...headers, ...records];
const table = sheet.tables.add(`A8:K${lastDataRow}`, true, "Mega645Results");
table.style = "TableStyleMedium2";

sheet.getRange(`A8:K${lastDataRow}`).format.font = { name: "Aptos", size: 10 };
sheet.getRange(`A${firstDataRow}:C${lastDataRow}`).format.horizontalAlignment = "center";
sheet.getRange(`D${firstDataRow}:I${lastDataRow}`).format = {
  font: { name: "Aptos", size: 10, bold: true, color: "#1F4E78" },
  horizontalAlignment: "center",
  numberFormat: "00",
};
sheet.getRange(`J${firstDataRow}:J${lastDataRow}`).format.numberFormat = "#,##0 \"VND\"";
sheet.getRange(`J${firstDataRow}:K${lastDataRow}`).format.horizontalAlignment = "right";
sheet.getRange(`A8:K${lastDataRow}`).format.borders = { preset: "inside", style: "thin", color: "#E7E6E6" };

sheet.getRange("A:A").format.columnWidth = 12;
sheet.getRange("B:B").format.columnWidth = 14;
sheet.getRange("C:C").format.columnWidth = 14;
sheet.getRange("D:I").format.columnWidth = 11;
sheet.getRange("J:J").format.columnWidth = 20;
sheet.getRange("K:K").format.columnWidth = 17;
sheet.freezePanes.freezeRows(8);

workbook.recalculate();

const formulaErrors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!",
  options: { useRegex: true, maxResults: 50 },
  summary: "formula error scan",
});
if (formulaErrors.ndjson.includes("#REF!") || formulaErrors.ndjson.includes("#DIV/0!")) {
  throw new Error(`Formula check failed: ${formulaErrors.ndjson}`);
}

await fs.mkdir(outputDir, { recursive: true });
const preview = await workbook.render({ sheetName: "Mega 6-45", range: "A1:K24", scale: 1.5, format: "png" });
await fs.writeFile(`${outputDir}/preview.png`, new Uint8Array(await preview.arrayBuffer()));
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);

console.log(JSON.stringify({ outputPath, records: records.length, first: records[0], latest: records.at(-1) }));
