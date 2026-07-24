import fs from "fs";
import path from "path";
import { SpreadsheetFile } from "@oai/artifact-tool";

const outputDir = path.resolve("../outputs/global_2026_07_24");
const workbookName = fs.readdirSync(outputDir).find((name) => name.endsWith(".xlsx"));
if (!workbookName) throw new Error("Workbook not found");

const raw = fs.readFileSync(path.join(outputDir, workbookName));
const wb = await SpreadsheetFile.importXlsx(new Uint8Array(raw));
wb.recalculate();

const errorScan = wb.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "final formula error scan",
});

const ranking = wb.worksheets.getItem("Country Ranking");
const portfolio = wb.worksheets.getItem("Portfolio");
const inputs = wb.worksheets.getItem("Model Inputs");
const scenarios = wb.worksheets.getItem("Scenarios");
const sourceSheet = wb.worksheets.getItem("Sources & Checks");

const scores = ranking.getRange("P8:P27").values.flat().map(Number);
const targetWeights = portfolio.getRange("G12:G31").values.flat().map(Number);
const assetWeights = inputs.getRange("B5:B8").values.flat().map(Number);
const scenarioProbabilities = scenarios.getRange("B5:B7").values.flat().map(Number);
const sourceChecks = sourceSheet.getRange("D5:D10").values.flat();

const result = {
  workbookName,
  sheets: Array.from(wb.worksheets).map((sheet) => sheet.name),
  scoreMin: Math.min(...scores),
  scoreMax: Math.max(...scores),
  targetWeightSum: targetWeights.reduce((sum, value) => sum + value, 0),
  assetWeightSum: assetWeights.reduce((sum, value) => sum + value, 0),
  scenarioProbabilitySum: scenarioProbabilities.reduce((sum, value) => sum + value, 0),
  sourceChecks,
  formulaErrors: errorScan,
  topTen: ranking.getRange("A8:B17").values,
};

console.log(JSON.stringify(result, null, 2));
