import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const workbookPath = "outputs/crr_dashboard/plantilla_carga_web_crr.xlsx";
const sourcePath = "/Users/jazz/Downloads/plantilla_carga_web_crr.xlsx";

const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(workbookPath));
const sheet = workbook.worksheets.getItem("Carga Web Cirugias");
const sourceWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(sourcePath));
const sourceSheet = sourceWorkbook.worksheets.getItem("Carga Web Cirugias");

const sourceRows = sourceSheet.getRange("A4:AA500").values;
const currentRows = sheet.getRange("A4:AB500").values;

const programmedByEpisode = new Map();
for (const row of sourceRows) {
  const episode = String(row[1] || "").trim();
  if (!episode) continue;
  programmedByEpisode.set(episode, row[22] ?? "");
}

const programmedValues = currentRows.map((row) => {
  const episode = String(row[1] || "").trim();
  return [programmedByEpisode.get(episode) ?? ""];
});

sheet.getRange("X4:X500").clear({ applyTo: "all" });
sheet.getRange("X4:X500").values = programmedValues;
sheet.getRange("X4:X500").format = {
  fill: "#fff0e7",
  font: { color: "#10365a", bold: true },
  borders: { preset: "inside", style: "thin", color: "#e8743b" },
};
sheet.getRange("X4:X500").format.numberFormat = "0";

const diffFormulas = Array.from({ length: 497 }, (_, index) => {
  const row = index + 4;
  return [`=IFERROR(IF(OR(W${row}="",X${row}=""),"",W${row}-X${row}),"")`];
});
const accuracyFormulas = Array.from({ length: 497 }, (_, index) => {
  const row = index + 4;
  return [`=IFERROR(IF(Y${row}="","",IF(Y${row}>15,"Subestimacion",IF(Y${row}<-15,"Sobrestimacion","Asertivo"))),"")`];
});
sheet.getRange("Y4:Y500").formulas = diffFormulas;
sheet.getRange("Z4:Z500").formulas = accuracyFormulas;

const errorScan = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "scan de errores finales",
});
console.log(errorScan.ndjson);

const exported = await SpreadsheetFile.exportXlsx(workbook);
await exported.save(workbookPath);
await exported.save("outputs/crr_dashboard/plantilla_carga_web_crr_enlazada_minsal.xlsx");
console.log(JSON.stringify({ repaired: true, workbookPath }, null, 2));
