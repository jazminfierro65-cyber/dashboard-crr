import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const templatePath = "/Users/jazz/Documents/New project/plantilla_carga_web_crr.xltx";
const referencePath = "/Users/jazz/Downloads/plantilla_carga_web_crr.xlsx";

function delayFormula(row) {
  return `=IFERROR(IF(OR(Q${row}="",S${row}=""),"",IF(IF(ISNUMBER(S${row}),MOD(S${row},1),TIMEVALUE(S${row}))>=IF(ISNUMBER(Q${row}),MOD(Q${row},1),TIMEVALUE(Q${row})),(IF(ISNUMBER(S${row}),MOD(S${row},1),TIMEVALUE(S${row}))-IF(ISNUMBER(Q${row}),MOD(Q${row},1),TIMEVALUE(Q${row})))*1440,(IF(ISNUMBER(S${row}),MOD(S${row},1),TIMEVALUE(S${row}))+1-IF(ISNUMBER(Q${row}),MOD(Q${row},1),TIMEVALUE(Q${row})))*1440)),"")`;
}

function durationFormula(row) {
  return `=IFERROR(IF(OR(S${row}="",T${row}=""),"",IF(IF(ISNUMBER(T${row}),MOD(T${row},1),TIMEVALUE(T${row}))>=IF(ISNUMBER(S${row}),MOD(S${row},1),TIMEVALUE(S${row})),(IF(ISNUMBER(T${row}),MOD(T${row},1),TIMEVALUE(T${row}))-IF(ISNUMBER(S${row}),MOD(S${row},1),TIMEVALUE(S${row})))*1440,(IF(ISNUMBER(T${row}),MOD(T${row},1),TIMEVALUE(T${row}))+1-IF(ISNUMBER(S${row}),MOD(S${row},1),TIMEVALUE(S${row})))*1440)),"")`;
}

function diffFormula(row) {
  return `=IFERROR(IF(OR(W${row}="",X${row}=""),"",W${row}-X${row}),"")`;
}

function accuracyFormula(row) {
  return `=IFERROR(IF(Y${row}="","",IF(Y${row}>15,"Subestimacion",IF(Y${row}<-15,"Sobrestimacion","Asertivo"))),"")`;
}

const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(templatePath));
const sheet = wb.worksheets.getItem("Carga Web Cirugias");

let programmedByEpisode = new Map();
try {
  const refWb = await SpreadsheetFile.importXlsx(await FileBlob.load(referencePath));
  const refSheet = refWb.worksheets.getItem("Carga Web Cirugias");
  const refRows = refSheet.getRange("A4:AA500").values;
  programmedByEpisode = new Map(
    refRows
      .map((row) => [String(row[1] || "").trim(), row[22] ?? ""])
      .filter(([episode]) => episode),
  );
} catch {
  programmedByEpisode = new Map();
}

const currentRows = sheet.getRange("A4:AB500").values;
const programmedValues = currentRows.map((row) => {
  const episode = String(row[1] || "").trim();
  const recovered = programmedByEpisode.get(episode);
  return [Number.isFinite(Number(recovered)) && recovered !== "" ? Number(recovered) : ""];
});

const rowCount = 497;
sheet.getRange("V4:V500").formulas = Array.from({ length: rowCount }, (_, index) => [delayFormula(index + 4)]);
sheet.getRange("W4:W500").formulas = Array.from({ length: rowCount }, (_, index) => [durationFormula(index + 4)]);
sheet.getRange("X4:X500").clear({ applyTo: "all" });
sheet.getRange("X4:X500").values = programmedValues;
sheet.getRange("Y4:Y500").formulas = Array.from({ length: rowCount }, (_, index) => [diffFormula(index + 4)]);
sheet.getRange("Z4:Z500").formulas = Array.from({ length: rowCount }, (_, index) => [accuracyFormula(index + 4)]);

sheet.getRange("Q4:Q500").format.numberFormat = "hh:mm";
sheet.getRange("S4:U500").format.numberFormat = "hh:mm";
sheet.getRange("V4:Y500").format.numberFormat = "0";
sheet.getRange("X4:X500").format = {
  fill: "#fff0e7",
  font: { color: "#10365a", bold: true },
  borders: { preset: "inside", style: "thin", color: "#e8743b" },
};
sheet.getRange("X3").format = {
  fill: "#16639f",
  font: { color: "#ffffff", bold: true },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  wrapText: true,
};
sheet.getRange("A:AB").format.autofitColumns();

await wb.render({ sheetName: "Carga Web Cirugias", range: "A1:N8", scale: 1 });
const errorScan = await wb.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "scan errores finales",
});
console.log(errorScan.ndjson);

const output = await SpreadsheetFile.exportXlsx(wb);
await output.save(templatePath);
console.log(JSON.stringify({ fixed: true, templatePath, recoveredProgrammedRows: programmedValues.filter(([v]) => v !== "").length }, null, 2));
