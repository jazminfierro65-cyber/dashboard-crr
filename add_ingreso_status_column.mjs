import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const workbookPath = "outputs/crr_dashboard/plantilla_carga_web_crr.xlsx";
const linkedPath = "outputs/crr_dashboard/plantilla_carga_web_crr_enlazada_minsal.xlsx";

function colName(index) {
  let name = "";
  let n = index;
  while (n > 0) {
    const rem = (n - 1) % 26;
    name = String.fromCharCode(65 + rem) + name;
    n = Math.floor((n - 1) / 26);
  }
  return name;
}

function statusFormula(row) {
  return `=IF(Q${row}="","",IF(IF(ISNUMBER(Q${row}),MOD(Q${row},1),TIMEVALUE(Q${row}))<=TIME(8,15,0),"Ingreso adecuado","Atraso"))`;
}

function categoryFormula(row) {
  return `=IF(H${row}<>"",IFERROR(VLOOKUP(IFERROR(LEFT(H${row},FIND("-",H${row})-1),H${row}),'Catalogo MINSAL'!A:E,5,FALSE),""),IF(I${row}<>"",IFERROR(VLOOKUP(IFERROR(LEFT(I${row},FIND("-",I${row})-1),I${row}),'Catalogo MINSAL'!A:E,5,FALSE),""),""))`;
}

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

const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(workbookPath));
const sheet = workbook.worksheets.getItem("Carga Web Cirugias");
const source = sheet.getRange("A3:AB500").values;

const output = source.map((row, index) => {
  if (index === 0) {
    if (row[17] === "Clasificación ingreso 08:00-08:15" && row[18] === "Clasificación ingreso 08:00-08:15") {
      return [...row.slice(0, 18), ...row.slice(19), "Observaciones"].slice(0, 28);
    }
    if (row[17] === "Clasificación ingreso 08:00-08:15") return row.slice(0, 28);
    return [...row.slice(0, 17), "Clasificación ingreso 08:00-08:15", ...row.slice(17)].slice(0, 28);
  }
  if (source[0][17] === "Clasificación ingreso 08:00-08:15" && source[0][18] === "Clasificación ingreso 08:00-08:15") {
    return [...row.slice(0, 18), ...row.slice(19), ""].slice(0, 28);
  }
  if (source[0][17] === "Clasificación ingreso 08:00-08:15") return row.slice(0, 28);
  return [...row.slice(0, 17), "", ...row.slice(17)].slice(0, 28);
});

sheet.getRange("A3:AB500").values = output;

const rowCount = 497;
sheet.getRange("J4:J500").formulas = Array.from({ length: rowCount }, (_, index) => [categoryFormula(index + 4)]);
sheet.getRange("R4:R500").formulas = Array.from({ length: rowCount }, (_, index) => [statusFormula(index + 4)]);
sheet.getRange("V4:V500").formulas = Array.from({ length: rowCount }, (_, index) => [delayFormula(index + 4)]);
sheet.getRange("W4:W500").formulas = Array.from({ length: rowCount }, (_, index) => [durationFormula(index + 4)]);
sheet.getRange("Y4:Y500").formulas = Array.from({ length: rowCount }, (_, index) => [diffFormula(index + 4)]);
sheet.getRange("Z4:Z500").formulas = Array.from({ length: rowCount }, (_, index) => [accuracyFormula(index + 4)]);

sheet.getRange("R4:R500").conditionalFormats.add("containsText", {
  text: "Ingreso adecuado",
  format: { fill: "#ddf4e9", font: { color: "#177e55", bold: true } },
});
sheet.getRange("R4:R500").conditionalFormats.add("containsText", {
  text: "Atraso",
  format: { fill: "#fff0e7", font: { color: "#e8743b", bold: true } },
});

sheet.getRange("R3:R500").format = {
  wrapText: true,
  verticalAlignment: "center",
};
sheet.getRange("R3").format = {
  fill: "#16639f",
  font: { color: "#ffffff", bold: true },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  wrapText: true,
};
sheet.getRange("V4:Y500").format.numberFormat = "0";
sheet.getRange("A:AB").format.autofitColumns();

await workbook.render({ sheetName: "Carga Web Cirugias", range: "A1:AB12", scale: 1 });
const errorScan = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "scan de errores finales",
});
console.log(errorScan.ndjson);

const exported = await SpreadsheetFile.exportXlsx(workbook);
await exported.save(workbookPath);
await exported.save(linkedPath);
console.log(JSON.stringify({ workbookPath, linkedPath, addedColumn: "R", header: output[0][17] }, null, 2));
