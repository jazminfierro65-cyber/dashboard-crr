import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const minsalPath = "/Users/jazz/Downloads/Códigos 2025 actualizado MINSAL (2).xlsx";
const outputDir = path.resolve("outputs/crr_dashboard");
const templatePath = path.join(outputDir, "plantilla_carga_web_crr.xlsx");
const outputPath = path.join(outputDir, "plantilla_carga_web_crr_enlazada_minsal.xlsx");
const webTemplatePath = path.join(outputDir, "plantilla_carga_web_crr.xlsx");

const blue = "#16639f";
const orange = "#e8743b";
const softOrange = "#fff0e7";
const border = "#d7e3ee";

function clean(value) {
  if (value === null || value === undefined) return "";
  const text = String(value).trim();
  if (/^#(NAME\?|N\/A|REF!|VALUE!|DIV\/0!)$/i.test(text)) return "";
  return text;
}

function normHeader(value) {
  return clean(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .toLowerCase();
}

function codeParts(value) {
  const original = clean(value);
  if (!original) return { original: "", base: "", additional: "" };
  const compact = original.replace(/\.0$/, "").replace(/,/g, "");
  const [baseRaw, additionalRaw = ""] = compact.split("-");
  return {
    original,
    base: clean(baseRaw).replace(/\D/g, ""),
    additional: clean(additionalRaw).replace(/\D/g, ""),
  };
}

function categoryFromActivity(activity, glosa = "") {
  const text = `${clean(activity)} ${clean(glosa)}`
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase();
  if (/\bcmay\b|mayor|cirugia mayor/.test(text)) return "Cirugía Mayor";
  if (/\bcmen\b|menor|cirugia menor/.test(text)) return "Cirugía Menor";
  if (/procedimiento|\bproc\b/.test(text)) return "Procedimiento";
  return "";
}

function findHeaderIndex(headers, candidates) {
  const normalized = headers.map(normHeader);
  for (const candidate of candidates) {
    const wanted = normHeader(candidate);
    const exact = normalized.findIndex((header) => header === wanted);
    if (exact >= 0) return exact;
    const partial = normalized.findIndex((header) => header.includes(wanted));
    if (partial >= 0) return partial;
  }
  return -1;
}

function sheetByName(workbook, names) {
  for (const name of names) {
    try {
      return workbook.worksheets.getItem(name);
    } catch {
      // Keep looking for the expected sheet name.
    }
  }
  return workbook.worksheets.getActiveWorksheet();
}

function title(sheet, range, text) {
  const r = sheet.getRange(range);
  r.merge();
  r.values = [[text]];
  r.format = {
    fill: blue,
    font: { color: "#ffffff", bold: true, size: 15 },
    horizontalAlignment: "center",
    verticalAlignment: "center",
  };
  r.format.rowHeightPx = 38;
}

function styleTable(sheet, range) {
  const r = sheet.getRange(range);
  r.format = {
    wrapText: true,
    verticalAlignment: "center",
    borders: { preset: "inside", style: "thin", color: border },
  };
  r.getRow(0).format = {
    fill: blue,
    font: { color: "#ffffff", bold: true },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
  };
}

const minsalWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(minsalPath));
const minsalSheet = sheetByName(minsalWorkbook, ["Arancel y tipo de actividad", "Arancel"]);
const minsalValues = minsalSheet.getRange("A1:H3000").values;
const minsalHeaderRow = minsalValues.findIndex((row) =>
  row.some((cell) => ["codigo", "código"].includes(normHeader(cell))),
);

if (minsalHeaderRow < 0) throw new Error("No se encontró columna de códigos en el archivo MINSAL.");

const minsalHeaders = minsalValues[minsalHeaderRow];
const codeIdx = findHeaderIndex(minsalHeaders, ["CÓDIGO", "Codigo"]);
const additionalIdx = findHeaderIndex(minsalHeaders, ["CÓDIGO ADICIONAL", "Codigo adicional"]);
const glosaIdx = findHeaderIndex(minsalHeaders, ["GLOSA", "Glosa"]);
const activityIdx = findHeaderIndex(minsalHeaders, ["TIPO DE ACTIVIDAD", "Actividad"]);

const catalogMap = new Map();
for (const row of minsalValues.slice(minsalHeaderRow + 1)) {
  const parts = codeParts(row[codeIdx]);
  if (!parts.base) continue;
  const glosa = glosaIdx >= 0 ? clean(row[glosaIdx]) : "";
  const activity = activityIdx >= 0 ? clean(row[activityIdx]) : "";
  const category = categoryFromActivity(activity, glosa);
  if (!category) continue;
  catalogMap.set(parts.base, [
    parts.base,
    additionalIdx >= 0 ? clean(row[additionalIdx]) : parts.additional,
    glosa,
    activity,
    category,
  ]);
}

const templateWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(templatePath));
const cirugias = sheetByName(templateWorkbook, ["Carga Web Cirugias"]);
const catalogoMinsal = templateWorkbook.worksheets.getOrAdd("Catalogo MINSAL");

const catalogRows = [
  ["Codigo", "Codigo Adicional", "Glosa MINSAL", "Tipo Actividad MINSAL", "Categoria"],
  ...Array.from(catalogMap.values()).sort((a, b) => a[0].localeCompare(b[0])),
];

title(catalogoMinsal, "A1:E1", "Catalogo MINSAL 2025 para clasificar cirugias CRR");
catalogoMinsal.getRange(`A3:E${catalogRows.length + 2}`).values = catalogRows;
styleTable(catalogoMinsal, `A3:E${catalogRows.length + 2}`);
catalogoMinsal.getRange("A:E").format.autofitColumns();
catalogoMinsal.freezePanes.freezeRows(3);

const categoryFormula = (row) =>
  `=IF(H${row}<>"",XLOOKUP(IFERROR(LEFT(H${row},FIND("-",H${row})-1),H${row}),'Catalogo MINSAL'!A:A,'Catalogo MINSAL'!E:E,""),IF(I${row}<>"",XLOOKUP(IFERROR(LEFT(I${row},FIND("-",I${row})-1),I${row}),'Catalogo MINSAL'!A:A,'Catalogo MINSAL'!E:E,""),""))`;

cirugias.getRange("J4:J500").formulas = Array.from({ length: 497 }, (_, index) => [categoryFormula(index + 4)]);
cirugias.getRange("J4:J500").format = {
  fill: softOrange,
  font: { color: "#10365a", bold: true },
  borders: { preset: "inside", style: "thin", color: orange },
};
cirugias.getRange("J4:J500").dataValidation = {
  allowBlank: true,
  list: { inCellDropDown: true, source: ["Cirugía Mayor", "Cirugía Menor", "Procedimiento"] },
};

cirugias.getRange("U4:U500").formulas = Array.from({ length: 497 }, (_, index) => {
  const row = index + 4;
  return [`=IFERROR(IF(OR(Q${row}="",R${row}=""),"",IF(TIMEVALUE(R${row})>=TIMEVALUE(Q${row}),(TIMEVALUE(R${row})-TIMEVALUE(Q${row}))*1440,(TIMEVALUE(R${row})+1-TIMEVALUE(Q${row}))*1440)),"")`];
});
cirugias.getRange("V4:V500").formulas = Array.from({ length: 497 }, (_, index) => {
  const row = index + 4;
  return [`=IFERROR(IF(OR(R${row}="",S${row}=""),"",IF(TIMEVALUE(S${row})>=TIMEVALUE(R${row}),(TIMEVALUE(S${row})-TIMEVALUE(R${row}))*1440,(TIMEVALUE(S${row})+1-TIMEVALUE(R${row}))*1440)),"")`];
});
cirugias.getRange("X4:X500").formulas = Array.from({ length: 497 }, (_, index) => {
  const row = index + 4;
  return [`=IFERROR(IF(OR(V${row}="",W${row}=""),"",V${row}-W${row}),"")`];
});
cirugias.getRange("Y4:Y500").formulas = Array.from({ length: 497 }, (_, index) => {
  const row = index + 4;
  return [`=IFERROR(IF(X${row}="","",IF(X${row}>15,"Subestimacion",IF(X${row}<-15,"Sobrestimacion","Asertivo"))),"")`];
});

await templateWorkbook.render({ sheetName: "Carga Web Cirugias", range: "A1:AA12", scale: 1 });
await templateWorkbook.render({ sheetName: "Catalogo MINSAL", range: "A1:E12", scale: 1 });

const errorScan = await templateWorkbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "scan de errores finales",
});
console.log(errorScan.ndjson);

await fs.mkdir(outputDir, { recursive: true });
const exported = await SpreadsheetFile.exportXlsx(templateWorkbook);
await exported.save(outputPath);
await exported.save(webTemplatePath);

console.log(JSON.stringify({ outputPath, webTemplatePath, catalogCodes: catalogRows.length - 1 }, null, 2));
