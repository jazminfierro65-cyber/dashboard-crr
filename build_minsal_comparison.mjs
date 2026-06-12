import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const minsalPath = "/Users/jazz/Downloads/Códigos 2025 actualizado MINSAL (2).xlsx";
const templatePath = "/Users/jazz/Downloads/plantilla_carga_web_crr.xlsx";
const outputDir = path.resolve("outputs/crr_dashboard");
const outputPath = path.join(outputDir, "comparativa_codigos_minsal2_crr.xlsx");
const filledTemplatePath = path.join(outputDir, "plantilla_carga_web_crr_con_codigos_minsal2.xlsx");

const blue = "#16639f";
const orange = "#e8743b";
const softBlue = "#e8f3fb";
const softOrange = "#fff0e7";
const softGreen = "#ddf4e9";
const softRed = "#fee2e2";
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
  if (!original) return { original: "", base: "", additional: "", key: "" };
  const compact = original.replace(/\.0$/, "").replace(/,/g, "");
  const [baseRaw, additionalRaw = ""] = compact.split("-");
  const base = clean(baseRaw).replace(/\D/g, "");
  const additional = clean(additionalRaw).replace(/\D/g, "");
  return {
    original,
    base,
    additional,
    key: base,
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
    const idx = normalized.findIndex((header) => header === normHeader(candidate));
    if (idx >= 0) return idx;
  }
  for (const candidate of candidates) {
    const wanted = normHeader(candidate);
    const idx = normalized.findIndex((header) => header.includes(wanted));
    if (idx >= 0) return idx;
  }
  return -1;
}

function sheetByName(workbook, preferredNames) {
  for (const name of preferredNames) {
    try {
      return workbook.worksheets.getItem(name);
    } catch {
      // Try next expected name.
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
  row.some((cell) => normHeader(cell) === "codigo" || normHeader(cell) === "código"),
);

if (minsalHeaderRow < 0) {
  throw new Error("No se encontró fila de encabezados en el archivo MINSAL.");
}

const minsalHeaders = minsalValues[minsalHeaderRow];
const minsalCodeIdx = findHeaderIndex(minsalHeaders, ["CÓDIGO", "Codigo"]);
const minsalAdditionalIdx = findHeaderIndex(minsalHeaders, ["CÓDIGO ADICIONAL", "Codigo adicional"]);
const minsalGlosaIdx = findHeaderIndex(minsalHeaders, ["GLOSA", "Glosa"]);
const minsalActivityIdx = findHeaderIndex(minsalHeaders, ["TIPO DE ACTIVIDAD", " Tipo de actividad", "Actividad"]);

if (minsalCodeIdx < 0 || minsalGlosaIdx < 0) {
  throw new Error("No se encontraron columnas mínimas CÓDIGO y GLOSA en MINSAL.");
}

const minsalByCode = new Map();
for (const row of minsalValues.slice(minsalHeaderRow + 1)) {
  const parts = codeParts(row[minsalCodeIdx]);
  if (!parts.key) continue;
  const item = {
    codigo: parts.base,
    codigoAdicional: minsalAdditionalIdx >= 0 ? clean(row[minsalAdditionalIdx]) : parts.additional,
    glosa: minsalGlosaIdx >= 0 ? clean(row[minsalGlosaIdx]) : "",
    tipoActividad: minsalActivityIdx >= 0 ? clean(row[minsalActivityIdx]) : "",
  };
  item.categoria = categoryFromActivity(item.tipoActividad, item.glosa);
  if (!minsalByCode.has(parts.key)) minsalByCode.set(parts.key, item);
}

const templateWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(templatePath));
const templateSheet = sheetByName(templateWorkbook, ["Carga Web Cirugias"]);
const templateValues = templateSheet.getRange("A3:AD250").values;
const templateHeaders = templateValues[0];

const idx = {
  fecha: findHeaderIndex(templateHeaders, ["Fecha Quirofano"]),
  episodio: findHeaderIndex(templateHeaders, ["Episodio"]),
  paciente: findHeaderIndex(templateHeaders, ["Paciente"]),
  descripcion: findHeaderIndex(templateHeaders, ["Descripción Intervención", "Descripcion Intervencion"]),
  cirugia: findHeaderIndex(templateHeaders, ["Cirugia", "Cirugía"]),
  codigo1: findHeaderIndex(templateHeaders, ["Codigo 1", "Código 1"]),
  codigo2: findHeaderIndex(templateHeaders, ["Codigo 2", "Código 2"]),
  categoria: findHeaderIndex(templateHeaders, ["Categoria", "Categoría"]),
  especialidad: findHeaderIndex(templateHeaders, ["Especialidad"]),
};

if (idx.codigo1 < 0 && idx.codigo2 < 0) {
  throw new Error("La plantilla no contiene columnas Codigo 1 o Codigo 2.");
}

const comparisons = [];
const seenUnique = new Map();
const notFound = [];
const catalogRows = [];

for (const row of templateValues.slice(1)) {
  const hasAnyData = row.some((cell) => clean(cell));
  if (!hasAnyData) continue;

  for (const [source, colIndex] of [
    ["Codigo 1", idx.codigo1],
    ["Codigo 2", idx.codigo2],
  ]) {
    if (colIndex < 0) continue;
    const parts = codeParts(row[colIndex]);
    if (!parts.key) continue;

    const minsal = minsalByCode.get(parts.key);
    const templateCategory = idx.categoria >= 0 ? clean(row[idx.categoria]) : "";
    const finalCategory = clean(templateCategory) || minsal?.categoria || "";
    const record = {
      fecha: idx.fecha >= 0 ? clean(row[idx.fecha]) : "",
      episodio: idx.episodio >= 0 ? clean(row[idx.episodio]) : "",
      paciente: idx.paciente >= 0 ? clean(row[idx.paciente]) : "",
      descripcion:
        (idx.descripcion >= 0 ? clean(row[idx.descripcion]) : "") ||
        (idx.cirugia >= 0 ? clean(row[idx.cirugia]) : ""),
      especialidad: idx.especialidad >= 0 ? clean(row[idx.especialidad]) : "",
      source,
      original: parts.original,
      compared: parts.base,
      templateCategory,
      found: minsal ? "Sí" : "No",
      minsalCode: minsal?.codigo || "",
      minsalAdditional: minsal?.codigoAdicional || parts.additional,
      minsalGlosa: minsal?.glosa || "",
      minsalActivity: minsal?.tipoActividad || "",
      minsalCategory: minsal?.categoria || "",
      finalCategory,
      observation: minsal
        ? finalCategory
          ? "Código encontrado y categoría asignada"
          : "Código encontrado sin categoría MINSAL"
        : "Código no encontrado en MINSAL",
    };
    comparisons.push(record);
    if (!minsal) notFound.push(record);
    if (!seenUnique.has(parts.key)) seenUnique.set(parts.key, record);
    if (minsal?.categoria) {
      catalogRows.push([parts.original, minsal.categoria, minsal.glosa, minsal.tipoActividad]);
    }
  }
}

const foundCount = comparisons.filter((row) => row.found === "Sí").length;
const summaryRows = [
  ["Indicador", "Valor"],
  ["Archivo MINSAL", path.basename(minsalPath)],
  ["Archivo plantilla CRR", path.basename(templatePath)],
  ["Filas con códigos comparados", comparisons.length],
  ["Códigos encontrados en MINSAL", foundCount],
  ["Códigos no encontrados", notFound.length],
  ["Códigos únicos revisados", seenUnique.size],
  ["Regla usada", "Se compara el código base antes del guion, manteniendo el código adicional como referencia"],
];

const comparisonHeaders = [
  "Fecha Quirofano",
  "Episodio",
  "Paciente",
  "Descripción Intervención",
  "Especialidad",
  "Origen Codigo",
  "Codigo Original CRR",
  "Codigo Comparado",
  "Categoria Plantilla",
  "Encontrado MINSAL",
  "Codigo MINSAL",
  "Codigo Adicional MINSAL",
  "Glosa MINSAL",
  "Tipo Actividad MINSAL",
  "Categoria MINSAL",
  "Categoria Final",
  "Observacion",
];

function comparisonToRow(record) {
  return [
    record.fecha,
    record.episodio,
    record.paciente,
    record.descripcion,
    record.especialidad,
    record.source,
    record.original,
    record.compared,
    record.templateCategory,
    record.found,
    record.minsalCode,
    record.minsalAdditional,
    record.minsalGlosa,
    record.minsalActivity,
    record.minsalCategory,
    record.finalCategory,
    record.observation,
  ];
}

const wb = Workbook.create();
const resumen = wb.worksheets.getOrAdd("Resumen", { renameFirstIfOnlyNewSpreadsheet: true });
const detalle = wb.worksheets.add("Comparativa CRR MINSAL");
const unicos = wb.worksheets.add("Codigos Unicos");
const faltantes = wb.worksheets.add("No Encontrados");

title(resumen, "A1:B1", "Resumen comparativa codigos CRR vs MINSAL 2025");
resumen.getRange(`A3:B${summaryRows.length + 2}`).values = summaryRows;
styleTable(resumen, `A3:B${summaryRows.length + 2}`);
resumen.getRange("A:B").format.autofitColumns();

title(detalle, "A1:Q1", "Comparativa de codigos solicitados en plantilla CRR");
const detailRows = [comparisonHeaders, ...comparisons.map(comparisonToRow)];
detalle.getRange(`A3:Q${detailRows.length + 2}`).values = detailRows;
styleTable(detalle, `A3:Q${detailRows.length + 2}`);
detalle.getRange("J4:J250").conditionalFormats.add("containsText", {
  text: "No",
  format: { fill: softRed, font: { color: "#b91c1c", bold: true } },
});
detalle.getRange("J4:J250").conditionalFormats.add("containsText", {
  text: "Sí",
  format: { fill: softGreen, font: { color: "#177e55", bold: true } },
});
detalle.getRange("P4:P250").format = {
  fill: softOrange,
  font: { color: "#10365a", bold: true },
  borders: { preset: "inside", style: "thin", color: orange },
};
detalle.getRange("A:Q").format.autofitColumns();
detalle.freezePanes.freezeRows(3);

title(unicos, "A1:Q1", "Codigos unicos recuperados desde MINSAL");
const uniqueRows = [comparisonHeaders, ...Array.from(seenUnique.values()).map(comparisonToRow)];
unicos.getRange(`A3:Q${uniqueRows.length + 2}`).values = uniqueRows;
styleTable(unicos, `A3:Q${uniqueRows.length + 2}`);
unicos.getRange("A:Q").format.autofitColumns();
unicos.freezePanes.freezeRows(3);

title(faltantes, "A1:Q1", "Codigos no encontrados para revision manual");
const missingRows = [comparisonHeaders, ...notFound.map(comparisonToRow)];
faltantes.getRange(`A3:Q${missingRows.length + 2}`).values = missingRows;
styleTable(faltantes, `A3:Q${Math.max(missingRows.length + 2, 4)}`);
faltantes.getRange("A:Q").format.autofitColumns();
faltantes.freezePanes.freezeRows(3);

for (const sheet of [resumen, detalle, unicos, faltantes]) {
  sheet.getRange("A1:Q250").format.font = { name: "Arial" };
}

await wb.render({ sheetName: "Resumen", range: "A1:B10", scale: 2 });
await wb.render({ sheetName: "Comparativa CRR MINSAL", range: "A1:Q8", scale: 1 });

const errorScan = await wb.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 50 },
  summary: "final formula error scan",
});
console.log(errorScan.ndjson);

await fs.mkdir(outputDir, { recursive: true });
const exported = await SpreadsheetFile.exportXlsx(wb);
await exported.save(outputPath);

const catalog = sheetByName(templateWorkbook, ["Catalogos", "Catálogos"]);
const uniqueCatalogRows = [
  ["Codigo", "Categoria", "Glosa MINSAL", "Tipo Actividad MINSAL"],
  ...Array.from(new Map(catalogRows.map((row) => [row[0], row])).values()),
];
catalog.getRange(`F3:I${uniqueCatalogRows.length + 2}`).values = uniqueCatalogRows;
styleTable(catalog, `F3:I${uniqueCatalogRows.length + 2}`);
catalog.getRange("F:I").format.autofitColumns();

await templateWorkbook.render({ sheetName: "Carga Web Cirugias", range: "A1:Z8", scale: 1 });
const filledTemplate = await SpreadsheetFile.exportXlsx(templateWorkbook);
await filledTemplate.save(filledTemplatePath);

console.log(JSON.stringify({
  outputPath,
  filledTemplatePath,
  compared: comparisons.length,
  found: foundCount,
  notFound: notFound.length,
  unique: seenUnique.size,
}, null, 2));
