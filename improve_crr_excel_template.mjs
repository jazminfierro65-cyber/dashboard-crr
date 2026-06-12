import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const workbookPath = "outputs/crr_dashboard/plantilla_carga_web_crr.xlsx";
const linkedPath = "outputs/crr_dashboard/plantilla_carga_web_crr_enlazada_minsal.xlsx";

const blue = "#16639f";
const orange = "#e8743b";
const softBlue = "#e8f3fb";
const softOrange = "#fff0e7";
const softGreen = "#ddf4e9";
const softGray = "#f5f7fa";
const border = "#d7e3ee";

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

const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(workbookPath));
const sheet = workbook.worksheets.getItem("Carga Web Cirugias");

sheet.getRange("A1:AB1").merge();
sheet.getRange("A1:AB1").values = [["Plantilla de carga web CRR - ingreso diario de pabellones"]];
sheet.getRange("A1:AB1").format = {
  fill: blue,
  font: { color: "#ffffff", bold: true, size: 16 },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
sheet.getRange("A2:AB2").values = [[
  "Obligatorio",
  "Obligatorio",
  "Obligatorio",
  "Obligatorio",
  "Obligatorio",
  "Opcional",
  "Obligatorio",
  "Obligatorio",
  "Opcional",
  "Automático",
  "Opcional",
  "Obligatorio",
  "Opcional",
  "Obligatorio",
  "Obligatorio",
  "Obligatorio",
  "Obligatorio primera cirugía",
  "Automático",
  "Obligatorio si se realiza",
  "Obligatorio si se realiza",
  "Opcional",
  "Automático",
  "Automático",
  "Manual",
  "Automático",
  "Automático",
  "Obligatorio si suspendido",
  "Opcional",
]];
sheet.getRange("A2:AB2").format = {
  fill: softGray,
  font: { color: "#60768b", bold: true, size: 10 },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  wrapText: true,
};

styleTable(sheet, "A3:AB500");

const manualRequired = "A4:E500,G4:H500,L4:L500,N4:Q500,S4:T500";
const manualOptional = "F4:F500,I4:I500,K4:K500,M4:M500,U4:U500,AA4:AB500";
const manualTime = "Q4:Q500,S4:U500";
const formulaCols = "J4:J500,R4:R500,V4:W500,Y4:Z500";
const manualProgrammed = "X4:X500";

sheet.getRange(manualRequired).format = {
  fill: "#ffffff",
  borders: { preset: "inside", style: "thin", color: border },
};
sheet.getRange(manualOptional).format = {
  fill: "#fbfdff",
  borders: { preset: "inside", style: "thin", color: border },
};
sheet.getRange(manualProgrammed).format = {
  fill: softOrange,
  font: { color: "#10365a", bold: true },
  borders: { preset: "inside", style: "thin", color: orange },
};
sheet.getRange(formulaCols).format = {
  fill: softBlue,
  font: { color: "#10365a", bold: true },
  borders: { preset: "inside", style: "thin", color: "#b9d7ed" },
};

sheet.getRange("A4:A500").format.numberFormat = "yyyy-mm-dd";
sheet.getRange(manualTime).format.numberFormat = "hh:mm";
sheet.getRange("V4:Y500").format.numberFormat = "0";
sheet.getRange("X4:X500").format.numberFormat = "0";

sheet.getRange("P4:P500").dataValidation = {
  allowBlank: false,
  list: { inCellDropDown: true, source: ["Atendido", "Realizado", "Suspendido", "Agendado"] },
};
sheet.getRange("O4:O500").dataValidation = {
  allowBlank: false,
  list: {
    inCellDropDown: true,
    source: ["Quirofano 01", "Quirofano 02", "Quirofano 03", "Quirofano 04", "Quirofano 11", "Quirofano 12", "Quirofano 13"],
  },
};
sheet.getRange("J4:J500").dataValidation = {
  allowBlank: true,
  list: { inCellDropDown: true, source: ["Cirugía Mayor", "Cirugía Menor", "Procedimiento"] },
};
sheet.getRange("R4:R500").conditionalFormats.add("containsText", {
  text: "Ingreso adecuado",
  format: { fill: softGreen, font: { color: "#177e55", bold: true } },
});
sheet.getRange("R4:R500").conditionalFormats.add("containsText", {
  text: "Atraso",
  format: { fill: softOrange, font: { color: orange, bold: true } },
});
sheet.getRange("Z4:Z500").conditionalFormats.add("containsText", {
  text: "Subestimacion",
  format: { fill: "#fee2e2", font: { color: "#b91c1c", bold: true } },
});
sheet.getRange("Z4:Z500").conditionalFormats.add("containsText", {
  text: "Asertivo",
  format: { fill: softGreen, font: { color: "#177e55", bold: true } },
});
sheet.getRange("Z4:Z500").conditionalFormats.add("containsText", {
  text: "Sobrestimacion",
  format: { fill: softBlue, font: { color: blue, bold: true } },
});

sheet.getRange("A:AB").format.autofitColumns();
sheet.getRange("A1:AB3").format.rowHeightPx = 34;
sheet.freezePanes.freezeRows(3);

const guide = workbook.worksheets.getOrAdd("Guia de carga CRR");
title(guide, "A1:F1", "Guia rapida para cargar datos CRR en la web");
guide.getRange("A3:F16").values = [
  ["Paso", "Qué completar", "Columnas", "Quién lo completa", "Regla", "Impacto en web"],
  ["1", "Identificación del paciente", "Fecha, Episodio, Paciente, Apellido, Edad", "Administrativo / pabellón", "No dejar vacío", "Permite filtrar por día y auditar caso a caso"],
  ["2", "Cirugía y códigos", "Descripción Intervención, Codigo 1, Codigo 2", "Equipo quirúrgico / administrativo", "Codigo 1 es obligatorio; Codigo 2 si corresponde", "Calcula categoría quirúrgica"],
  ["3", "Categoría", "Categoria", "Automático", "Se obtiene desde Catalogo MINSAL", "Gráfico de cirugía mayor/menor/procedimiento"],
  ["4", "Equipo y pabellón", "Especialidad, Anestesista, Cirujano, Pabellon", "Pabellón", "Usar nombres consistentes", "Filtros por especialidad y resumen por pabellón"],
  ["5", "Estado", "Estado", "Pabellón", "Elegir Atendido, Realizado, Suspendido o Agendado", "Calcula suspensiones y causas"],
  ["6", "Horas", "Ingreso, Inicio, Término, Salida", "Pabellón", "Formato hh:mm", "Calcula atraso, demora y duración real"],
  ["7", "Ingreso primera hora", "Clasificación ingreso 08:00-08:15", "Automático", "Hasta 08:15 adecuado; mayor a 08:15 atraso", "Indicador operativo de primera hora"],
  ["8", "Tiempo programado", "Tiempo programado de cirugía en tabla (Min)", "Jefatura / programación", "Ingresar minutos manualmente", "Calcula asertividad"],
  ["9", "Suspensiones", "Causa Suspension", "Pabellón / jefatura", "Obligatorio si Estado = Suspendido", "Gráfico y tabla de causas"],
  ["10", "Actualizar web", "Guardar Excel", "Usuario", "Luego presionar Actualizar desde Excel en la web", "Dashboard queda actualizado"],
  ["", "", "", "", "", ""],
  ["Colores", "Blanco", "Entrada manual", "", "", ""],
  ["Colores", "Naranjo", "Dato manual crítico", "", "", ""],
];
styleTable(guide, "A3:F16");
guide.getRange("A:F").format.autofitColumns();
guide.freezePanes.freezeRows(3);

const errorScan = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "scan de errores finales",
});
console.log(errorScan.ndjson);

await workbook.render({ sheetName: "Carga Web Cirugias", range: "A1:N12", scale: 1 });
await workbook.render({ sheetName: "Guia de carga CRR", range: "A1:F16", scale: 1 });
const exported = await SpreadsheetFile.exportXlsx(workbook);
await exported.save(workbookPath);
await exported.save(linkedPath);
console.log(JSON.stringify({ workbookPath, linkedPath, guide: "Guia de carga CRR" }, null, 2));
