import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = path.resolve("outputs/crr_dashboard");
const outputPath = path.join(outputDir, "plantilla_carga_web_crr.xlsx");

const wb = Workbook.create();
const cirugias = wb.worksheets.getOrAdd("Carga Web Cirugias", { renameFirstIfOnlyNewSpreadsheet: true });
const eventos = wb.worksheets.add("Carga Eventos Diarios");
const catalogos = wb.worksheets.add("Catalogos");
const instrucciones = wb.worksheets.add("Instrucciones");

const blue = "#16639f";
const orange = "#e8743b";
const softOrange = "#fff0e7";
const softBlue = "#e8f3fb";
const softGreen = "#ddf4e9";
const border = "#d7e3ee";

function title(sheet, range, text) {
  const r = sheet.getRange(range);
  r.merge();
  r.values = [[text]];
  r.format = {
    fill: blue,
    font: { color: "#ffffff", bold: true, size: 16 },
    horizontalAlignment: "center",
    verticalAlignment: "center",
  };
  r.format.rowHeightPx = 40;
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

const surgeryHeaders = [
  "Fecha Quirofano",
  "Episodio",
  "Paciente",
  "Apellido",
  "Edad",
  "Diagnostico",
  "Descripción Intervención",
  "Codigo 1",
  "Codigo 2",
  "Categoria",
  "Cirugia",
  "Especialidad",
  "Anestesista",
  "Cirujano",
  "Pabellon/Quirofano",
  "Estado",
  "Ingreso Pabellon",
  "Inicio Cirugia",
  "Termino Cirugia",
  "Salida Pabellon",
  "Demora Ingreso-Inicio Min",
  "Duracion Cirugia Formula Min",
  "Tiempo programado de cirugía en tabla (Min)",
  "Diferencia Real - Programado Min",
  "Asertividad",
  "Causa Suspension",
  "Observaciones",
];

title(cirugias, "A1:AA1", "Plantilla de carga web CRR - Cirugias por pabellon");
cirugias.getRange("A3:AA3").values = [surgeryHeaders];
cirugias.getRange("A4:AA8").values = [
  ["2026-05-26", "I000000001", "NOMBRE PACIENTE", "EJEMPLO", "45a 2m", "", "COLECISTECTOMIA POR VIDEOLAPAROSCOPIA", "1802004-7", "1802067-9", "", "COLECISTECTOMIA", "CIRUGIA GENERAL", "Nombre Anestesista", "Nombre Cirujano", "Quirofano 01", "Atendido", "08:15", "08:35", "09:30", "09:45", "", "", 45, "", "", "", "Fila de ejemplo editable"],
  ["2026-05-26", "I000000002", "PACIENTE", "SUSPENDIDO", "60a", "", "HERNIA INGUINAL", "1802001-2", "", "", "HERNIA INGUINAL", "CIRUGIA GENERAL", "", "Nombre Cirujano", "Quirofano 02", "Suspendido", "", "", "", "", "", "", 60, "", "", "Paciente no se presenta", "Registrar causa si estado es Suspendido"],
  ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
  ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
  ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
];
cirugias.getRange("J4:J200").formulas = Array.from({ length: 197 }, (_, i) => {
  const row = i + 4;
  return [`=IF(H${row}<>"",XLOOKUP(H${row},Catalogos!F:F,Catalogos!G:G,""),IF(I${row}<>"",XLOOKUP(I${row},Catalogos!F:F,Catalogos!G:G,""),""))`];
});
cirugias.getRange("U4:U200").formulas = Array.from({ length: 197 }, (_, i) => {
  const row = i + 4;
  return [`=IF(OR(Q${row}="",R${row}=""),"",IF(TIMEVALUE(R${row})>=TIMEVALUE(Q${row}),(TIMEVALUE(R${row})-TIMEVALUE(Q${row}))*1440,(TIMEVALUE(R${row})+1-TIMEVALUE(Q${row}))*1440))`];
});
cirugias.getRange("V4:V200").formulas = Array.from({ length: 197 }, (_, i) => {
  const row = i + 4;
  return [`=IF(OR(R${row}="",S${row}=""),"",IF(TIMEVALUE(S${row})>=TIMEVALUE(R${row}),(TIMEVALUE(S${row})-TIMEVALUE(R${row}))*1440,(TIMEVALUE(S${row})+1-TIMEVALUE(R${row}))*1440))`];
});
cirugias.getRange("X4:X200").formulas = Array.from({ length: 197 }, (_, i) => {
  const row = i + 4;
  return [`=IF(OR(V${row}="",W${row}=""),"",V${row}-W${row})`];
});
cirugias.getRange("Y4:Y200").formulas = Array.from({ length: 197 }, (_, i) => {
  const row = i + 4;
  return [`=IF(X${row}="","",IF(X${row}>15,"Subestimacion",IF(X${row}<-15,"Sobrestimacion","Asertivo")))`];
});
styleTable(cirugias, "A3:AA200");
cirugias.getRange("W4:W200").format = {
  fill: softOrange,
  font: { color: "#10365a", bold: true },
  borders: { preset: "inside", style: "thin", color: orange },
};
cirugias.getRange("U4:X200").format.numberFormat = "0";
cirugias.getRange("Y4:Y200").conditionalFormats.add("containsText", { text: "Subestimacion", format: { fill: "#fee2e2", font: { color: "#b91c1c", bold: true } } });
cirugias.getRange("Y4:Y200").conditionalFormats.add("containsText", { text: "Asertivo", format: { fill: softGreen, font: { color: "#177e55", bold: true } } });
cirugias.getRange("Y4:Y200").conditionalFormats.add("containsText", { text: "Sobrestimacion", format: { fill: softBlue, font: { color: blue, bold: true } } });
cirugias.getRange("A:AA").format.autofitColumns();
cirugias.freezePanes.freezeRows(3);

title(eventos, "A1:J1", "Plantilla de carga web CRR - Eventos diarios");
eventos.getRange("A3:J3").values = [[
  "Fecha",
  "Tipo Evento",
  "Especialidad",
  "Pabellon/Quirofano",
  "Nodo Critico",
  "Minutos",
  "Cantidad",
  "Causa",
  "Responsable",
  "Observaciones",
]];
eventos.getRange("A4:J8").values = [
  ["2026-05-26", "Demora ingreso >15", "UROLOGIA", "Quirofano 01", "Gestion de tabla", 20, 1, "Paciente ingresa tarde", "Pabellon", "Ejemplo"],
  ["2026-05-26", "Suspension", "CIRUGIA GENERAL", "Quirofano 02", "Admision / citacion", 0, 1, "Paciente no se presenta", "Admision", "Ejemplo"],
  ["", "", "", "", "", "", "", "", "", ""],
  ["", "", "", "", "", "", "", "", "", ""],
  ["", "", "", "", "", "", "", "", "", ""],
];
styleTable(eventos, "A3:J200");
eventos.getRange("A:J").format.autofitColumns();
eventos.freezePanes.freezeRows(3);

title(catalogos, "A1:G1", "Catalogos sugeridos");
catalogos.getRange("A3:D13").values = [
  ["Estados", "Tipos Evento", "Nodos Criticos", "Asertividad"],
  ["Atendido", "Demora ingreso >15", "Admision / citacion", "Asertivo"],
  ["Realizado", "Suspension", "Preoperatorio", "Subestimacion"],
  ["Recepcionado", "Recambio >15", "Anestesia", "Sobrestimacion"],
  ["Agendado", "Nodo critico", "Esterilizacion / instrumental", ""],
  ["Suspendido", "", "Disponibilidad de cama", ""],
  ["", "", "Pabellon", ""],
  ["", "", "Equipo quirurgico", ""],
  ["", "", "Traslado paciente", ""],
  ["", "", "Gestion de tabla", ""],
  ["", "", "Recuperacion", ""],
];
styleTable(catalogos, "A3:D13");
catalogos.getRange("F3:G12").values = [
  ["Codigo", "Categoria"],
  ["1802004-7", "Cirugía Mayor"],
  ["1802067-9", "Cirugía Mayor"],
  ["1802001-2", "Cirugía Menor"],
  ["PROC-001", "Procedimiento"],
  ["", ""],
  ["", ""],
  ["", ""],
  ["", ""],
  ["", ""],
];
styleTable(catalogos, "F3:G12");
catalogos.getRange("A:G").format.autofitColumns();

title(instrucciones, "A1:F1", "Como usar esta plantilla");
instrucciones.getRange("A3:F12").values = [
  ["Paso", "Accion", "Hoja", "Campo clave", "Resultado", "Nota"],
  ["1", "Copiar o pegar datos desde el sistema origen", "Carga Web Cirugias", "Fecha, paciente, pabellon, horas", "Base lista para web", "No borrar encabezados"],
  ["2", "Completar manualmente tiempo programado", "Carga Web Cirugias", "Tiempo programado de cirugía en tabla (Min)", "Calcula diferencia y asertividad", "Columna destacada en naranjo"],
  ["3", "Completar o mantener mapa codigo-categoria", "Catalogos", "Codigo y Categoria", "Categoria se enlaza con Codigo 1 o Codigo 2", "Use Cirugía Mayor, Cirugía Menor o Procedimiento"],
  ["4", "Revisar suspendidos", "Carga Web Cirugias", "Estado y Causa Suspension", "Permite resumen por pabellon", "Si Estado=Suspended/Suspendido, registrar causa"],
  ["5", "Registrar eventos operativos", "Carga Eventos Diarios", "Tipo, nodo, causa", "Demoras y nodos criticos", "Sirve para tablero diario"],
  ["6", "Exportar para cargar en web", "Excel", "Guardar como CSV UTF-8", "Archivo liviano para importar", "Exportar una hoja a la vez si la web lo solicita"],
  ["7", "Validar formulas", "Carga Web Cirugias", "Duracion, diferencia, asertividad", "Indicadores coherentes", "Asertividad = real - programado"],
  ["", "", "", "", "", ""],
  ["Regla asertividad", "Diferencia > 15 Subestimacion", "", "", "", ""],
  ["Regla asertividad", "Diferencia entre -15 y 15 Asertivo", "", "", "", ""],
];
styleTable(instrucciones, "A3:F12");
instrucciones.getRange("A:F").format.autofitColumns();

if (typeof wb.recalculate === "function") wb.recalculate();

await fs.mkdir(outputDir, { recursive: true });
const exported = await SpreadsheetFile.exportXlsx(wb);
await exported.save(outputPath);
console.log(outputPath);
