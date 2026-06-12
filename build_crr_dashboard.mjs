import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = path.resolve("outputs/crr_dashboard");
const outputPath = path.join(outputDir, "modelo_dashboard_cirugias_crr.xlsx");

const wb = Workbook.create();
const dashboard = wb.worksheets.getOrAdd("Dashboard CRR", { renameFirstIfOnlyNewSpreadsheet: true });
const data = wb.worksheets.add("Datos Cirugias");
const baseDiaria = wb.worksheets.add("Base Diaria CRR");
const calc = wb.worksheets.add("Calculos");
const cats = wb.worksheets.add("Catalogos");
const informe = wb.worksheets.add("Informe Semana 21");
const mejora = wb.worksheets.add("Plan Mejora");
const programa = wb.worksheets.add("Programa Diario");

const palette = {
  navy: "#17324D",
  teal: "#0F766E",
  green: "#15803D",
  amber: "#D97706",
  red: "#B91C1C",
  blue: "#2563EB",
  slate: "#334155",
  softBlue: "#EAF2F8",
  softTeal: "#E6F4F1",
  softAmber: "#FEF3C7",
  softRed: "#FEE2E2",
  softGreen: "#DCFCE7",
  border: "#CBD5E1",
  body: "#F8FAFC",
  white: "#FFFFFF",
};

function title(sheet, range, text) {
  const r = sheet.getRange(range);
  r.merge();
  r.values = [[text]];
  r.format = {
    fill: palette.navy,
    font: { color: palette.white, bold: true, size: 18 },
    horizontalAlignment: "center",
    verticalAlignment: "center",
  };
  r.format.rowHeightPx = 42;
}

function section(sheet, range, text, fill = palette.teal) {
  const r = sheet.getRange(range);
  r.merge();
  r.values = [[text]];
  r.format = {
    fill,
    font: { color: palette.white, bold: true, size: 12 },
    horizontalAlignment: "left",
    verticalAlignment: "center",
  };
  r.format.rowHeightPx = 26;
}

function styleTable(sheet, range, headerRows = 1) {
  const r = sheet.getRange(range);
  r.format = {
    fill: palette.white,
    font: { color: "#0F172A", size: 10 },
    borders: { preset: "inside", style: "thin", color: "#E2E8F0" },
    verticalAlignment: "center",
    wrapText: true,
  };
  const header = r.getResizedRange(-(r.rowCount - headerRows), 0);
  header.format = {
    fill: palette.slate,
    font: { color: palette.white, bold: true, size: 10 },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "outside", style: "thin", color: palette.border },
  };
}

function writeRows(sheet, topLeft, rows) {
  const [col, row] = topLeft.match(/([A-Z]+)([0-9]+)/).slice(1);
  const endCol = String.fromCharCode(col.charCodeAt(0) + rows[0].length - 1);
  const endRow = Number(row) + rows.length - 1;
  sheet.getRange(`${topLeft}:${endCol}${endRow}`).values = rows;
}

const causas = [
  "Paciente no se presenta",
  "Evaluacion preanestesica incompleta",
  "Examenes o imagenes pendientes",
  "Falta de cama UPC/UTI",
  "Falta de instrumental o insumo critico",
  "Falta de equipo quirurgico",
  "Urgencia desplaza tabla electiva",
  "Condicion clinica no apta",
  "Falla de pabellon/equipamiento",
  "Sobrecupo o programacion insuficiente",
];
const nodos = [
  "Admision / citacion",
  "Preoperatorio",
  "Anestesia",
  "Esterilizacion / instrumental",
  "Disponibilidad de cama",
  "Pabellon",
  "Equipo quirurgico",
  "Traslado paciente",
  "Recuperacion",
  "Gestion de tabla",
];
const especialidades = ["Traumatologia", "Cirugia general", "Ginecologia", "Urologia", "Otorrino", "Oftalmologia"];
const estados = ["Realizada", "Suspendida"];
const pabellones = ["Pab 1", "Pab 2", "Pab 3", "Pab 4"];
const responsables = ["Jefatura pabellon", "Anestesia", "Equipo quirurgico", "Admision", "Esterilizacion", "Gestion camas"];

title(cats, "A1:F1", "Catalogos y definiciones operativas");
writeRows(cats, "A3", [["Causas de suspension"], ...causas.map((x) => [x])]);
writeRows(cats, "C3", [["Nodos criticos"], ...nodos.map((x) => [x])]);
writeRows(cats, "E3", [["Especialidades"], ...especialidades.map((x) => [x])]);
writeRows(cats, "F3", [["Responsables"], ...responsables.map((x) => [x])]);
styleTable(cats, "A3:A13");
styleTable(cats, "C3:C13");
styleTable(cats, "E3:E9");
styleTable(cats, "F3:F9");
cats.getRange("A15:F22").values = [
  ["Indicador", "Formula sugerida", "Meta inicial", "Semaforo verde", "Semaforo amarillo", "Semaforo rojo"],
  ["% suspensiones", "Suspensiones / cirugias programadas", "<= 6%", "<= 6%", "6% a 10%", "> 10%"],
  ["Retraso promedio ingreso pabellon", "Promedio minutos retraso", "<= 15 min", "<= 15", "16 a 30", "> 30"],
  ["% inicio a tiempo", "Cirugias con retraso <= 10 min / realizadas", ">= 85%", ">= 85%", "70% a 84%", "< 70%"],
  ["Reprogramacion oportuna", "Suspendidas con reprogramacion / suspendidas", ">= 90%", ">= 90%", "75% a 89%", "< 75%"],
  ["Nodo critico principal", "Nodo con mayor frecuencia de suspension/retraso", "Gestion semanal", "Plan activo", "Plan parcial", "Sin responsable"],
  ["Mejora cerrada", "Acciones cerradas / acciones totales", ">= 80%", ">= 80%", "50% a 79%", "< 50%"],
  ["Uso del tablero", "Actualizar al cierre de tabla diaria", "Diario", "Diario", "2-3 veces/sem", "Sin rutina"],
];
styleTable(cats, "A15:F22");
cats.getRange("A:F").format.autofitColumns();

const headers = [
  "Fecha",
  "Mes",
  "ID Cirugia",
  "Especialidad",
  "Pabellon",
  "Equipo quirurgico",
  "Hora programada",
  "Hora ingreso pabellon",
  "Min retraso",
  "Estado",
  "Causa suspension",
  "Nodo critico",
  "Responsable",
  "Accion inmediata",
  "Reprogramada S/N",
  "Observaciones",
];
const sample = [
  ["2026-01-03", "2026-01", "CRR-0001", "Traumatologia", "Pab 1", "Equipo A", "08:00", "08:12", 12, "Realizada", "", "Traslado paciente", "Admision", "Llamado temprano a paciente", "N/A", ""],
  ["2026-01-04", "2026-01", "CRR-0002", "Cirugia general", "Pab 2", "Equipo B", "09:30", "", "", "Suspendida", "Falta de cama UPC/UTI", "Disponibilidad de cama", "Gestion camas", "Confirmar cama 24 h antes", "S", ""],
  ["2026-01-08", "2026-01", "CRR-0003", "Ginecologia", "Pab 3", "Equipo C", "10:00", "10:35", 35, "Realizada", "", "Anestesia", "Anestesia", "Checklist preanestesico", "N/A", "Ingreso tardio por evaluacion"],
  ["2026-01-12", "2026-01", "CRR-0004", "Urologia", "Pab 1", "Equipo D", "08:00", "", "", "Suspendida", "Examenes o imagenes pendientes", "Preoperatorio", "Equipo quirurgico", "Bloqueo si examenes no estan listos", "S", ""],
  ["2026-01-18", "2026-01", "CRR-0005", "Otorrino", "Pab 4", "Equipo E", "11:00", "11:08", 8, "Realizada", "", "Pabellon", "Jefatura pabellon", "Preparacion sala previa", "N/A", ""],
  ["2026-02-02", "2026-02", "CRR-0006", "Oftalmologia", "Pab 2", "Equipo F", "08:15", "08:19", 4, "Realizada", "", "Admision / citacion", "Admision", "Confirmacion telefonica", "N/A", ""],
  ["2026-02-05", "2026-02", "CRR-0007", "Traumatologia", "Pab 3", "Equipo A", "09:00", "", "", "Suspendida", "Falta de instrumental o insumo critico", "Esterilizacion / instrumental", "Esterilizacion", "Kit critico validado dia previo", "N", ""],
  ["2026-02-09", "2026-02", "CRR-0008", "Cirugia general", "Pab 1", "Equipo B", "08:30", "09:05", 35, "Realizada", "", "Gestion de tabla", "Jefatura pabellon", "Brief de tabla 07:30", "N/A", ""],
  ["2026-02-14", "2026-02", "CRR-0009", "Ginecologia", "Pab 4", "Equipo C", "12:00", "12:18", 18, "Realizada", "", "Traslado paciente", "Admision", "Coordinar traslado con hospitalizacion", "N/A", ""],
  ["2026-02-19", "2026-02", "CRR-0010", "Urologia", "Pab 2", "Equipo D", "10:30", "", "", "Suspendida", "Condicion clinica no apta", "Preoperatorio", "Anestesia", "Filtro clinico 48 h antes", "S", ""],
  ["2026-03-03", "2026-03", "CRR-0011", "Otorrino", "Pab 3", "Equipo E", "08:00", "08:45", 45, "Realizada", "", "Anestesia", "Anestesia", "Prechequeo sala recuperacion", "N/A", ""],
  ["2026-03-06", "2026-03", "CRR-0012", "Oftalmologia", "Pab 1", "Equipo F", "09:45", "09:50", 5, "Realizada", "", "Pabellon", "Jefatura pabellon", "Apertura de sala anticipada", "N/A", ""],
  ["2026-03-10", "2026-03", "CRR-0013", "Traumatologia", "Pab 2", "Equipo A", "11:00", "", "", "Suspendida", "Urgencia desplaza tabla electiva", "Gestion de tabla", "Jefatura pabellon", "Cupo protegido electivo", "S", ""],
  ["2026-03-15", "2026-03", "CRR-0014", "Cirugia general", "Pab 4", "Equipo B", "08:20", "08:32", 12, "Realizada", "", "Esterilizacion / instrumental", "Esterilizacion", "Trazabilidad de cajas", "N/A", ""],
  ["2026-03-21", "2026-03", "CRR-0015", "Ginecologia", "Pab 1", "Equipo C", "10:10", "", "", "Suspendida", "Paciente no se presenta", "Admision / citacion", "Admision", "Confirmacion 72/24 h", "N", ""],
  ["2026-04-01", "2026-04", "CRR-0016", "Urologia", "Pab 3", "Equipo D", "08:00", "08:06", 6, "Realizada", "", "Traslado paciente", "Admision", "Camillero asignado por bloque", "N/A", ""],
  ["2026-04-04", "2026-04", "CRR-0017", "Otorrino", "Pab 4", "Equipo E", "09:30", "", "", "Suspendida", "Falta de equipo quirurgico", "Equipo quirurgico", "Confirmar roster 48 h antes", "S", ""],
  ["2026-04-08", "2026-04", "CRR-0018", "Oftalmologia", "Pab 2", "Equipo F", "10:00", "10:25", 25, "Realizada", "", "Pabellon", "Jefatura pabellon", "Lista de preparacion de sala", "N/A", ""],
  ["2026-04-12", "2026-04", "CRR-0019", "Traumatologia", "Pab 1", "Equipo A", "08:45", "09:20", 35, "Realizada", "", "Disponibilidad de cama", "Gestion camas", "Prealerta camas postoperatorias", "N/A", ""],
  ["2026-04-18", "2026-04", "CRR-0020", "Cirugia general", "Pab 3", "Equipo B", "11:15", "", "", "Suspendida", "Falla de pabellon/equipamiento", "Pabellon", "Jefatura pabellon", "Mantencion preventiva", "S", ""],
];

title(data, "A1:P1", "Base de datos diaria - cirugias CRR");
data.getRange("A2:P2").values = [headers];
data.getRange(`A3:P${sample.length + 2}`).values = sample;
styleTable(data, `A2:P${sample.length + 2}`);
data.getRange("A3:A200").format.numberFormat = "yyyy-mm-dd";
data.getRange("I3:I200").format.numberFormat = "0";
data.getRange("D3:D200").dataValidation = { allowBlank: true, list: { inCellDropDown: true, source: especialidades } };
data.getRange("E3:E200").dataValidation = { allowBlank: true, list: { inCellDropDown: true, source: pabellones } };
data.getRange("J3:J200").dataValidation = { allowBlank: true, list: { inCellDropDown: true, source: estados } };
data.getRange("K3:K200").dataValidation = { allowBlank: true, list: { inCellDropDown: true, source: causas } };
data.getRange("L3:L200").dataValidation = { allowBlank: true, list: { inCellDropDown: true, source: nodos } };
data.getRange("M3:M200").dataValidation = { allowBlank: true, list: { inCellDropDown: true, source: responsables } };
data.getRange("O3:O200").dataValidation = { allowBlank: true, list: { inCellDropDown: true, source: ["S", "N", "N/A"] } };
data.getRange("A:P").format.autofitColumns();
data.freezePanes.freezeRows(2);
data.getRange("I3:I200").conditionalFormats.addCellIs({
  operator: "greaterThan",
  formula: 30,
  format: { fill: palette.softRed, font: { color: palette.red, bold: true } },
});

title(baseDiaria, "A1:P1", "Base diaria CRR - indicadores para actualizacion continua");
baseDiaria.getRange("A2:P2").values = [[
  "Fecha",
  "Semana",
  "Fuente",
  "Programados",
  "Operados",
  "Suspendidos",
  "% suspension",
  "Retraso primera hora min",
  "Ocupacion diaria",
  "Rendimiento pac/dia",
  "Promedio recambio min",
  "Subprogramacion min",
  "Recambios total",
  "Recambios cumplen <15",
  "Recambios no cumplen",
  "Observacion ejecutiva",
]];
baseDiaria.getRange("A3:P7").values = [
  ["2026-05-18", 21, "_ANALISIS CRR SEMANA 21.pdf", 13, 11, 2, "", 33, "", "", "", "", "", "", "", "Oftalmologia y TMT infantil con suspension; retraso por instrumental"],
  ["2026-05-19", 21, "_ANALISIS CRR SEMANA 21.pdf", 12, 10, 1, "", "", "", "", "", "", "", "", "", "Recambio afectado por piso humedo y paciente en baño"],
  ["2026-05-20", 21, "_ANALISIS CRR SEMANA 21.pdf", 10, 8, 2, "", 52.5, "", "", "", "", "", "", "", "Retrasos por reunion anestesia, instrumental y cirujano en policlinico"],
  ["2026-05-22", 21, "_ANALISIS CRR SEMANA 21.pdf", 13, 12, 2, "", "", "", "", "", "", "", "", "", "Revisión de cajas y suspension TMT mano"],
  ["Semana 21", 21, "_ANALISIS CRR SEMANA 21.pdf", 48, 41, 7, 0.1458, 24.86, 1.014, 5.86, 17.78, 2.14, 41, 18, 15, "Resumen semanal extraido del informe CRR semana 21"],
];
styleTable(baseDiaria, "A2:P7");
baseDiaria.getRange("A3:A6").format.numberFormat = "yyyy-mm-dd";
baseDiaria.getRange("G3:G200").format.numberFormat = "0.0%";
baseDiaria.getRange("I3:I200").format.numberFormat = "0.0%";
baseDiaria.getRange("A:P").format.autofitColumns();
baseDiaria.freezePanes.freezeRows(2);
baseDiaria.getRange("G3:G200").conditionalFormats.addCellIs({
  operator: "greaterThan",
  formula: 0.1,
  format: { fill: palette.softRed, font: { color: palette.red, bold: true } },
});
baseDiaria.getRange("H3:H200").conditionalFormats.addCellIs({
  operator: "greaterThan",
  formula: 15,
  format: { fill: palette.softAmber, font: { color: palette.amber, bold: true } },
});

title(informe, "A1:H1", "Informe CRR Semana 21 incorporado a la base");
informe.getRange("A3:D12").values = [
  ["Indicador", "Valor semana 21", "Semana anterior", "Referencia / meta"],
  ["Pacientes programados", 48, "", ""],
  ["Pacientes operados", 41, "", ""],
  ["Retraso primera hora", 24.86, 37.44, "< 15 min"],
  ["Ocupacion diaria", 1.014, 0.865, "> 75%"],
  ["Rendimiento diario", 5.86, 5.22, "> 6 pac/dia"],
  ["Promedio recambio", 17.78, 17.21, "< 15 min"],
  ["Subprogramacion diaria", 2.14, 17.7, ""],
  ["Suspensiones", 0.1458, 0.2034, "<= 6% sugerido"],
  ["Total recambios", 41, "", ""],
];
styleTable(informe, "A3:D12");
informe.getRange("B6:C6").format.numberFormat = "0.0%";
informe.getRange("B11:C11").format.numberFormat = "0.0%";
section(informe, "A14:H14", "Eventos del informe", palette.blue);
informe.getRange("A15:H35").values = [
  ["Fecha", "Tipo evento", "Especialidad", "Motivo / causa", "Cantidad", "Programados", "Nodo critico sugerido", "Observacion"],
  ["2026-05-18", "Retraso primera hora", "Oftalmologia", "Pendiente entrega instrumental", 1, "", "Esterilizacion / instrumental", "Hora informada 08:33"],
  ["2026-05-20", "Retraso primera hora", "Otorrino", "Reunion de anestesia", 1, "", "Anestesia", "Pabellon comienza a las 09:00 hrs"],
  ["2026-05-20", "Retraso primera hora", "Oftalmologia", "Pendiente entrega instrumental", 1, "", "Esterilizacion / instrumental", "Hora informada 08:45"],
  ["2026-05-19", "Recambio concesionaria", "", "Piso humedo", 1, 6, "Pabellon", "16,7% recambio afectado"],
  ["2026-05-19", "Recambio concesionaria", "", "Piso humedo", 1, 7, "Pabellon", "14,3% recambio afectado"],
  ["2026-05-18", "Recambio otros", "Servicio de traumatologia", "Revision de cajas", 5, "", "Esterilizacion / instrumental", "Se repite en informe"],
  ["2026-05-19", "Recambio otros", "Servicio de cirugia adulto", "Paciente en el baño", 1, "", "Traslado paciente", "Paciente ingresa al baño antes de pabellon"],
  ["2026-05-20", "Recambio otros", "Unidad de oftalmologia", "Cirujano en policlinico", 4, "", "Equipo quirurgico", "Se repite en informe"],
  ["2026-05-22", "Recambio otros", "Servicio de traumatologia", "Revision de cajas", 3, "", "Esterilizacion / instrumental", "Se repite en informe"],
  ["2026-05-18", "Suspension", "Oftalmologia", "Prolongacion de tabla", 1, 7, "Gestion de tabla", "Asertivo 6; suspendido 1"],
  ["2026-05-18", "Suspension", "TMT infantil", "Estudio incompleto", 1, 6, "Preoperatorio", "Asertivo 3; sobrestimacion 2; suspendido 1"],
  ["2026-05-19", "Suspension", "C. general", "Penultima cirugia termina 18:57", 1, 6, "Gestion de tabla", "Subestimacion y prolongacion de tabla"],
  ["2026-05-20", "Suspension", "Otorrino", "Penultima cirugia termina 17:30", 2, 5, "Gestion de tabla", "Bloque quirurgico hasta 17:00 hrs"],
  ["2026-05-22", "Suspension", "TMT mano", "Penultima cirugia termina 17:37", 2, 6, "Gestion de tabla", "Suspendido 2"],
  ["", "", "", "", "", "", "", ""],
  ["Uso", "Actualizar esta tabla al cargar informes futuros", "", "", "", "", "", "Cada fila debe representar un evento o causa accionable"],
  ["Validacion", "La suma de suspensiones evento = 7", "", "", "", "", "", "Coincide con 48 programados y 14,58%"],
  ["Validacion", "Recambios no cumplen = 15", "", "", "", "", "", "2 concesionaria + 13 otros motivos"],
  ["Validacion", "Cumplen estandar = 18", "", "", "", "", "", "43,90% del total"],
  ["Fuente", "/Users/jazz/Downloads/_ANALISIS CRR SEMANA  21.pdf", "", "", "", "", "", "PDF original del informe"],
];
styleTable(informe, "A15:H35");
informe.getRange("A16:A30").format.numberFormat = "yyyy-mm-dd";
informe.getRange("A:H").format.autofitColumns();
informe.freezePanes.freezeRows(15);

title(calc, "A1:M1", "Calculos para dashboard");
calc.getRange("A3:M3").values = [[
  "Mes",
  "Programadas",
  "Realizadas",
  "Suspendidas",
  "% susp.",
  "Retraso prom.",
  "% a tiempo",
  "Reprog. oportuna",
  "Meta susp.",
  "Meta retraso",
  "Meta a tiempo",
  "Meta reprogram.",
  "Alerta",
]];
const months = ["2026-01", "2026-02", "2026-03", "2026-04"];
calc.getRange("A4:A7").values = months.map((m) => [m]);
calc.getRange("B4:M7").formulas = months.map((m, idx) => {
  const row = idx + 4;
  return [
    `=COUNTIF('Datos Cirugias'!$B$3:$B$200,A${row})`,
    `=COUNTIFS('Datos Cirugias'!$B$3:$B$200,A${row},'Datos Cirugias'!$J$3:$J$200,"Realizada")`,
    `=COUNTIFS('Datos Cirugias'!$B$3:$B$200,A${row},'Datos Cirugias'!$J$3:$J$200,"Suspendida")`,
    `=IFERROR(D${row}/B${row},0)`,
    `=IFERROR(SUMIFS('Datos Cirugias'!$I$3:$I$200,'Datos Cirugias'!$B$3:$B$200,A${row},'Datos Cirugias'!$J$3:$J$200,"Realizada")/C${row},0)`,
    `=IFERROR(COUNTIFS('Datos Cirugias'!$B$3:$B$200,A${row},'Datos Cirugias'!$J$3:$J$200,"Realizada",'Datos Cirugias'!$I$3:$I$200,"<=10")/C${row},0)`,
    `=IFERROR(COUNTIFS('Datos Cirugias'!$B$3:$B$200,A${row},'Datos Cirugias'!$J$3:$J$200,"Suspendida",'Datos Cirugias'!$O$3:$O$200,"S")/D${row},0)`,
    0.06,
    15,
    0.85,
    0.9,
    `=IF(OR(E${row}>I${row},F${row}>J${row},G${row}<K${row}),"Revisar","OK")`,
  ];
});
styleTable(calc, "A3:M7");
calc.getRange("E4:E7").format.numberFormat = "0.0%";
calc.getRange("G4:H7").format.numberFormat = "0.0%";
calc.getRange("I4:I7").format.numberFormat = "0.0%";
calc.getRange("K4:L7").format.numberFormat = "0.0%";
calc.getRange("A10:C10").values = [["Causa suspension", "Casos", "% del total"]];
calc.getRange("A11:A20").values = causas.map((x) => [x]);
calc.getRange("B11:C20").formulas = causas.map((_, i) => {
  const row = i + 11;
  return [
    `=COUNTIF('Datos Cirugias'!$K$3:$K$200,A${row})`,
    `=IFERROR(B${row}/SUM($B$11:$B$20),0)`,
  ];
});
styleTable(calc, "A10:C20");
calc.getRange("C11:C20").format.numberFormat = "0.0%";
calc.getRange("E10:G10").values = [["Nodo critico", "Casos", "% del total"]];
calc.getRange("E11:E20").values = nodos.map((x) => [x]);
calc.getRange("F11:G20").formulas = nodos.map((_, i) => {
  const row = i + 11;
  return [
    `=COUNTIF('Datos Cirugias'!$L$3:$L$200,E${row})`,
    `=IFERROR(F${row}/SUM($F$11:$F$20),0)`,
  ];
});
styleTable(calc, "E10:G20");
calc.getRange("G11:G20").format.numberFormat = "0.0%";
calc.getRange("I10:K10").values = [["Especialidad", "Suspendidas", "Retraso prom."]];
calc.getRange("I11:I16").values = especialidades.map((x) => [x]);
calc.getRange("J11:K16").formulas = especialidades.map((_, i) => {
  const row = i + 11;
  return [
    `=COUNTIFS('Datos Cirugias'!$D$3:$D$200,I${row},'Datos Cirugias'!$J$3:$J$200,"Suspendida")`,
    `=IFERROR(SUMIFS('Datos Cirugias'!$I$3:$I$200,'Datos Cirugias'!$D$3:$D$200,I${row},'Datos Cirugias'!$J$3:$J$200,"Realizada")/COUNTIFS('Datos Cirugias'!$D$3:$D$200,I${row},'Datos Cirugias'!$J$3:$J$200,"Realizada"),0)`,
  ];
});
styleTable(calc, "I10:K16");
calc.getRange("A:M").format.autofitColumns();

title(dashboard, "A1:M1", "Dashboard CRR - suspension, retrasos, nodos criticos y mejoramiento");
dashboard.getRange("A2:M2").values = [["Actualizacion sugerida: cierre diario de tabla quirurgica | Datos de ejemplo incluidos, reemplazables en 'Datos Cirugias'"]];
dashboard.getRange("A2:M2").merge();
dashboard.getRange("A2:M2").format = { fill: palette.softBlue, font: { color: palette.slate, italic: true }, horizontalAlignment: "center" };

dashboard.getRange("A4:B6").values = [["Cirugias programadas", ""], ["Realizadas", ""], ["Suspendidas", ""]];
dashboard.getRange("C4:D6").values = [["% suspensiones", ""], ["Retraso prom. min", ""], ["% inicio a tiempo", ""]];
dashboard.getRange("E4:F6").values = [["Reprogramacion oportuna", ""], ["Nodo critico principal", ""], ["Causa principal", ""]];
dashboard.getRange("B4:B6").formulas = [["=SUM(Calculos!B4:B7)"], ["=SUM(Calculos!C4:C7)"], ["=SUM(Calculos!D4:D7)"]];
dashboard.getRange("D4:D6").formulas = [["=IFERROR(B6/B4,0)"], ["=IFERROR(AVERAGE(Calculos!F4:F7),0)"], ["=IFERROR(SUMPRODUCT(Calculos!C4:C7,Calculos!G4:G7)/SUM(Calculos!C4:C7),0)"]];
dashboard.getRange("F4:F6").formulas = [["=IFERROR(SUMPRODUCT(Calculos!D4:D7,Calculos!H4:H7)/SUM(Calculos!D4:D7),0)"], ["=INDEX(Calculos!E11:E20,MATCH(MAX(Calculos!F11:F20),Calculos!F11:F20,0))"], ["=INDEX(Calculos!A11:A20,MATCH(MAX(Calculos!B11:B20),Calculos!B11:B20,0))"]];
dashboard.getRange("A4:F6").format = {
  fill: palette.white,
  borders: { preset: "outside", style: "thin", color: palette.border },
  verticalAlignment: "center",
  wrapText: true,
};
dashboard.getRange("A4:A6").format.fill = palette.softTeal;
dashboard.getRange("C4:C6").format.fill = palette.softAmber;
dashboard.getRange("E4:E6").format.fill = palette.softBlue;
dashboard.getRange("B4:B6").format.font = { bold: true, size: 14, color: palette.teal };
dashboard.getRange("D4:D6").format.font = { bold: true, size: 14, color: palette.amber };
dashboard.getRange("F4:F6").format.font = { bold: true, size: 12, color: palette.blue };
dashboard.getRange("D4").format.numberFormat = "0.0%";
dashboard.getRange("D6").format.numberFormat = "0.0%";
dashboard.getRange("F4").format.numberFormat = "0.0%";

section(dashboard, "A8:F8", "Tendencia mensual");
dashboard.getRange("A9:H13").values = [["Mes", "Programadas", "Suspendidas", "% susp.", "Retraso prom.", "% a tiempo", "Reprog. oportuna", "Alerta"]];
dashboard.getRange("A10:A13").formulas = [["=Calculos!A4"], ["=Calculos!A5"], ["=Calculos!A6"], ["=Calculos!A7"]];
dashboard.getRange("B10:H13").formulas = [4, 5, 6, 7].map((r) => [
  `=Calculos!B${r}`,
  `=Calculos!D${r}`,
  `=Calculos!E${r}`,
  `=Calculos!F${r}`,
  `=Calculos!G${r}`,
  `=Calculos!H${r}`,
  `=Calculos!M${r}`,
]);
styleTable(dashboard, "A9:H13");
dashboard.getRange("D10:D13").format.numberFormat = "0.0%";
dashboard.getRange("F10:G13").format.numberFormat = "0.0%";
dashboard.getRange("H10:H13").conditionalFormats.add('containsText', {
  text: "Revisar",
  format: { fill: palette.softRed, font: { color: palette.red, bold: true } },
});

section(dashboard, "A16:C16", "Principales causas de suspension", palette.amber);
dashboard.getRange("A17:C22").formulas = [
  ["=Calculos!A10", "=Calculos!B10", "=Calculos!C10"],
  ...[11, 12, 13, 14, 15].map((r) => [`=Calculos!A${r}`, `=Calculos!B${r}`, `=Calculos!C${r}`]),
];
styleTable(dashboard, "A17:C22");
dashboard.getRange("C18:C22").format.numberFormat = "0.0%";

section(dashboard, "E16:G16", "Nodos criticos", palette.blue);
dashboard.getRange("E17:G22").formulas = [
  ["=Calculos!E10", "=Calculos!F10", "=Calculos!G10"],
  ...[11, 12, 13, 14, 15].map((r) => [`=Calculos!E${r}`, `=Calculos!F${r}`, `=Calculos!G${r}`]),
];
styleTable(dashboard, "E17:G22");
dashboard.getRange("G18:G22").format.numberFormat = "0.0%";

section(dashboard, "I4:M4", "Lectura ejecutiva y foco de mejoramiento", palette.green);
dashboard.getRange("I5:M11").values = [
  ["Pregunta de gestion", "Lectura sugerida", "Decision esperada", "Frecuencia", "Duenio"],
  ["Que causa suspende mas?", "Atacar el mayor Pareto de suspensiones.", "Plan de accion especifico por causa", "Semanal", "Jefatura pabellon"],
  ["Donde se origina el retraso?", "Identificar nodo con mas eventos.", "Intervenir el nodo, no solo el evento", "Semanal", "CRR + nodo"],
  ["La tabla inicia a tiempo?", "Medir retraso promedio y % a tiempo.", "Ajustar citacion, traslado y briefing", "Diario", "Pabellon"],
  ["Se reprograma oportunamente?", "Evitar perdida de oportunidad quirurgica.", "Reprogramar con fecha y responsable", "Diario", "Admision"],
  ["Hay mejora sostenida?", "Cerrar acciones con indicador asociado.", "Escalar acciones vencidas", "Quincenal", "Direccion CRR"],
  ["Que se informa?", "KPIs + causa + nodo + accion.", "Reporte breve a comite quirurgico", "Mensual", "CRR"],
];
styleTable(dashboard, "I5:M11");
section(dashboard, "I13:M13", "Informe Semana 21 integrado", palette.blue);
dashboard.getRange("I14:M19").values = [
  ["Indicador", "Valor", "Meta", "Estado", "Uso diario"],
  ["Programados / operados", "48 / 41", "", "Base cargada", "Comparar contra tabla diaria"],
  ["Suspension semanal", "14,58%", "<= 6%", "Revisar", "Abrir causa y responsable"],
  ["Retraso primera hora", "24,86 min", "< 15 min", "Revisar", "Registrar motivo antes de cierre"],
  ["Ocupacion diaria", "101,4%", "> 75%", "OK", "Validar sobrecarga de tabla"],
  ["Promedio recambio", "17,78 min", "< 15 min", "Revisar", "Separar concesionaria vs otros"],
];
styleTable(dashboard, "I14:M19");
dashboard.getRange("L16:L19").conditionalFormats.add('containsText', {
  text: "Revisar",
  format: { fill: palette.softRed, font: { color: palette.red, bold: true } },
});
dashboard.getRange("A:M").format.columnWidthPx = 112;
dashboard.getRange("A1:M30").format.rowHeightPx = 24;
dashboard.freezePanes.freezeRows(2);

dashboard.charts.add("ColumnClustered", dashboard.getRange("A9:D13"), "Auto").setPosition(dashboard.getRange("A24:F36"));
dashboard.charts.add("ColumnClustered", dashboard.getRange("A17:B22"), "Auto").setPosition(dashboard.getRange("H24:M36"));
dashboard.charts.add("BarClustered", dashboard.getRange("E17:F22"), "Auto").setPosition(dashboard.getRange("A38:F50"));
dashboard.charts.add("Line", dashboard.getRange("A9:A13,E9:F13"), "Auto").setPosition(dashboard.getRange("H38:M50"));

title(mejora, "A1:J1", "Plan de mejoramiento CRR");
mejora.getRange("A3:J3").values = [["ID", "Problema / brecha", "Causa raiz probable", "Nodo", "Accion de mejora", "Responsable", "Fecha inicio", "Fecha compromiso", "Estado", "Indicador asociado"]];
mejora.getRange("A4:J11").values = [
  ["M-001", "Suspension por cama UPC/UTI", "Reserva no confirmada con anticipacion", "Disponibilidad de cama", "Huddle diario camas-pabellon antes de validar tabla", "Gestion camas", "2026-04-01", "2026-04-30", "En curso", "% suspensiones por cama"],
  ["M-002", "Retraso por traslado", "Paciente no disponible al llamado", "Traslado paciente", "Prellamado a hospitalizacion 30 min antes", "Admision", "2026-04-05", "2026-04-20", "Cerrada", "Retraso prom. min"],
  ["M-003", "Instrumental incompleto", "Checklist no aplicado dia previo", "Esterilizacion / instrumental", "Validar cajas criticas a las 16:00 del dia anterior", "Esterilizacion", "2026-04-08", "2026-05-05", "En curso", "Suspensiones por instrumental"],
  ["M-004", "Evaluacion preanestesica tardia", "Informacion clinica incompleta", "Anestesia", "Filtro preoperatorio 48-72 h antes", "Anestesia", "2026-04-10", "2026-05-10", "Pendiente", "% inicio a tiempo"],
  ["M-005", "Sobrecupo de tabla", "Duracion quirurgica subestimada", "Gestion de tabla", "Revisar duracion historica por especialidad", "Jefatura pabellon", "2026-04-12", "2026-05-15", "Pendiente", "% suspensiones"],
  ["M-006", "Paciente no se presenta", "Confirmacion insuficiente", "Admision / citacion", "Confirmacion 72 h y 24 h con registro", "Admision", "2026-04-15", "2026-05-01", "En curso", "No presentacion"],
  ["M-007", "Equipo quirurgico no disponible", "Cambios de roster no comunicados", "Equipo quirurgico", "Confirmacion de equipo 48 h antes", "Equipo quirurgico", "2026-04-18", "2026-05-08", "Pendiente", "Suspensiones por equipo"],
  ["M-008", "Fallas de equipamiento", "Mantencion preventiva incompleta", "Pabellon", "Calendario mensual de mantencion y respaldo", "Jefatura pabellon", "2026-04-20", "2026-05-20", "En curso", "Fallas pabellon"],
];
styleTable(mejora, "A3:J11");
mejora.getRange("I4:I100").dataValidation = { allowBlank: true, list: { inCellDropDown: true, source: ["Pendiente", "En curso", "Cerrada", "Escalada"] } };
mejora.getRange("D4:D100").dataValidation = { allowBlank: true, list: { inCellDropDown: true, source: nodos } };
mejora.getRange("F4:F100").dataValidation = { allowBlank: true, list: { inCellDropDown: true, source: responsables } };
mejora.getRange("I4:I100").conditionalFormats.add('containsText', { text: "Cerrada", format: { fill: palette.softGreen, font: { color: palette.green, bold: true } } });
mejora.getRange("I4:I100").conditionalFormats.add('containsText', { text: "Pendiente", format: { fill: palette.softAmber, font: { color: palette.amber, bold: true } } });
mejora.getRange("I4:I100").conditionalFormats.add('containsText', { text: "Escalada", format: { fill: palette.softRed, font: { color: palette.red, bold: true } } });
mejora.getRange("A:J").format.autofitColumns();
mejora.freezePanes.freezeRows(3);

title(programa, "A1:J1", "Programa de actualizacion diaria CRR");
programa.getRange("A3:J11").values = [
  ["Paso", "Hora sugerida", "Responsable", "Entrada requerida", "Donde registrar", "Validacion", "Salida esperada", "Semaforo", "Escalamiento", "Nota"],
  ["1. Cierre de tabla", "Fin de jornada", "Pabellon", "Programados, operados, suspendidos", "Base Diaria CRR", "Programados = operados + suspendidos + pendientes justificados", "Fila diaria completa", "Obligatorio", "Jefatura pabellon", "Registrar aunque no existan suspensiones"],
  ["2. Retraso primera hora", "Antes de 10:00", "Pabellon / Anestesia", "Hora ingreso y motivo si >15 min", "Informe Semana 21 / Base Diaria", "Motivo y nodo critico no vacios", "Causa accionable", ">15 min", "Anestesia o nodo responsable", "Separar instrumental, anestesia, traslado y equipo"],
  ["3. Recambio", "Durante jornada", "Pabellon", "Promedio recambio y eventos >=15 min", "Informe Semana 21", "Clasificar concesionaria u otros", "Pareto recambio", ">15 min", "Jefatura pabellon", "Usar cantidad si el informe agrupa eventos"],
  ["4. Suspensiones", "Cierre diario", "CRR / Admision", "Fecha, especialidad, causa, cantidad", "Informe Semana 21", "Causa y nodo critico asignados", "Pareto suspension", ">6%", "Comite quirurgico", "Registrar reprogramacion si existe"],
  ["5. Plan mejora", "Semanal", "Duenio de nodo", "Accion, responsable, fecha compromiso", "Plan Mejora", "Accion vencida debe estar escalada", "Acciones cerradas o escaladas", "Vencida", "Direccion CRR", "Relacionar cada accion con indicador"],
  ["6. Revision ejecutiva", "Semanal", "CRR", "Dashboard actualizado", "Dashboard CRR", "Sin errores y con datos de semana", "Reporte breve", "Revisar", "Direccion", "Usar lectura ejecutiva del dashboard"],
  ["7. Carga de PDF", "Cuando llegue informe", "Analista CRR", "Informe PDF semanal", "Informe Semana 21 o nueva hoja semanal", "Cuadrar programados, operados y suspensiones", "Base historica ampliada", "Diferencias", "CRR", "Mantener fuente visible"],
  ["8. Calidad de datos", "Mensual", "CRR", "Revision de campos vacios", "Datos Cirugias / Base Diaria", "Causas, nodos y responsables completos", "Base confiable", ">5% vacios", "Jefatura", "Depurar catalogos cada mes"],
];
styleTable(programa, "A3:J11");
programa.getRange("A13:F20").values = [
  ["Campo minimo diario", "Tipo", "Obligatorio", "Ejemplo", "Validacion", "Comentario"],
  ["Fecha", "Fecha", "Si", "2026-05-27", "Formato fecha", "Una fila por dia o por evento"],
  ["Programados", "Numero", "Si", 12, ">= operados + suspendidos", "Permite calcular tasa"],
  ["Operados", "Numero", "Si", 10, ">=0", "Volumen real"],
  ["Suspendidos", "Numero", "Si", 2, ">=0", "Debe tener causa si >0"],
  ["Retraso primera hora", "Numero min", "Si", 24.86, "Meta <15", "Usar promedio diario o evento"],
  ["Nodo critico", "Lista", "Si si hay evento", "Gestion de tabla", "Catalogo", "Debe quedar asociado a responsable"],
  ["Accion inmediata", "Texto", "Si si hay alerta", "Confirmar instrumental 24 h antes", "No vacio", "Sirve para seguimiento"],
];
styleTable(programa, "A13:F20");
programa.getRange("A:J").format.autofitColumns();
programa.freezePanes.freezeRows(3);

for (const s of [dashboard, data, baseDiaria, calc, cats, informe, mejora, programa]) {
  s.getRange("A1:Z200").format.font.name = "Calibri";
}

if (typeof wb.recalculate === "function") wb.recalculate();

const key = await wb.inspect({
  kind: "table",
  range: "Dashboard CRR!A1:M22",
  include: "values,formulas",
  tableMaxRows: 24,
  tableMaxCols: 13,
});
console.log(key.ndjson);

const errors = await wb.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "final formula error scan",
});
console.log(errors.ndjson);

for (const [sheetName, range] of [
  ["Dashboard CRR", "A1:M50"],
  ["Datos Cirugias", "A1:P25"],
  ["Base Diaria CRR", "A1:P12"],
  ["Calculos", "A1:M22"],
  ["Catalogos", "A1:F22"],
  ["Informe Semana 21", "A1:H35"],
  ["Plan Mejora", "A1:J12"],
  ["Programa Diario", "A1:J20"],
]) {
  await wb.render({ sheetName, range, format: "png", scale: 1 });
}

await fs.mkdir(outputDir, { recursive: true });
const output = await SpreadsheetFile.exportXlsx(wb);
await output.save(outputPath);
console.log(outputPath);
