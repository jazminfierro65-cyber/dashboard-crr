const EXCEL_URL = "./outputs/crr_dashboard/plantilla_carga_web_crr.xlsx";
const palette = ["#16639f", "#e8743b", "#177e55", "#7c3aed", "#0f766e", "#b91c1c", "#64748b"];

let records = [];
let kpiReferences = [];

const filters = {
  month: document.querySelector("#monthFilter"),
  week: document.querySelector("#weekFilter"),
  date: document.querySelector("#dateFilter"),
  specialty: document.querySelector("#specialtyFilter"),
  subspecialty: document.querySelector("#subspecialtyFilter"),
  room: document.querySelector("#roomFilter"),
  reload: document.querySelector("#reloadExcel"),
  clear: document.querySelector("#clearFilters"),
};

function clean(value) {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

function normalizeRoom(value) {
  const text = clean(value);
  const match = text.match(/quirofano\s*0*(\d+)/i);
  if (!match) return text;
  return `Quirofano ${String(Number(match[1])).padStart(2, "0")}`;
}

function normalizeDate(value) {
  if (!value) return "";
  if (typeof value === "number") {
    const date = new Date(Date.UTC(1899, 11, 30 + value));
    return date.toISOString().slice(0, 10);
  }
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  const text = clean(value);
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;
  const match = text.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/);
  if (match) return `${match[3]}-${match[2].padStart(2, "0")}-${match[1].padStart(2, "0")}`;
  return text;
}

function normalizeTime(value) {
  if (value === null || value === undefined || value === "") return "";
  if (typeof value === "number") {
    const total = Math.round(value * 24 * 60);
    return `${String(Math.floor(total / 60) % 24).padStart(2, "0")}:${String(total % 60).padStart(2, "0")}`;
  }
  if (value instanceof Date) return `${String(value.getHours()).padStart(2, "0")}:${String(value.getMinutes()).padStart(2, "0")}`;
  return clean(value);
}

function numericValue(value) {
  if (value === null || value === undefined || value === "") return "";
  const number = Number(value);
  return Number.isFinite(number) ? number : "";
}

function timeToMinutes(value) {
  const text = normalizeTime(value);
  const match = text.match(/^(\d{1,2}):(\d{2})/);
  if (!match) return null;
  return Number(match[1]) * 60 + Number(match[2]);
}

function minutesBetween(start, end) {
  const startMinutes = timeToMinutes(start);
  const endMinutes = timeToMinutes(end);
  if (startMinutes === null || endMinutes === null) return "";
  const difference = endMinutes - startMinutes;
  return difference >= 0 ? difference : difference + 1440;
}

function normalizeCode(value) {
  const text = clean(value);
  if (!text) return "";
  const base = text.split("-")[0].replace(/\D/g, "");
  return base || text.replace(/\D/g, "");
}

function normalizeCategory(value) {
  const text = clean(value);
  if (!text || text.startsWith("=")) return "";
  if (/may/i.test(text)) return "Cirugía Mayor";
  if (/men/i.test(text)) return "Cirugía Menor";
  if (/proc/i.test(text)) return "Procedimiento";
  return text;
}

function buildCategoryMap(workbook) {
  const sheet = workbook.Sheets["Catalogo MINSAL"];
  if (!sheet) return new Map();
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "", range: 2 });
  return new Map(
    rows
      .map((row) => [normalizeCode(row["Codigo"]), normalizeCategory(row["Categoria"])])
      .filter(([code, category]) => code && category),
  );
}

function categoryFromCodes(code1, code2, categoryMap) {
  return categoryMap.get(normalizeCode(code1)) || categoryMap.get(normalizeCode(code2)) || "";
}

function buildScheduleMap(workbook) {
  const sheet = workbook.Sheets["Horario Pabellones"];
  if (!sheet) return new Map();
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "", range: 2 });
  return new Map(
    rows
      .map((row) => {
        const room = normalizeRoom(row["Quirofano"]);
        const start = normalizeTime(row["Inicio funcionamiento"]);
        const end = normalizeTime(row["Termino funcionamiento"]);
        const minutesFromSheet = numericValue(row["Minutos disponibles"]);
        const minutes = minutesFromSheet === "" ? minutesBetween(start, end) : minutesFromSheet;
        return [room, { start, end, minutes }];
      })
      .filter(([room, value]) => room && Number.isFinite(Number(value.minutes)) && Number(value.minutes) > 0),
  );
}

function sameDayMinutesBetween(previousEnd, nextStart) {
  const previousMinutes = timeToMinutes(previousEnd);
  const nextMinutes = timeToMinutes(nextStart);
  if (previousMinutes === null || nextMinutes === null || nextMinutes < previousMinutes) return "";
  return nextMinutes - previousMinutes;
}

function classifyAccuracy(difference) {
  if (difference === "" || !Number.isFinite(Number(difference))) return "";
  if (Number(difference) > 15) return "Subestimacion";
  if (Number(difference) < -15) return "Sobrestimacion";
  return "Asertivo";
}

function classifyEntryTime(value) {
  const minutes = timeToMinutes(value);
  if (minutes === null) return "";
  return minutes >= 8 * 60 && minutes <= 8 * 60 + 15 ? "Ingreso adecuado primera hora" : "Atraso primera hora";
}

function classifyTurnover(value) {
  if (value === "" || !Number.isFinite(Number(value))) return "";
  return Number(value) <= 15 ? "Cumple" : "No cumple";
}

function formatDate(value) {
  if (!value || value === "todos") return "Todos";
  const [year, month, day] = value.split("-");
  return `${day}-${month}-${year}`;
}

function monthKey(value) {
  return value ? value.slice(0, 7) : "";
}

function formatMonth(value) {
  if (!value || value === "todos") return "Todos los meses";
  const [year, month] = value.split("-");
  const date = new Date(Number(year), Number(month) - 1, 1);
  return new Intl.DateTimeFormat("es-CL", { month: "long", year: "numeric" }).format(date);
}

function parseDate(value) {
  if (!value) return null;
  const [year, month, day] = value.split("-").map(Number);
  return new Date(year, month - 1, day);
}

function isoDate(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function weekStart(value) {
  const date = parseDate(value);
  if (!date) return "";
  const day = date.getDay() || 7;
  date.setDate(date.getDate() - day + 1);
  return isoDate(date);
}

function recordWeek(item) {
  return item.weekStart || weekStart(item.date);
}

function weekRangeLabel(startValue) {
  const start = parseDate(startValue);
  if (!start) return "Sin semana";
  const end = new Date(start);
  end.setDate(start.getDate() + 6);
  return `${formatDate(isoDate(start))} al ${formatDate(isoDate(end))}`;
}

function percent(value) {
  return `${(value * 100).toLocaleString("es-CL", { maximumFractionDigits: 1 })}%`;
}

function currentTimestamp() {
  return new Intl.DateTimeFormat("es-CL", {
    dateStyle: "short",
    timeStyle: "short",
  }).format(new Date());
}

function isDone(item) {
  return /atendido|realizado/i.test(item.status);
}

function isSuspended(item) {
  return /suspend/i.test(item.status);
}

function suspensionDetail(item) {
  if (item.suspensionDetail) return item.suspensionDetail;
  if (item.cause && item.suspensionJustification) return `${item.cause}: ${item.suspensionJustification}`;
  return item.suspensionJustification || item.cause || "Sin causa/motivo registrado";
}

function suspensionCause(item) {
  if (item.cause) return item.cause;
  if (item.suspensionJustification) return `Sin causa registrada (motivo: ${item.suspensionJustification})`;
  return "Sin causa registrada";
}

function isLate(item) {
  return /atraso/i.test(item.entryClass);
}

function isAccurate(item) {
  return /asertivo/i.test(item.accuracy);
}

function isYes(value) {
  return /^s[ií]$/i.test(clean(value));
}

function tally(items, keyFn) {
  return items.reduce((acc, item) => {
    const key = keyFn(item) || "Sin dato";
    acc[key] = (acc[key] || 0) + 1;
    return acc;
  }, {});
}

function topLabel(items, keyFn, fallback = "Sin dato") {
  const entries = Object.entries(tally(items, keyFn))
    .filter(([label]) => label && !/^sin\s/i.test(label))
    .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]));
  if (!entries.length) return fallback;
  const [label, count] = entries[0];
  return `${label} (${count})`;
}

async function loadExcel() {
  if (!window.XLSX) {
    alert("No se pudo cargar el lector de Excel.");
    return;
  }
  filters.reload.textContent = "Actualizando...";
  const response = await fetch(`${EXCEL_URL}?t=${Date.now()}`);
  const workbook = XLSX.read(await response.arrayBuffer(), { type: "array", cellDates: true });
  const sheet = workbook.Sheets["Carga Web Cirugias"] || workbook.Sheets[workbook.SheetNames[0]];
  const referenceSheet = workbook.Sheets["Referencias KPI"];
  const scheduleMap = buildScheduleMap(workbook);
  const categoryMap = buildCategoryMap(workbook);
  kpiReferences = referenceSheet ? XLSX.utils.sheet_to_json(referenceSheet, { defval: "" }) : [];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "", range: 2 });
  records = rows
    .filter((row) => row["Fecha Quirofano"] || row["Paciente"] || row["Episodio"])
    .map((row, index) => {
      const first = clean(row["Nombre"]) || clean(row["Paciente"]);
      const last = clean(row["Apellido"]);
      const surgeryStart = normalizeTime(row["Inicio Cirugia"]);
      const surgeryEnd = normalizeTime(row["Termino Cirugia"]);
      const durationFromExcel = numericValue(row["Duracion Cirugia Formula Min"]);
      const programmed = numericValue(row["Tiempo programado de cirugía en tabla (Min)"]);
      const duration = durationFromExcel === "" ? minutesBetween(surgeryStart, surgeryEnd) : durationFromExcel;
      const diffFromExcel = numericValue(row["Diferencia Real - Programado Min"]);
      const diff = diffFromExcel === "" && duration !== "" && programmed !== "" ? duration - programmed : diffFromExcel;
      const accuracy = clean(row["Asertividad"]) || classifyAccuracy(diff);
      const entryTime = normalizeTime(row["Ingreso Pabellon"]);
      const exitTime = normalizeTime(row["Salida Pabellon"]);
      const turnover = numericValue(row["Tiempo recambio entre cirugías (Min)"]);
      const turnoverCompliance = clean(row["Cumple recambio <=15 min"]) || classifyTurnover(turnover);
      const firstHourJustification = clean(row["Justificación Atraso Primera Hora"]);
      const turnoverJustification = clean(row["Justificación Recambio >15 Min"]);
      const suspensionJustification = clean(row["Motivo de Suspensión"]) || clean(row["Motivo de Suspension"]) || clean(row["Motivo Justificado"]);
      const suspensionDetailFromExcel = clean(row["Detalle Suspensión Dashboard"]);
      const suspensionDetailText = suspensionDetailFromExcel.startsWith("=") ? "" : suspensionDetailFromExcel;
      const category = normalizeCategory(row["Categoria"]) || categoryFromCodes(row["Codigo 1"], row["Codigo 2"], categoryMap);
      const cma = clean(row["Ingreso CMA ambulatorio"]);
      const hospitalized = clean(row["Hospitalizado posterior"]);
      const room = normalizeRoom(row["Quirofano"]) || normalizeRoom(row["Pabellon/Quirofano"]) || "Sin pabellón";
      const schedule = scheduleMap.get(room);
      const workingFromExcel = numericValue(row["Horario Funcionamiento Pabellon (Min)"]);
      const workingMinutes = workingFromExcel === "" && schedule ? schedule.minutes : workingFromExcel;
      const date = normalizeDate(row["Fecha Quirofano"]);
      return {
        id: clean(row["ID Registro Dashboard"]) || clean(row["Episodio"]) || `row-${index}`,
        date,
        weekStart: normalizeDate(row["Semana Inicio"]) || weekStart(date),
        dashboardValidation: clean(row["Validación Carga Dashboard"]),
        patient: [first, last].filter(Boolean).join(" "),
        room,
        status: clean(row["Estado"]) || "Sin estado",
        entryTime,
        exitTime,
        entryClass: clean(row["Clasificación ingreso 08:00-08:15"]),
        category: category || "Sin categoría",
        specialty: clean(row["Especialidad"]) || "Sin especialidad",
        subspecialty: clean(row["Subespecialidad"]) || "Sin subespecialidad",
        surgery: clean(row["Descripción Intervención"]) || clean(row["Cirugia"]),
        cause: clean(row["Causas Suspension"]) || clean(row["Causa Suspension"]),
        suspensionDetail: suspensionDetailText,
        firstHourJustification,
        turnoverJustification,
        suspensionJustification,
        surgeryStart,
        surgeryEnd,
        duration,
        programmed,
        workingMinutes,
        diff,
        accuracy,
        turnover,
        turnoverCompliance,
        cma,
        hospitalized,
      };
    });
  applyFirstHourEntryClass(records);
  applyTurnover(records);
  initFilters();
  render();
  filters.reload.textContent = "Actualizar Excel";
}

function applyFirstHourEntryClass(items) {
  const groups = new Map();
  items.forEach((item) => {
    const minutes = timeToMinutes(item.entryTime);
    if (minutes === null) {
      item.entryClass = "";
      return;
    }
    const key = `${item.date}||${item.room}`;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push({ item, minutes });
  });
  groups.forEach((group) => {
    group.sort((a, b) => a.minutes - b.minutes);
    group.forEach(({ item, minutes }, index) => {
      item.entryClass = index === 0 ? classifyEntryTime(item.entryTime) : "No aplica";
      if (index === 0 && minutes < 8 * 60) item.entryClass = "Ingreso adecuado primera hora";
    });
  });
}

function applyTurnover(items) {
  const groups = new Map();
  items.forEach((item) => {
    if (!item.date || !item.room || !item.entryTime) return;
    const key = `${item.date}||${item.room}`;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(item);
  });
  groups.forEach((group) => {
    group.sort((a, b) => (timeToMinutes(a.entryTime) ?? 0) - (timeToMinutes(b.entryTime) ?? 0));
    group.forEach((item, index) => {
      if (item.turnover !== "") return;
      const previous = group
        .slice(0, index)
        .filter((candidate) => candidate.exitTime && timeToMinutes(candidate.exitTime) < timeToMinutes(item.entryTime))
        .sort((a, b) => (timeToMinutes(b.exitTime) ?? 0) - (timeToMinutes(a.exitTime) ?? 0))[0];
      item.turnover = previous ? sameDayMinutesBetween(previous.exitTime, item.entryTime) : "";
      item.turnoverCompliance = classifyTurnover(item.turnover);
    });
  });
}

function initFilters() {
  const currentMonth = filters.month.value || "latest";
  const currentWeek = filters.week.value || "latest";
  const currentDate = filters.date.value || "latest";
  const currentSpecialty = filters.specialty.value || "todos";
  const currentSubspecialty = filters.subspecialty.value || "todos";
  const currentRoom = filters.room.value || "todos";
  const dates = [...new Set(records.map((item) => item.date).filter(Boolean))].sort();
  const months = [...new Set(dates.map(monthKey).filter(Boolean))].sort();
  const weeks = [...new Set(records.map(recordWeek).filter(Boolean))].sort();
  const specialties = [...new Set(records.map((item) => item.specialty))].sort();
  const subspecialties = [...new Set(records.map((item) => item.subspecialty))].sort();
  const rooms = [...new Set(records.map((item) => item.room))].sort();
  filters.month.innerHTML = `<option value="todos">Todos los meses</option>${months.map((month) => `<option value="${month}">${formatMonth(month)}</option>`).join("")}`;
  filters.week.innerHTML = `<option value="todos">Todas las semanas</option>${weeks.map((week) => `<option value="${week}">${weekRangeLabel(week)}</option>`).join("")}`;
  filters.date.innerHTML = `<option value="todos">Todos los días</option>${dates.map((date) => `<option value="${date}">${formatDate(date)}</option>`).join("")}`;
  filters.specialty.innerHTML = `<option value="todos">Todas</option>${specialties.map((specialty) => `<option>${specialty}</option>`).join("")}`;
  filters.subspecialty.innerHTML = `<option value="todos">Todas</option>${subspecialties.map((subspecialty) => `<option>${subspecialty}</option>`).join("")}`;
  filters.room.innerHTML = `<option value="todos">Todos</option>${rooms.map((room) => `<option>${room}</option>`).join("")}`;
  filters.month.value = currentMonth === "latest" ? monthKey(dates.at(-1)) || "todos" : months.includes(currentMonth) ? currentMonth : "todos";
  filters.week.value = currentWeek === "latest" ? weekStart(dates.at(-1)) || "todos" : weeks.includes(currentWeek) ? currentWeek : "todos";
  filters.date.value = currentDate === "latest" ? "todos" : dates.includes(currentDate) ? currentDate : "todos";
  filters.specialty.value = specialties.includes(currentSpecialty) ? currentSpecialty : "todos";
  filters.subspecialty.value = subspecialties.includes(currentSubspecialty) ? currentSubspecialty : "todos";
  filters.room.value = rooms.includes(currentRoom) ? currentRoom : "todos";
}

function filteredRecords() {
  return records.filter((item) => {
    const dateOk =
      filters.date.value !== "todos"
        ? item.date === filters.date.value
        : filters.week.value !== "todos"
          ? recordWeek(item) === filters.week.value
        : filters.month.value === "todos" || monthKey(item.date) === filters.month.value;
    const specialtyOk = filters.specialty.value === "todos" || item.specialty === filters.specialty.value;
    const subspecialtyOk = filters.subspecialty.value === "todos" || item.subspecialty === filters.subspecialty.value;
    const roomOk = filters.room.value === "todos" || item.room === filters.room.value;
    return dateOk && specialtyOk && subspecialtyOk && roomOk;
  });
}

function trendRecords() {
  return records.filter((item) => {
    const periodOk =
      filters.date.value !== "todos"
        ? item.date === filters.date.value
        : filters.week.value !== "todos"
          ? recordWeek(item) === filters.week.value
          : filters.month.value === "todos" || monthKey(item.date) === filters.month.value;
    const specialtyOk = filters.specialty.value === "todos" || item.specialty === filters.specialty.value;
    const subspecialtyOk = filters.subspecialty.value === "todos" || item.subspecialty === filters.subspecialty.value;
    const roomOk = filters.room.value === "todos" || item.room === filters.room.value;
    return periodOk && specialtyOk && subspecialtyOk && roomOk;
  });
}

function render() {
  const data = filteredRecords();
  const done = data.filter(isDone);
  const suspended = data.filter(isSuspended);
  const late = data.filter(isLate);
  const accurate = data.filter(isAccurate);
  const turnoverValues = data
    .filter((item) => item.turnover !== "")
    .map((item) => Number(item.turnover))
    .filter((value) => Number.isFinite(value) && value >= 0);
  const avgTurnover = average(turnoverValues);
  const turnoverCompliant = data.filter((item) => item.turnover !== "" && Number(item.turnover) <= 15);
  updateSummary(data, done, late, suspended, accurate);
  renderKpiSet("week", data, done, late, suspended, accurate, turnoverValues, avgTurnover, turnoverCompliant);
  renderStatus(data, done, late, suspended, accurate);
  renderOperationalMotives(data, late);
  renderDelayJustificationSummary(data);
  renderRanking(trendRecords());
  renderDailyRoomProduction(data);
  renderAttentionRows(data);
  renderTurnover(data);
  renderCauses(suspended);
  renderCmaFlow(data);
  renderTrend(trendRecords());
  renderWeeklySummary(trendRecords());
  renderWeeklyReasonBars(trendRecords());
  renderReferenceBenchmark(trendRecords());
  renderOccupancyInfo(trendRecords());
  renderDonut("categoryChart", "categoryLegend", tally(data, (item) => item.category));
  renderDonut("specialtyChart", "specialtyLegend", tally(data, (item) => item.specialty));
}

function setText(id, value) {
  const element = document.querySelector(`#${id}`);
  if (element) element.textContent = value;
}

function renderKpiSet(prefix, data, done, late, suspended, accurate, turnoverValues, avgTurnover, turnoverCompliant) {
  const id = (base) => (prefix ? `${prefix}Kpi${base}` : `kpi${base}`);
  setText(id("Total"), data.length);
  const periodText =
    filters.date.value !== "todos"
      ? `Día ${formatDate(filters.date.value)}`
      : filters.week.value !== "todos"
        ? `Semana del ${weekRangeLabel(filters.week.value)}`
        : `Mes ${formatMonth(filters.month.value)}`;
  setText(`${id("Total")}Text`, periodText);
  setText(`${id("Total")}Done`, done.length);
  setText(id("Done"), data.length ? percent(done.length / data.length) : "0%");
  setText(`${id("Done")}Text`, `${done.length} realizados / atendidos`);
  setText(id("Late"), data.length ? percent(late.length / data.length) : "0%");
  setText(`${id("Late")}Text`, `${late.length} ingresos con atraso`);
  setText(id("Susp"), data.length ? percent(suspended.length / data.length) : "0%");
  setText(`${id("Susp")}Text`, `${suspended.length} casos suspendidos`);
  setText(id("Accurate"), data.length ? percent(accurate.length / data.length) : "0%");
  setText(id("Turnover"), turnoverValues.length ? `${avgTurnover.toLocaleString("es-CL", { maximumFractionDigits: 1 })} min` : "0 min");
  setText(
    `${id("Turnover")}Text`,
    turnoverValues.length ? `${turnoverCompliant.length}/${turnoverValues.length} cumplen <=15 min` : "0 recambios calculados",
  );
}

function updateSummary(data, done, late, suspended, accurate) {
  const dateText =
    filters.date.value !== "todos"
      ? formatDate(filters.date.value)
      : filters.week.value !== "todos"
        ? weekRangeLabel(filters.week.value)
        : formatMonth(filters.month.value);
  const specialtyText = filters.specialty.value === "todos" ? "todas las especialidades" : filters.specialty.value;
  const subspecialtyText = filters.subspecialty.value === "todos" ? "todas las subespecialidades" : filters.subspecialty.value;
  const roomText = filters.room.value === "todos" ? "todos los pabellones" : filters.room.value;
  const priority = suspended.length
    ? "Revisar causas de suspensión"
    : late.length
      ? "Revisar ingreso de primera hora"
      : data.length && accurate.length / data.length < 0.65
        ? "Revisar programación de tiempos"
        : "Mantener seguimiento";
  const filterSummary = document.querySelector("#filterSummary");
  if (filterSummary) filterSummary.textContent = `${dateText} · ${specialtyText} · ${subspecialtyText} · ${roomText}`;
  document.querySelector("#summaryDate").textContent = dateText;
  document.querySelector("#lastUpdated").textContent = currentTimestamp();
  document.querySelector("#mainPriority").textContent = priority;
}

function renderStatus(data, done, late, suspended, accurate) {
  const box = document.querySelector("#statusCallout");
  const title = document.querySelector("#statusTitle");
  const text = document.querySelector("#statusText");
  box.classList.remove("warning", "danger");
  if (!data.length) {
    title.textContent = "Sin producción para el filtro";
    text.textContent = "Cambie los filtros o actualice el Excel.";
  } else if (suspended.length || late.length / data.length > 0.35) {
    box.classList.add("danger");
    title.textContent = "Producción con alerta";
    text.textContent = `${late.length} atrasos y ${suspended.length} suspensiones requieren revisión de jefatura.`;
  } else if (late.length || accurate.length / data.length < 0.65) {
    box.classList.add("warning");
    title.textContent = "Producción en observación";
    text.textContent = "Conviene revisar ingresos de primera hora y programación de tiempos quirúrgicos.";
  } else {
    title.textContent = "Producción dentro de control";
    text.textContent = `${done.length} casos resueltos y sin alerta mayor dominante.`;
  }
  const topRoom = Object.entries(tally(late, (item) => item.room)).sort((a, b) => b[1] - a[1])[0];
  const topSpecialty = Object.entries(tally(data, (item) => item.specialty)).sort((a, b) => b[1] - a[1])[0];
  const topCause = Object.entries(tally(suspended, suspensionCause)).sort((a, b) => b[1] - a[1])[0];
  document.querySelector("#insightList").innerHTML = [
    ["Mayor volumen", topSpecialty ? `${topSpecialty[0]} concentra ${topSpecialty[1]} caso(s).` : "Sin especialidad dominante."],
    ["Pabellón con atraso", topRoom ? `${topRoom[0]} concentra ${topRoom[1]} atraso(s).` : "Sin atrasos registrados."],
    ["Suspensión principal", topCause ? `${topCause[0]}: ${topCause[1]} caso(s).` : "Sin suspensiones dominantes."],
  ]
    .map(([head, detail]) => `<div class="insight"><strong>${head}</strong><span>${detail}</span></div>`)
    .join("");
  document.querySelector("#actionList").innerHTML = [
    ["Acción sugerida", suspended.length ? "Validar causa de suspensión y responsable de cierre." : "Mantener control de producción diaria."],
    ["Ingreso a pabellón", late.length ? "Priorizar revisión de atrasos de primera hora por pabellón." : "Ingreso de primera hora sin alerta dominante."],
    ["Programación", accurate.length / Math.max(data.length, 1) < 0.65 ? "Ajustar tiempos programados según duración real por cirugía." : "Asertividad quirúrgica dentro de lectura favorable."],
  ]
    .map(([head, detail]) => `<div class="action"><strong>${head}</strong><span>${detail}</span></div>`)
    .join("");
  renderCauseBreakdown(data, late, suspended);
}

function average(values) {
  const nums = values.map(Number).filter(Number.isFinite);
  return nums.length ? nums.reduce((sum, value) => sum + value, 0) / nums.length : 0;
}

function occupancyStats(items) {
  const roomDays = new Map();
  items.forEach((item) => {
    if (!item.date || !item.room || item.room === "Sin pabellón") return;
    const key = `${item.date}||${item.room}`;
    const working = Number(item.workingMinutes);
    if (!roomDays.has(key)) {
      roomDays.set(key, { date: item.date, room: item.room, workingMinutes: working, starts: [], ends: [] });
    } else if (Number.isFinite(working) && working > 0) {
      roomDays.get(key).workingMinutes = working;
    }
    if (isDone(item)) {
      const start = timeToMinutes(item.surgeryStart);
      const end = timeToMinutes(item.surgeryEnd);
      if (start !== null && end !== null) {
        roomDays.get(key).starts.push(start);
        roomDays.get(key).ends.push(end);
      }
    }
  });
  const calculated = [...roomDays.values()].filter(
    (item) => Number.isFinite(Number(item.workingMinutes)) && Number(item.workingMinutes) > 0,
  );
  calculated.forEach((item) => {
    item.usedMinutes = item.starts.length && item.ends.length ? Math.max(...item.ends) - Math.min(...item.starts) : 0;
    if (item.usedMinutes < 0) item.usedMinutes += 1440;
  });
  const usedMinutes = calculated.reduce((sum, item) => sum + item.usedMinutes, 0);
  const availableMinutes = calculated.reduce((sum, item) => sum + Number(item.workingMinutes), 0);
  const dailyRates = calculated.map((item) => (item.usedMinutes / 60) / (Number(item.workingMinutes) / 60));
  const missing = [...roomDays.values()].filter(
    (item) => !Number.isFinite(Number(item.workingMinutes)) || Number(item.workingMinutes) <= 0,
  );
  return {
    usedMinutes,
    availableMinutes,
    missing,
    roomDayCount: roomDays.size,
    calculatedDayCount: calculated.length,
    rate: dailyRates.length ? average(dailyRates) : null,
  };
}

function referenceFor(indicator) {
  return kpiReferences.find((row) => clean(row.Indicador).toLowerCase() === indicator.toLowerCase()) || {};
}

function referenceText(row) {
  const value = numericValue(row.Referencia);
  const unit = clean(row.Unidad);
  const rule = clean(row.Regla);
  const number = value === "" ? clean(row.Referencia) : value.toLocaleString("es-CL", { maximumFractionDigits: 1 });
  if (rule === "menor") return `< ${number}${unit ? ` ${unit}` : ""}`;
  if (rule === "mayor") return `> ${number}${unit ? ` ${unit}` : ""}`;
  if (rule === "mayor_igual") return `>= ${number}${unit ? ` ${unit}` : ""}`;
  return `${number}${unit ? ` ${unit}` : ""}`;
}

function evaluateReference(latest, row) {
  const reference = numericValue(row.Referencia);
  if (reference === "" || latest === "" || !Number.isFinite(Number(latest))) return "neutral";
  const value = Number(latest);
  const rule = clean(row.Regla);
  if (rule === "menor") return value < reference ? "green" : "danger";
  if (rule === "mayor") return value > reference ? "green" : "danger";
  if (rule === "mayor_igual") return value >= reference ? "green" : "danger";
  if (rule === "referencia") return value >= reference ? "green" : "warning";
  return "neutral";
}

function latestWeekItems(items) {
  const dated = items.filter((item) => item.date && recordWeek(item));
  const latest = [...new Set(dated.map(recordWeek))].sort().at(-1);
  return latest ? dated.filter((item) => recordWeek(item) === latest) : [];
}

function renderReferenceBenchmark(items) {
  const box = document.querySelector("#benchmarkRows");
  const title = document.querySelector("#benchmarkWeek");
  if (!box) return;
  const weekItems = latestWeekItems(items);
  if (!weekItems.length) {
    title.textContent = "Sin semana disponible";
    box.innerHTML = `<div class="empty-note">Sin datos suficientes para calcular la última semana.</div>`;
    return;
  }
  const start = weekStart(weekItems[0].date);
  const done = weekItems.filter(isDone);
  const late = weekItems.filter(isLate);
  const suspended = weekItems.filter(isSuspended);
  const firstHourDelayValues = late
    .map((item) => {
      const minutes = timeToMinutes(item.entryTime);
      return minutes === null ? null : minutes - (8 * 60 + 15);
    })
    .filter((value) => Number.isFinite(value) && value > 0);
  const turnoverValues = weekItems
    .filter((item) => item.turnover !== "")
    .map((item) => Number(item.turnover))
    .filter((value) => Number.isFinite(value) && value >= 0);
  const positiveDiffs = weekItems
    .map((item) => Number(item.diff))
    .filter((value) => Number.isFinite(value) && value > 0);
  const roomDayCount =
    new Set(weekItems.filter((item) => item.date && item.room).map((item) => `${item.date}||${item.room}`)).size || 1;
  const occupancy = occupancyStats(weekItems);
  const firstHourDelayAvg = firstHourDelayValues.length ? average(firstHourDelayValues) : 0;
  const suspensionRate = weekItems.length ? (suspended.length / weekItems.length) * 100 : 0;
  const turnoverAvg = turnoverValues.length ? average(turnoverValues) : "";
  const subprogrammingAvg = positiveDiffs.length ? average(positiveDiffs) : 0;
  const dailyYield = done.length / roomDayCount;

  title.textContent = weekRangeLabel(start);
  const rows = [
    {
      referenceLabel: "Pacientes Programados",
      latestLabel: "Pacientes Operados",
      ref: referenceFor("Pacientes Programados"),
      latest: done.length,
      latestText: done.length.toLocaleString("es-CL"),
      latestDetail: `${weekItems.length} programados`,
    },
    {
      referenceLabel: "Retraso Primera Hora",
      latestLabel: "Retraso 1° Hora",
      ref: referenceFor("Retraso Primera Hora"),
      latest: firstHourDelayAvg,
      latestText: `${firstHourDelayAvg.toLocaleString("es-CL", { maximumFractionDigits: 1 })} min`,
      latestDetail: `${firstHourDelayValues.length} atraso${firstHourDelayValues.length === 1 ? "" : "s"} después de 08:15`,
    },
    {
      referenceLabel: "Ocupación Diaria",
      latestLabel: "Promedio Ocupación Diaria",
      ref: referenceFor("Ocupación Diaria"),
      latest: occupancy.rate === null ? "" : occupancy.rate * 100,
      latestText: occupancy.rate === null ? "S/D" : percent(occupancy.rate),
      latestDetail: occupancy.rate !== null
        ? `Mismo cálculo Excel · promedio ${occupancy.calculatedDayCount} pabellón-día`
        : "Falta horario funcionamiento",
    },
    {
      referenceLabel: "Rendimiento Diario",
      latestLabel: "Rendimiento Diario",
      ref: referenceFor("Rendimiento Diario"),
      latest: dailyYield,
      latestText: dailyYield.toLocaleString("es-CL", { maximumFractionDigits: 2 }),
      latestDetail: `${done.length} operados / ${roomDayCount} pabellón-día`,
    },
    {
      referenceLabel: "Promedio Recambio",
      latestLabel: "Promedio Recambio",
      ref: referenceFor("Promedio Recambio"),
      latest: turnoverAvg,
      latestText: turnoverValues.length ? turnoverAvg.toLocaleString("es-CL", { maximumFractionDigits: 1 }) : "S/D",
      latestDetail: turnoverValues.length ? "min" : "sin recambios",
    },
    {
      referenceLabel: "Subprogramación Diaria",
      latestLabel: "Subprogramación Diaria",
      ref: referenceFor("Subprogramación Diaria"),
      latest: subprogrammingAvg,
      latestText: subprogrammingAvg.toLocaleString("es-CL", { maximumFractionDigits: 1 }),
      latestDetail: "min promedio",
    },
    {
      referenceLabel: "Suspensiones",
      latestLabel: "Suspensiones",
      ref: referenceFor("Suspensiones"),
      latest: suspensionRate,
      latestText: percent(suspensionRate / 100),
      latestDetail: `${suspended.length} caso${suspended.length === 1 ? "" : "s"}`,
    },
  ];
  box.innerHTML = rows
    .map((row) => {
      const status = evaluateReference(row.latest, row.ref);
      return `<article class="benchmark-row ${status}">
        <section>
          <span>${row.referenceLabel}</span>
          <strong>${referenceText(row.ref)}</strong>
        </section>
        <section>
          <span>${row.latestLabel}</span>
          <strong>${row.latestText}</strong>
          <small>${row.latestDetail}</small>
        </section>
      </article>`;
    })
    .join("");
}

function renderOccupancyInfo(items) {
  const box = document.querySelector("#occupancyInfo");
  if (!box) return;
  const weekItems = latestWeekItems(items);
  if (!weekItems.length) {
    box.innerHTML = `<article class="occupancy-card full"><strong>Sin datos semanales</strong><span>No hay registros con fecha para calcular ocupación diaria.</span></article>`;
    return;
  }

  const occupancy = occupancyStats(weekItems);
  const status = occupancy.rate === null ? "Pendiente" : occupancy.rate >= 0.75 ? "Sobre referencia" : "Bajo referencia";
  const resultDetail = occupancy.rate === null
    ? `${occupancy.usedMinutes.toLocaleString("es-CL", { maximumFractionDigits: 0 })} min usados. Falta cargar inicio y término de funcionamiento.`
    : `Promedio de ${occupancy.calculatedDayCount} pabellón-día · ${occupancy.usedMinutes.toLocaleString("es-CL", { maximumFractionDigits: 0 })} min usados / ${occupancy.availableMinutes.toLocaleString("es-CL", { maximumFractionDigits: 0 })} min programados`;
  const missingText = occupancy.missing.length
    ? `${occupancy.missing.length} pabellón-día sin horario de funcionamiento`
    : "Horario disponible completo para la semana";

  box.innerHTML = `
    <article class="occupancy-card formula">
      <span>Definición</span>
      <strong>Horas ocupadas / horas programadas</strong>
      <small>Mismo cálculo del Excel: (horas totales ocupadas / horas totales programadas) x 100.</small>
    </article>
    <article class="occupancy-card result ${occupancy.rate === null ? "warning" : occupancy.rate >= 0.75 ? "green" : "danger"}">
      <span>Última semana</span>
      <strong>${occupancy.rate === null ? "S/D" : percent(occupancy.rate)}</strong>
      <small>${resultDetail}</small>
    </article>
    <article class="occupancy-card action">
      <span>Valor semanal</span>
      <strong>${status}</strong>
      <small>${missingText}</small>
    </article>`;
}

function breakdownRows(title, entries, emptyText) {
  const rows = entries.length
    ? entries
        .slice(0, 5)
        .map(([label, value]) => `<li><span>${label}</span><strong>${value}</strong></li>`)
        .join("")
    : `<li class="muted-row">${emptyText}</li>`;
  return `<section>
    <h3>${title}</h3>
    <ul>${rows}</ul>
  </section>`;
}

function renderCauseBreakdown(data, late, suspended) {
  const accuracyOut = data.filter((item) => /subestimacion|sobrestimacion/i.test(item.accuracy));
  const longTurnover = data.filter((item) => Number(item.turnover) > 15);
  const causeEntries = Object.entries(tally(suspended, suspensionCause)).sort((a, b) => b[1] - a[1]);
  const lateEntries = Object.entries(tally(late, (item) => item.room)).sort((a, b) => b[1] - a[1]);
  const accuracyEntries = Object.entries(tally(accuracyOut, (item) => item.accuracy || "Pendiente")).sort((a, b) => b[1] - a[1]);
  const turnoverEntries = Object.entries(tally(longTurnover, (item) => item.room)).sort((a, b) => b[1] - a[1]);
  document.querySelector("#causeBreakdown").innerHTML = [
    breakdownRows("Causas de suspensión", causeEntries, "Sin suspensiones en el filtro."),
    breakdownRows("Atraso primera hora", lateEntries, "Sin atrasos de primera hora."),
    breakdownRows("Asertividad fuera de rango", accuracyEntries, "Sin subestimación o sobrestimación."),
    breakdownRows("Recambios sobre 15 min", turnoverEntries, "Sin recambios sobre meta."),
  ].join("");
}

function renderOperationalMotives(data, late) {
  const longTurnover = data.filter((item) => Number(item.turnover) > 15);
  const lateByRoom = Object.entries(tally(late, (item) => item.room)).sort((a, b) => b[1] - a[1]);
  const turnoverByRoom = Object.entries(tally(longTurnover, (item) => item.room)).sort((a, b) => b[1] - a[1]);
  const totalIssues = late.length + longTurnover.length;
  const rows = [
    {
      className: late.length ? "warning" : "green",
      title: "Retraso primera hora",
      value: late.length,
      detail: "ingresos fuera de 08:00 a 08:15",
      entries: lateByRoom,
      empty: "Sin atrasos de primera hora.",
    },
    {
      className: longTurnover.length ? "danger" : "green",
      title: "Recambio sobre 15 min",
      value: longTurnover.length,
      detail: "recambios fuera de meta",
      entries: turnoverByRoom,
      empty: "Sin recambios sobre 15 minutos.",
    },
  ];
  document.querySelector("#operationalMotives").innerHTML = [
    `<article class="ops-card summary ${totalIssues ? "warning" : "green"}">
      <span>Motivos detectados</span>
      <strong>${totalIssues}</strong>
      <small>${data.length} caso${data.length === 1 ? "" : "s"} en el filtro actual</small>
    </article>`,
    ...rows.map(
      (row) => `<article class="ops-card ${row.className}">
        <div class="ops-card-head">
          <div>
            <span>${row.title}</span>
            <strong>${row.value}</strong>
          </div>
          <b>${row.detail}</b>
        </div>
        <ul>
          ${
            row.entries.length
              ? row.entries
                  .slice(0, 5)
                  .map(([room, count]) => `<li><span>${room}</span><strong>${count}</strong></li>`)
                  .join("")
              : `<li class="muted-row">${row.empty}</li>`
          }
        </ul>
      </article>`,
    ),
  ].join("");
}

function renderTurnover(data) {
  const grouped = new Map();
  data.forEach((item) => {
    const value = Number(item.turnover);
    if (item.turnover === "" || !Number.isFinite(value) || value < 0) return;
    if (!grouped.has(item.room)) grouped.set(item.room, []);
    grouped.get(item.room).push(value);
  });
  const rows = [...grouped.entries()]
    .map(([room, values]) => ({ room, avg: average(values), count: values.length }))
    .sort((a, b) => b.avg - a.avg)
    .slice(0, 6);
  renderMetricBars("turnoverBars", rows, {
    label: (row) => row.room,
    value: (row) => `${row.avg.toLocaleString("es-CL", { maximumFractionDigits: 1 })} min`,
    detail: (row) => `${row.count} recambio${row.count === 1 ? "" : "s"}`,
    score: (row) => row.avg,
    color: "blue",
    empty: "Sin recambios calculados para este filtro.",
  });
}

function renderCauses(suspended) {
  const rows = Object.entries(tally(suspended, suspensionDetail))
    .map(([cause, count]) => ({ cause, count }))
    .sort((a, b) => b.count - a.count)
    .slice(0, 6);
  renderMetricBars("causeBars", rows, {
    label: (row) => row.cause,
    value: (row) => `${row.count}`,
    detail: () => "suspensión",
    score: (row) => row.count,
    color: "red",
    empty: "Sin suspensiones para este filtro.",
  });
}

function renderCmaFlow(data) {
  const withCma = data.filter((item) => isYes(item.cma)).length;
  const hospitalized = data.filter((item) => isYes(item.hospitalized)).length;
  const pendingCma = data.filter((item) => !clean(item.cma)).length;
  const pendingHosp = data.filter((item) => !clean(item.hospitalized)).length;
  const rows = [
    ["CMA ambulatorio", withCma, `${data.length ? percent(withCma / data.length) : "0%"} del filtro`],
    ["Quedan hospitalizados", hospitalized, `${data.length ? percent(hospitalized / data.length) : "0%"} del filtro`],
    ["CMA pendiente", pendingCma, "sin dato cargado"],
    ["Hospitalización pendiente", pendingHosp, "sin dato cargado"],
  ];
  document.querySelector("#cmaFlow").innerHTML = rows
    .map(
      ([label, value, detail], index) => `<article class="flow-card ${index > 1 ? "warning" : ""}">
        <span>${label}</span>
        <strong>${value}</strong>
        <small>${detail}</small>
      </article>`,
    )
    .join("");
}

function renderWeeklySummary(items) {
  const monthlyBox = document.querySelector("#monthlySummary");
  const grouped = new Map();
  const monthTurnovers = [];
  const monthly = {
    total: 0,
    done: 0,
    late: 0,
    suspended: 0,
    accurate: 0,
    topCause: {},
  };
  items.forEach((item) => {
    const key = recordWeek(item);
    if (!key) return;
    monthly.total += 1;
    if (isDone(item)) monthly.done += 1;
    if (isLate(item)) monthly.late += 1;
    if (isSuspended(item)) {
      monthly.suspended += 1;
      const cause = suspensionCause(item);
      monthly.topCause[cause] = (monthly.topCause[cause] || 0) + 1;
    }
    if (isAccurate(item)) monthly.accurate += 1;
    const monthTurnover = Number(item.turnover);
    if (item.turnover !== "" && Number.isFinite(monthTurnover)) monthTurnovers.push(monthTurnover);
    if (!grouped.has(key)) {
      grouped.set(key, {
        total: 0,
        done: 0,
        late: 0,
        suspended: 0,
        accurate: 0,
        turnovers: [],
        topCause: {},
        items: [],
      });
    }
    const row = grouped.get(key);
    row.items.push(item);
    row.total += 1;
    if (isDone(item)) row.done += 1;
    if (isLate(item)) row.late += 1;
    if (isSuspended(item)) {
      row.suspended += 1;
      const cause = suspensionCause(item);
      row.topCause[cause] = (row.topCause[cause] || 0) + 1;
    }
    if (isAccurate(item)) row.accurate += 1;
    const turnover = Number(item.turnover);
    if (item.turnover !== "" && Number.isFinite(turnover)) row.turnovers.push(turnover);
  });
  const rows = [...grouped.entries()].sort((a, b) => a[0].localeCompare(b[0]));
  const box = document.querySelector("#weeklySummary");
  if (!rows.length) {
    monthlyBox.innerHTML = "";
    box.innerHTML = `<div class="empty-note">Sin semanas disponibles para el filtro actual.</div>`;
    return;
  }
  const monthDoneRate = monthly.total ? monthly.done / monthly.total : 0;
  const monthLateRate = monthly.total ? monthly.late / monthly.total : 0;
  const monthAccurateRate = monthly.total ? monthly.accurate / monthly.total : 0;
  const monthTurnoverAvg = average(monthTurnovers);
  const monthOccupancy = occupancyStats(items);
  const mainCause = Object.entries(monthly.topCause).sort((a, b) => b[1] - a[1])[0];
  const selectedMonth = filters.month.value === "todos" ? "Periodo seleccionado" : formatMonth(filters.month.value);
  const monthlyStatusClass = monthly.suspended || monthLateRate > 0.2 ? "danger" : monthAccurateRate < 0.65 ? "warning" : "green";
  const monthlyStatusText = monthlyStatusClass === "danger" ? "Mes con alerta" : monthlyStatusClass === "warning" ? "Mes en observación" : "Mes en control";
  monthlyBox.innerHTML = `<article class="month-card ${monthlyStatusClass}">
    <div class="month-title">
      <div>
        <span>Lectura mensual</span>
        <strong>${selectedMonth}</strong>
      </div>
      <b>${monthlyStatusText}</b>
    </div>
    <div class="month-metrics">
      <span><strong>${monthly.total}</strong> casos</span>
      <span><strong>${percent(monthDoneRate)}</strong> resolución</span>
      <span><strong>${monthly.late}</strong> atrasos 1a hora</span>
      <span><strong>${monthly.suspended}</strong> suspensiones</span>
      <span><strong>${percent(monthAccurateRate)}</strong> asertividad</span>
      <span><strong>${monthOccupancy.rate === null ? "S/D" : percent(monthOccupancy.rate)}</strong> ocupación</span>
      <span><strong>${monthTurnovers.length ? `${monthTurnoverAvg.toLocaleString("es-CL", { maximumFractionDigits: 1 })} min` : "S/D"}</strong> recambio</span>
    </div>
    <p>${mainCause ? `Causa principal de suspensión: ${mainCause[0]} (${mainCause[1]})` : "Sin suspensión dominante en el mes."}</p>
  </article>`;
  box.innerHTML = rows
    .map(([start, row], index) => {
      const doneRate = row.total ? row.done / row.total : 0;
      const lateRate = row.total ? row.late / row.total : 0;
      const suspRate = row.total ? row.suspended / row.total : 0;
      const accurateRate = row.total ? row.accurate / row.total : 0;
      const turnoverAvg = average(row.turnovers);
      const weeklyOccupancy = occupancyStats(row.items);
      const cause = Object.entries(row.topCause).sort((a, b) => b[1] - a[1])[0];
      const lowOccupancy = weeklyOccupancy.rate !== null && weeklyOccupancy.rate < 0.75;
      const statusClass = row.suspended || lateRate > 0.2 ? "danger" : accurateRate < 0.65 || lowOccupancy ? "warning" : "green";
      const statusText = statusClass === "danger" ? "Alerta" : statusClass === "warning" ? "Observación" : "Control";
      return `<article class="week-card ${statusClass}">
        <div class="week-head">
          <div>
            <span>Semana ${index + 1}</span>
            <strong>${weekRangeLabel(start)}</strong>
          </div>
          <b>${statusText}</b>
        </div>
        <div class="week-metrics">
          <span><strong>${row.total}</strong> programados</span>
          <span><strong>${row.done}</strong> realizados</span>
          <span><strong>${percent(doneRate)}</strong> resolución</span>
          <span><strong>${row.late}</strong> atrasos 1a hora</span>
          <span><strong>${row.suspended}</strong> suspensiones</span>
          <span><strong>${weeklyOccupancy.rate === null ? "S/D" : percent(weeklyOccupancy.rate)}</strong> ocupación</span>
          <span><strong>${percent(accurateRate)}</strong> asertividad</span>
        </div>
        <p>${weeklyOccupancy.rate === null
          ? "Falta horario de funcionamiento para calcular ocupación semanal."
          : `${weeklyOccupancy.usedMinutes.toLocaleString("es-CL", { maximumFractionDigits: 0 })} min usados / ${weeklyOccupancy.availableMinutes.toLocaleString("es-CL", { maximumFractionDigits: 0 })} min disponibles.`
        } ${cause ? `Causa principal de suspensión: ${cause[0]} (${cause[1]})` : "Sin suspensión dominante."}</p>
      </article>`;
    })
    .join("");
}

function barRows(entries, max, emptyText, colorClass = "") {
  if (!entries.length) return `<li class="muted-row">${emptyText}</li>`;
  return entries
    .slice(0, 4)
    .map(([label, count]) => {
      const pct = Math.max(6, Math.min(100, (count / Math.max(max, 1)) * 100));
      return `<li class="${colorClass}">
        <div>
          <span>${label}</span>
          <strong>${count}</strong>
        </div>
        <i><b style="width:${pct}%"></b></i>
      </li>`;
    })
    .join("");
}

function renderWeeklyReasonBars(items) {
  const grouped = new Map();
  items.forEach((item) => {
    const key = recordWeek(item);
    if (!key) return;
    if (!grouped.has(key)) {
      grouped.set(key, {
        total: 0,
        done: 0,
        suspensions: {},
        delays: {},
      });
    }
    const row = grouped.get(key);
    row.total += 1;
    if (isDone(item)) row.done += 1;
    if (isSuspended(item)) {
      const cause = suspensionCause(item);
      row.suspensions[cause] = (row.suspensions[cause] || 0) + 1;
    }
    if (isLate(item)) {
      const label = "Atraso primera hora";
      row.delays[label] = (row.delays[label] || 0) + 1;
    }
    if (Number(item.turnover) > 15) {
      const label = "Recambio >15 min";
      row.delays[label] = (row.delays[label] || 0) + 1;
    }
  });
  const rows = [...grouped.entries()].sort((a, b) => a[0].localeCompare(b[0]));
  const box = document.querySelector("#weeklyReasonBars");
  if (!rows.length) {
    box.innerHTML = `<div class="empty-note">Sin información semanal de causas para el filtro actual.</div>`;
    return;
  }
  box.innerHTML = `<div class="weekly-reason-board">
    <div class="weekly-reason-board-head">
      <span>Semana</span>
      <span>Producción</span>
      <span>Causas de suspensión</span>
      <span>Motivos de atraso</span>
    </div>
    ${rows
    .map(([start, row], index) => {
      const suspensionEntries = Object.entries(row.suspensions).sort((a, b) => b[1] - a[1]);
      const delayEntries = Object.entries(row.delays).sort((a, b) => b[1] - a[1]);
      const maxValue = Math.max(...suspensionEntries.map(([, count]) => count), ...delayEntries.map(([, count]) => count), 1);
      const totalSusp = suspensionEntries.reduce((sum, [, count]) => sum + count, 0);
      const totalDelay = delayEntries.reduce((sum, [, count]) => sum + count, 0);
      const statusClass = totalSusp || totalDelay ? "warning" : "green";
      return `<article class="weekly-reason-row ${statusClass}">
        <div class="weekly-reason-week">
          <span>Semana ${index + 1}</span>
          <strong>${weekRangeLabel(start)}</strong>
          <b>${totalSusp + totalDelay} motivo${totalSusp + totalDelay === 1 ? "" : "s"}</b>
        </div>
        <div class="weekly-production-pair">
          <span><strong>${row.total}</strong> total programados</span>
          <span><strong>${row.done}</strong> total realizados</span>
        </div>
        <section class="weekly-reason-cell">
          <h3>Causas de suspensión <em>${totalSusp}</em></h3>
          <ul class="reason-bars">${barRows(suspensionEntries, maxValue, "Sin suspensiones esta semana.", "danger")}</ul>
        </section>
        <section class="weekly-reason-cell">
          <h3>Motivos de atraso <em>${totalDelay}</em></h3>
          <ul class="reason-bars">${barRows(delayEntries, maxValue, "Sin atrasos ni recambios sobre meta.", "warning")}</ul>
        </section>
      </article>`;
    })
    .join("")}
  </div>`;
}

function renderMetricBars(containerId, rows, config) {
  const box = document.querySelector(`#${containerId}`);
  if (!rows.length) {
    box.innerHTML = `<div class="empty-note">${config.empty}</div>`;
    return;
  }
  const max = Math.max(...rows.map(config.score), 1);
  box.innerHTML = rows
    .map((row) => {
      const pct = Math.max(4, Math.min(100, (config.score(row) / max) * 100));
      return `<div class="metric-row ${config.color}">
        <div>
          <strong>${config.label(row)}</strong>
          <span>${config.detail(row)}</span>
        </div>
        <b>${config.value(row)}</b>
        <i><span style="width:${pct}%"></span></i>
      </div>`;
    })
    .join("");
}

function topEntriesFrom(items, predicate, labelFn) {
  return Object.entries(tally(items.filter(predicate), labelFn)).sort((a, b) => b[1] - a[1]);
}

function justificationCard(title, total, entries, roomEntries, emptyText, className) {
  const max = Math.max(...entries.map(([, count]) => count), 1);
  const rows = entries.length
    ? entries
        .slice(0, 5)
        .map(([label, count]) => {
          const pct = Math.max(7, Math.min(100, (count / max) * 100));
          return `<li class="${className}">
            <div><span>${label}</span><strong>${count}</strong></div>
            <i><b style="width:${pct}%"></b></i>
          </li>`;
        })
        .join("")
    : `<li class="muted-row">${emptyText}</li>`;
  const rooms = roomEntries.length
    ? roomEntries
        .slice(0, 4)
        .map(([room, count]) => `<span class="chip warning">${room}: ${count}</span>`)
        .join("")
    : `<span class="chip green">Sin pabellón crítico</span>`;
  return `<article class="delay-card ${className}">
    <div class="delay-card-head">
      <div>
        <span>Indicador operativo</span>
        <strong>${title}</strong>
      </div>
      <b>${total}<small>casos</small></b>
    </div>
    <div class="delay-card-summary">
      <span>${entries.length ? "Motivo principal" : "Estado"}</span>
      <strong>${entries[0] ? entries[0][0] : emptyText}</strong>
    </div>
    <ul class="reason-bars">${rows}</ul>
    <div class="delay-room-title">Pabellones relacionados</div>
    <div class="delay-room-chips">${rooms}</div>
  </article>`;
}

function renderDelayJustificationSummary(data) {
  const box = document.querySelector("#delayJustificationSummary");
  if (!box) return;
  const late = data.filter(isLate);
  const longTurnover = data.filter((item) => Number(item.turnover) > 15);
  const lateEntries = topEntriesFrom(
    data,
    isLate,
    (item) => item.firstHourJustification || "Sin justificación registrada",
  );
  const turnoverEntries = topEntriesFrom(
    data,
    (item) => Number(item.turnover) > 15,
    (item) => item.turnoverJustification || "Sin justificación registrada",
  );
  const lateRooms = topEntriesFrom(data, isLate, (item) => item.room || "Sin pabellón");
  const turnoverRooms = topEntriesFrom(data, (item) => Number(item.turnover) > 15, (item) => item.room || "Sin pabellón");
  box.innerHTML = [
    justificationCard(
      "Atraso primera hora",
      late.length,
      lateEntries,
      lateRooms,
      "Sin atrasos de primera hora en el filtro.",
      late.length ? "warning" : "green",
    ),
    justificationCard(
      "Recambio mayor a 15 min",
      longTurnover.length,
      turnoverEntries,
      turnoverRooms,
      "Sin recambios mayores a 15 minutos en el filtro.",
      longTurnover.length ? "danger" : "green",
    ),
  ].join("");
}

function renderTrend(items) {
  const daily = new Map();
  items.forEach((item) => {
    if (!item.date) return;
    if (!daily.has(item.date)) daily.set(item.date, { done: 0, late: 0, suspended: 0 });
    const row = daily.get(item.date);
    if (isDone(item)) row.done += 1;
    if (isLate(item)) row.late += 1;
    if (isSuspended(item)) row.suspended += 1;
  });
  const entries = [...daily.entries()].sort((a, b) => a[0].localeCompare(b[0]));
  const canvas = document.querySelector("#trendChart");
  const ctx = canvas.getContext("2d");
  const width = canvas.width;
  const height = canvas.height;
  const pad = 34;
  ctx.clearRect(0, 0, width, height);
  ctx.fillStyle = "#fff";
  ctx.fillRect(0, 0, width, height);
  if (!entries.length) {
    ctx.fillStyle = "#60768b";
    ctx.textAlign = "center";
    ctx.font = "14px Inter, sans-serif";
    ctx.fillText("Sin datos", width / 2, height / 2);
    return;
  }
  const max = Math.max(...entries.flatMap(([, row]) => [row.done, row.late, row.suspended]), 1);
  ctx.strokeStyle = "#d7e3ee";
  ctx.lineWidth = 1;
  for (let i = 0; i <= 3; i += 1) {
    const y = pad + ((height - pad * 2) / 3) * i;
    ctx.beginPath();
    ctx.moveTo(pad, y);
    ctx.lineTo(width - pad, y);
    ctx.stroke();
  }
  const drawLine = (key, color) => {
    ctx.strokeStyle = color;
    ctx.lineWidth = 3;
    ctx.beginPath();
    entries.forEach(([, row], index) => {
      const x = pad + (entries.length === 1 ? 0.5 : index / (entries.length - 1)) * (width - pad * 2);
      const y = height - pad - (row[key] / max) * (height - pad * 2);
      if (index === 0) ctx.moveTo(x, y);
      else ctx.lineTo(x, y);
    });
    ctx.stroke();
    entries.forEach(([, row], index) => {
      const x = pad + (entries.length === 1 ? 0.5 : index / (entries.length - 1)) * (width - pad * 2);
      const y = height - pad - (row[key] / max) * (height - pad * 2);
      ctx.beginPath();
      ctx.arc(x, y, 4, 0, Math.PI * 2);
      ctx.fillStyle = color;
      ctx.fill();
    });
  };
  drawLine("done", "#16639f");
  drawLine("late", "#e8743b");
  drawLine("suspended", "#b91c1c");
  ctx.fillStyle = "#60768b";
  ctx.font = "11px Inter, sans-serif";
  ctx.textAlign = "center";
  entries.forEach(([date], index) => {
    const x = pad + (entries.length === 1 ? 0.5 : index / (entries.length - 1)) * (width - pad * 2);
    ctx.fillText(formatDate(date).slice(0, 5), x, height - 10);
  });
}

function renderRanking(data) {
  const title = document.querySelector("#roomRankingWeek");
  const weekItems =
    filters.date.value && filters.date.value !== "todos"
      ? data.filter((item) => recordWeek(item) === weekStart(filters.date.value))
      : latestWeekItems(data);
  if (title) title.textContent = weekItems.length ? weekRangeLabel(weekStart(weekItems[0].date)) : "Sin semana disponible";
  if (!weekItems.length) {
    document.querySelector("#roomRanking").innerHTML = `<div class="empty-note">Sin datos semanales para ranking operativo.</div>`;
    return;
  }
  const grouped = new Map();
  weekItems.forEach((item) => {
    if (!grouped.has(item.room)) {
      grouped.set(item.room, {
        total: 0,
        done: 0,
        late: 0,
        suspended: 0,
        accurate: 0,
        longTurnover: 0,
        items: [],
        reasons: {},
      });
    }
    const row = grouped.get(item.room);
    row.items.push(item);
    row.total += 1;
    if (isDone(item)) row.done += 1;
    if (isLate(item)) {
      row.late += 1;
      const reason = item.firstHourJustification || "Atraso primera hora sin justificación";
      row.reasons[reason] = (row.reasons[reason] || 0) + 1;
    }
    if (Number(item.turnover) > 15) {
      row.longTurnover += 1;
      const reason = item.turnoverJustification || "Recambio >15 min sin justificación";
      row.reasons[reason] = (row.reasons[reason] || 0) + 1;
    }
    if (isSuspended(item)) {
      row.suspended += 1;
      const reason = suspensionDetail(item);
      row.reasons[reason] = (row.reasons[reason] || 0) + 1;
    }
    if (isAccurate(item)) row.accurate += 1;
  });
  document.querySelector("#roomRanking").innerHTML = [...grouped.entries()]
    .map(([room, row]) => {
      const occupancy = occupancyStats(row.items);
      const occupancyRate = occupancy.rate;
      const scoreOccupancy = occupancyRate === null ? 0 : occupancyRate;
      const resolution = row.total ? row.done / row.total : 0;
      const mainReason = Object.entries(row.reasons).sort((a, b) => b[1] - a[1])[0];
      const topSpecialty = topLabel(row.items, (item) => item.specialty, "Sin especialidad");
      const topSubspecialty = topLabel(row.items, (item) => item.subspecialty, "Sin subespecialidad");
      const score =
        row.suspended * 5 +
        row.late * 3 +
        row.longTurnover * 3 +
        Math.max(0, 0.75 - scoreOccupancy) * 8 +
        Math.max(0, 0.85 - resolution) * 4;
      return { room, row, occupancy, occupancyRate, resolution, mainReason, topSpecialty, topSubspecialty, score };
    })
    .sort((a, b) => b.score - a.score || b.row.total - a.row.total || a.room.localeCompare(b.room))
    .map(({ room, row, occupancyRate, resolution, mainReason, topSpecialty, topSubspecialty, score }) => {
      const cls = row.suspended || score >= 8 ? "danger" : row.late || row.longTurnover || (occupancyRate !== null && occupancyRate < 0.75) ? "warning" : "green";
      const status = cls === "danger" ? "Crítico" : cls === "warning" ? "Alerta" : "Control";
      const barWidth = Math.min(100, Math.max(6, score * 8));
      return `<article class="rank-row ${cls}">
        <div class="rank-room-title">
          <strong>${room}</strong>
          <span>${status}</span>
        </div>
        <span>${row.done}/${row.total} realizados o atendidos · ocupación ${occupancyRate === null ? "S/D" : percent(occupancyRate)}</span>
        <div class="rank-service">
          <b>Especialidad: ${topSpecialty}</b>
          <small>Subespecialidad: ${topSubspecialty}</small>
        </div>
        <div class="rank-bar"><span style="width:${barWidth}%"></span></div>
        <div class="rank-metrics">
          <span class="chip">${row.total} casos</span>
          <span class="chip green">${percent(resolution)} resolución</span>
          <span class="chip">${occupancyRate === null ? "S/D" : percent(occupancyRate)} ocup.</span>
          <span class="chip warning">${row.late} atrasos</span>
          <span class="chip warning">${row.longTurnover} recambios >15</span>
          <span class="chip danger">${row.suspended} susp.</span>
        </div>
        <p>${mainReason ? `Justificación principal: ${mainReason[0]} (${mainReason[1]})` : "Sin justificaciones críticas registradas esta semana."}</p>
      </article>`;
    })
    .join("");
}

function dailyRoomCard(row) {
  const doneRate = row.total ? row.done / row.total : 0;
  const accurateRate = row.total ? row.accurate / row.total : 0;
  const turnoverAvg = average(row.turnovers);
  const occupiedMinutes = row.starts.length && row.ends.length ? Math.max(...row.ends) - Math.min(...row.starts) : 0;
  const occupancyRate = Number(row.workingMinutes) > 0 ? occupiedMinutes / Number(row.workingMinutes) : null;
  const causeText = Object.entries(row.causes)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 3)
    .map(([cause, count]) => `${cause} (${count})`);
  const specialtyText = Object.entries(row.specialties).sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))[0];
  const subspecialtyText = Object.entries(row.subspecialties).sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))[0];
  const delayText = Object.entries(row.delayReasons)
    .sort((a, b) => b[1] - a[1])
    .map(([reason, count]) => `${reason} (${count})`);
  const firstHourJustificationText = Object.entries(row.firstHourJustifications)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 2)
    .map(([reason, count]) => `Atraso: ${reason} (${count})`);
  const turnoverJustificationText = Object.entries(row.turnoverJustifications)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 2)
    .map(([reason, count]) => `Recambio: ${reason} (${count})`);
  const suspensionJustificationText = Object.entries(row.suspensionJustifications)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 2)
    .map(([reason, count]) => `Suspensión: ${reason} (${count})`);
  const justification = [...causeText, ...delayText, ...firstHourJustificationText, ...turnoverJustificationText, ...suspensionJustificationText];
  const cls = row.suspended ? "danger" : row.late || row.longTurnover || (occupancyRate !== null && occupancyRate < 0.75) ? "warning" : "green";
  return `<article class="daily-room-card ${cls}">
    <div class="daily-room-head">
      <div>
        <span>${formatDate(row.date)}</span>
        <strong>${row.room}</strong>
        <small class="daily-room-service">${specialtyText ? `Especialidad: ${specialtyText[0]} (${specialtyText[1]})` : "Especialidad: sin dato"}${subspecialtyText ? ` · Subespecialidad: ${subspecialtyText[0]} (${subspecialtyText[1]})` : ""}</small>
      </div>
      <b>${row.total} programado${row.total === 1 ? "" : "s"}</b>
    </div>
    <div class="daily-room-metrics">
      <span><strong>${row.total}</strong> producción</span>
      <span><strong>${row.done}</strong> realizados</span>
      <span><strong>${row.suspended}</strong> susp.</span>
      <span><strong>${row.late}</strong> atraso 1a hora</span>
      <span><strong>${row.longTurnover}</strong> recambio >15</span>
      <span><strong>${occupancyRate === null ? "S/D" : percent(occupancyRate)}</strong> ocupación</span>
      <span><strong>${percent(doneRate)}</strong> resolución</span>
    </div>
    <div class="daily-room-justification">
      <h3>Justificación</h3>
      <p>${justification.length ? justification.join(" · ") : "Sin suspensión, atraso de primera hora ni recambio mayor a 15 minutos."}</p>
      <small>${occupancyRate === null
        ? "Ocupación: falta horario de funcionamiento."
        : `Ocupación: ${occupiedMinutes.toLocaleString("es-CL", { maximumFractionDigits: 0 })} min ocupados / ${Number(row.workingMinutes).toLocaleString("es-CL", { maximumFractionDigits: 0 })} min programados.`
      } ${row.turnovers.length ? `Recambio promedio: ${turnoverAvg.toLocaleString("es-CL", { maximumFractionDigits: 1 })} min · Asertividad: ${percent(accurateRate)}` : `Asertividad: ${percent(accurateRate)}`}</small>
    </div>
  </article>`;
}

function renderDailyRoomProduction(data) {
  const grouped = new Map();
  data.forEach((item) => {
    const key = `${item.date || "Sin fecha"}||${item.room || "Sin pabellón"}`;
    if (!grouped.has(key)) {
      grouped.set(key, {
        date: item.date,
        room: item.room || "Sin pabellón",
        total: 0,
        done: 0,
        late: 0,
        suspended: 0,
        accurate: 0,
        starts: [],
        ends: [],
        workingMinutes: "",
        turnovers: [],
        longTurnover: 0,
        causes: {},
        delayReasons: {},
        firstHourJustifications: {},
        turnoverJustifications: {},
        suspensionJustifications: {},
        specialties: {},
        subspecialties: {},
      });
    }
    const row = grouped.get(key);
    row.total += 1;
    if (item.specialty && !/^sin\s/i.test(item.specialty)) row.specialties[item.specialty] = (row.specialties[item.specialty] || 0) + 1;
    if (item.subspecialty && !/^sin\s/i.test(item.subspecialty)) row.subspecialties[item.subspecialty] = (row.subspecialties[item.subspecialty] || 0) + 1;
    const working = Number(item.workingMinutes);
    if (Number.isFinite(working) && working > 0) row.workingMinutes = working;
    if (isDone(item)) {
      row.done += 1;
      const start = timeToMinutes(item.surgeryStart);
      const end = timeToMinutes(item.surgeryEnd);
      if (start !== null && end !== null) {
        row.starts.push(start);
        row.ends.push(end);
      }
    }
    if (isLate(item)) {
      row.late += 1;
      row.delayReasons["Atraso primera hora"] = (row.delayReasons["Atraso primera hora"] || 0) + 1;
      if (item.firstHourJustification) {
        row.firstHourJustifications[item.firstHourJustification] = (row.firstHourJustifications[item.firstHourJustification] || 0) + 1;
      }
    }
    if (isSuspended(item)) {
      row.suspended += 1;
      const cause = suspensionDetail(item);
      row.causes[cause] = (row.causes[cause] || 0) + 1;
      if (item.suspensionJustification) {
        row.suspensionJustifications[item.suspensionJustification] = (row.suspensionJustifications[item.suspensionJustification] || 0) + 1;
      }
    }
    if (isAccurate(item)) row.accurate += 1;
    const turnover = Number(item.turnover);
    if (item.turnover !== "" && Number.isFinite(turnover)) {
      row.turnovers.push(turnover);
      if (turnover > 15) {
        row.longTurnover += 1;
        row.delayReasons["Recambio >15 min"] = (row.delayReasons["Recambio >15 min"] || 0) + 1;
        if (item.turnoverJustification) {
          row.turnoverJustifications[item.turnoverJustification] = (row.turnoverJustifications[item.turnoverJustification] || 0) + 1;
        }
      }
    }
  });
  const rows = [...grouped.values()].sort((a, b) => {
    const dateOrder = (a.date || "").localeCompare(b.date || "");
    return dateOrder || normalizeRoom(a.room).localeCompare(normalizeRoom(b.room), "es", { numeric: true });
  });
  document.querySelector("#dailyRoomCount").textContent = `${rows.length} registro${rows.length === 1 ? "" : "s"}`;
  const box = document.querySelector("#dailyRoomProduction");
  if (!rows.length) {
    box.innerHTML = `<div class="empty-note">Sin producción por quirófano para el filtro actual.</div>`;
    return;
  }
  const byDate = new Map();
  rows.forEach((row) => {
    const key = row.date || "Sin fecha";
    if (!byDate.has(key)) byDate.set(key, []);
    byDate.get(key).push(row);
  });
  box.innerHTML = [...byDate.entries()]
    .map(([date, dayRows]) => {
      const total = dayRows.reduce((sum, row) => sum + row.total, 0);
      const done = dayRows.reduce((sum, row) => sum + row.done, 0);
      const suspended = dayRows.reduce((sum, row) => sum + row.suspended, 0);
      const late = dayRows.reduce((sum, row) => sum + row.late, 0);
      const longTurnover = dayRows.reduce((sum, row) => sum + row.longTurnover, 0);
      return `<section class="daily-room-day">
        <div class="daily-room-day-head">
          <div>
            <span>Día quirúrgico</span>
            <strong>${formatDate(date)}</strong>
          </div>
          <ul>
            <li>${dayRows.length} ${dayRows.length === 1 ? "pabellón" : "pabellones"}</li>
            <li>${total} programados</li>
            <li>${done} realizados</li>
            <li>${suspended} susp.</li>
            <li>${late} atraso 1a hora</li>
            <li>${longTurnover} recambio >15</li>
          </ul>
        </div>
        <div class="daily-room-day-grid">${dayRows.map(dailyRoomCard).join("")}</div>
      </section>`;
    })
    .join("");
}

function renderAttentionRows(data) {
  const attention = data
    .filter((item) => isLate(item) || isSuspended(item) || /subestimacion|sobrestimacion/i.test(item.accuracy) || Number(item.turnover) > 15)
    .sort((a, b) => a.room.localeCompare(b.room) || a.entryTime.localeCompare(b.entryTime));
  document.querySelector("#attentionCount").textContent = `${attention.length} caso${attention.length === 1 ? "" : "s"}`;
  document.querySelector("#attentionRows").innerHTML = attention.length
    ? attention
        .map((item) => {
          const statusClass = isSuspended(item) ? "danger" : isDone(item) ? "green" : "warning";
          const entryClass = isLate(item) ? "warning" : "green";
          const accuracyClass = /subestimacion/i.test(item.accuracy) ? "danger" : /sobrestimacion/i.test(item.accuracy) ? "warning" : "green";
          const turnoverValue = Number(item.turnover);
          const turnoverText = Number.isFinite(turnoverValue) ? `${turnoverValue.toLocaleString("es-CL", { maximumFractionDigits: 0 })} min` : "Pendiente";
          const turnoverClass = Number.isFinite(turnoverValue) && turnoverValue > 15 ? "warning" : "green";
          const detail = [
            isSuspended(item) ? suspensionDetail(item) : item.surgery,
            isLate(item) && item.firstHourJustification ? `Atraso: ${item.firstHourJustification}` : "",
            Number.isFinite(turnoverValue) && turnoverValue > 15 && item.turnoverJustification ? `Recambio: ${item.turnoverJustification}` : "",
          ]
            .filter(Boolean)
            .join(" · ");
          return `<tr>
            <td>${formatDate(item.date)}</td>
            <td><strong>${item.patient || "Sin paciente"}</strong></td>
            <td>${item.room}</td>
            <td><span class="badge ${statusClass}">${item.status}</span></td>
            <td>${item.entryTime}</td>
            <td><span class="badge ${entryClass}">${item.entryClass || "Pendiente"}</span></td>
            <td><span class="badge ${accuracyClass}">${item.accuracy || "Pendiente"}</span></td>
            <td><span class="badge ${turnoverClass}">${turnoverText}</span></td>
            <td>${item.specialty}</td>
            <td>${item.subspecialty}</td>
            <td>${detail || "Sin detalle"}</td>
          </tr>`;
        })
        .join("")
    : `<tr><td colspan="11">Sin casos críticos para este filtro.</td></tr>`;
}

function renderDonut(canvasId, legendId, data) {
  const canvas = document.querySelector(`#${canvasId}`);
  const legend = document.querySelector(`#${legendId}`);
  const entries = Object.entries(data).filter(([, value]) => value > 0).sort((a, b) => b[1] - a[1]);
  const total = entries.reduce((sum, [, value]) => sum + value, 0);
  if (canvas) canvas.hidden = true;
  if (!total) {
    legend.innerHTML = `<div class="chart-bars empty-note">Sin datos para el filtro actual.</div>`;
    return;
  }
  const max = Math.max(...entries.map(([, value]) => value), 1);
  legend.innerHTML = `<div class="chart-bars">
    <div class="chart-bars-total">
      <span>Total del filtro</span>
      <strong>${total}</strong>
      <small>casos</small>
    </div>
    ${entries
    .slice(0, 7)
    .map(([label, value], index) => {
      const width = Math.max(8, (value / max) * 100);
      return `<article class="chart-bar-row">
        <div class="chart-bar-label">
          <span class="legend-dot" style="background:${palette[index % palette.length]}"></span>
          <strong>${label}</strong>
          <small>${value} · ${percent(value / total)}</small>
        </div>
        <i><b style="width:${width}%; background:${palette[index % palette.length]}"></b></i>
      </article>`;
    })
    .join("")}
  </div>`;
}

filters.reload.addEventListener("click", loadExcel);
document.querySelector("#toggleCauseBreakdown").addEventListener("click", () => {
  const panel = document.querySelector("#causeBreakdown");
  const button = document.querySelector("#toggleCauseBreakdown");
  const isOpen = !panel.hidden;
  panel.hidden = isOpen;
  button.setAttribute("aria-expanded", String(!isOpen));
  button.textContent = isOpen ? "Desglose" : "Ocultar";
});
filters.clear.addEventListener("click", () => {
  const latestDate = records.map((item) => item.date).filter(Boolean).sort().at(-1) || "";
  filters.month.value = monthKey(latestDate) || "todos";
  filters.week.value = weekStart(latestDate) || "todos";
  filters.date.value = "todos";
  filters.specialty.value = "todos";
  filters.subspecialty.value = "todos";
  filters.room.value = "todos";
  render();
});
filters.month.addEventListener("change", () => {
  filters.week.value = "todos";
  filters.date.value = "todos";
  render();
});
filters.week.addEventListener("change", () => {
  if (filters.week.value !== "todos") filters.month.value = monthKey(filters.week.value);
  filters.date.value = "todos";
  render();
});
filters.date.addEventListener("change", () => {
  if (filters.date.value !== "todos") {
    filters.month.value = monthKey(filters.date.value);
    filters.week.value = weekStart(filters.date.value);
  }
  render();
});
filters.specialty.addEventListener("change", render);
filters.subspecialty.addEventListener("change", render);
filters.room.addEventListener("change", render);

function activateTab(tabName) {
  document.querySelectorAll("[data-tab-target]").forEach((button) => {
    const active = button.dataset.tabTarget === tabName;
    button.classList.toggle("active", active);
    button.setAttribute("aria-selected", String(active));
  });
  document.querySelectorAll("[data-tab-panel]").forEach((panel) => {
    panel.hidden = panel.dataset.tabPanel !== tabName;
  });
}

document.querySelectorAll("[data-tab-target]").forEach((button) => {
  button.setAttribute("role", "tab");
  button.setAttribute("aria-selected", button.classList.contains("active") ? "true" : "false");
  button.addEventListener("click", () => activateTab(button.dataset.tabTarget));
});

document.querySelector(".dashboard-tabs")?.setAttribute("role", "tablist");
activateTab(document.querySelector("[data-tab-target].active")?.dataset.tabTarget || "resumen");

loadExcel().catch((error) => {
  console.error(error);
  document.querySelector("#statusTitle").textContent = "No se pudo leer el Excel";
  document.querySelector("#statusText").textContent = error.message;
});
