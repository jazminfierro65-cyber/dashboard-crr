const STORAGE_KEY = "crr-pabellones-dia-data";
const SURGERY_PROGRAMMED_KEY = "crr-surgery-programmed-minutes";
const SURGERY_RECORDS_KEY = "crr-surgery-records-from-excel";
const EXCEL_TEMPLATE_URL = "./outputs/crr_dashboard/plantilla_carga_web_crr.xlsx";

const nodes = [
  "Admision / citacion",
  "Preoperatorio",
  "Anestesia",
  "Esterilizacion / instrumental",
  "Disponibilidad de cama",
  "Pabellon",
  "Equipo quirurgico",
  "Traslado paciente",
  "Gestion de tabla",
];

const initialDaily = [
  {
    date: "2026-05-18",
    programmed: 13,
    operated: 11,
    suspended: 2,
    firstDelay: 33,
    occupancy: 0.98,
    turnover: 18,
    note: "Instrumental pendiente y suspensiones por prolongacion de tabla / estudio incompleto.",
  },
  {
    date: "2026-05-19",
    programmed: 12,
    operated: 10,
    suspended: 1,
    firstDelay: 12,
    occupancy: 0.96,
    turnover: 19,
    note: "Recambio afectado por piso humedo y paciente en baño antes de entrar a pabellon.",
  },
  {
    date: "2026-05-20",
    programmed: 10,
    operated: 8,
    suspended: 2,
    firstDelay: 52.5,
    occupancy: 1.08,
    turnover: 17,
    note: "Reunion de anestesia, instrumental pendiente y cirujano en policlinico.",
  },
  {
    date: "2026-05-22",
    programmed: 13,
    operated: 12,
    suspended: 2,
    firstDelay: 8,
    occupancy: 1.04,
    turnover: 17,
    note: "Revision de cajas y cierre tardio de penultima cirugia.",
  },
];

const initialEvents = [
  event("2026-05-18", "Demora ingreso >15", "Oftalmologia", "Pab 1", "Esterilizacion / instrumental", 33, 1, "Pendiente entrega instrumental"),
  event("2026-05-20", "Demora ingreso >15", "Otorrino", "Pab 2", "Anestesia", 60, 1, "Reunion de anestesia; pabellon comienza 09:00"),
  event("2026-05-20", "Demora ingreso >15", "Oftalmologia", "Pab 3", "Esterilizacion / instrumental", 45, 1, "Pendiente entrega instrumental"),
  event("2026-05-19", "Recambio >15", "Cirugia adulto", "Pab 2", "Pabellon", 16, 1, "Piso humedo"),
  event("2026-05-19", "Recambio >15", "Cirugia adulto", "Pab 3", "Traslado paciente", 15, 1, "Paciente en el bano antes de entrar a pabellon"),
  event("2026-05-18", "Recambio >15", "Traumatologia", "Pab 1", "Esterilizacion / instrumental", 18, 5, "Revision de cajas"),
  event("2026-05-20", "Recambio >15", "Oftalmologia", "Pab 1", "Equipo quirurgico", 17, 4, "Cirujano en policlinico"),
  event("2026-05-22", "Recambio >15", "Traumatologia", "Pab 4", "Esterilizacion / instrumental", 17, 3, "Revision de cajas"),
  event("2026-05-18", "Suspensión", "Oftalmologia", "Pab 1", "Gestion de tabla", 0, 1, "Prolongacion de tabla"),
  event("2026-05-18", "Suspensión", "TMT infantil", "Pab 2", "Preoperatorio", 0, 1, "Estudio incompleto"),
  event("2026-05-19", "Suspensión", "C. general", "Pab 2", "Gestion de tabla", 0, 1, "Penultima cirugia termina 18:57"),
  event("2026-05-20", "Suspensión", "Otorrino", "Pab 3", "Gestion de tabla", 0, 2, "Penultima cirugia termina 17:30"),
  event("2026-05-22", "Suspensión", "TMT mano", "Pab 4", "Gestion de tabla", 0, 2, "Penultima cirugia termina 17:37"),
];

let state = loadState();

const dateFilter = document.querySelector("#dateFilter");
const nodeFilter = document.querySelector("#nodeFilter");
const eventTypeFilter = document.querySelector("#eventTypeFilter");
const selectedDayLabel = document.querySelector("#selectedDayLabel");
const dailyList = document.querySelector("#dailyList");
const nodeMap = document.querySelector("#nodeMap");
const suspensionBars = document.querySelector("#suspensionBars");
const eventRows = document.querySelector("#eventRows");
const eventForm = document.querySelector("#eventForm");
let surgeryRecords = loadSurgeryRecords();
let programmedTimes = loadProgrammedTimes();

function event(date, type, specialty, room, node, minutes, count, cause) {
  return { id: crypto.randomUUID(), date, type, specialty, room, node, minutes, count, cause };
}

function loadState() {
  const saved = localStorage.getItem(STORAGE_KEY);
  if (!saved) return { daily: initialDaily, events: initialEvents };
  try {
    return JSON.parse(saved);
  } catch {
    return { daily: initialDaily, events: initialEvents };
  }
}

function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function loadProgrammedTimes() {
  const saved = localStorage.getItem(SURGERY_PROGRAMMED_KEY);
  if (!saved) return {};
  try {
    return JSON.parse(saved);
  } catch {
    return {};
  }
}

function loadSurgeryRecords() {
  const saved = localStorage.getItem(SURGERY_RECORDS_KEY);
  if (saved) {
    try {
      const parsed = JSON.parse(saved);
      if (Array.isArray(parsed)) return parsed;
    } catch {
      localStorage.removeItem(SURGERY_RECORDS_KEY);
    }
  }
  return window.CRR_SURGERY_RECORDS || [];
}

function saveProgrammedTimes() {
  localStorage.setItem(SURGERY_PROGRAMMED_KEY, JSON.stringify(programmedTimes));
}

function saveSurgeryRecords() {
  localStorage.setItem(SURGERY_RECORDS_KEY, JSON.stringify(surgeryRecords));
}

function formatDate(value) {
  if (value === "todos") return "Todos los dias";
  const [year, month, day] = value.split("-");
  return `${day}-${month}-${year}`;
}

function percent(value) {
  return `${(value * 100).toLocaleString("es-CL", { maximumFractionDigits: 1 })}%`;
}

function timeToMinutes(value) {
  if (!value) return null;
  const [hours, minutes] = String(value).split(":").map(Number);
  if (!Number.isFinite(hours) || !Number.isFinite(minutes)) return null;
  return hours * 60 + minutes;
}

function filteredEvents() {
  const date = dateFilter.value;
  const node = nodeFilter.value;
  const type = eventTypeFilter?.value || "todos";
  return state.events.filter((item) => {
    const dateOk = date === "todos" || item.date === date;
    const nodeOk = node === "todos" || item.node === node;
    const typeOk = type === "todos" || item.type === type;
    return dateOk && nodeOk && typeOk;
  });
}

function filteredDaily() {
  const date = dateFilter.value;
  if (date === "todos") return state.daily;
  return state.daily.filter((item) => item.date === date);
}

function initControls() {
  const dates = [...new Set(state.daily.map((item) => item.date).concat(state.events.map((item) => item.date)))].sort();
  dateFilter.innerHTML = `<option value="todos">Todos los dias</option>${dates
    .map((date) => `<option value="${date}">${formatDate(date)}</option>`)
    .join("")}`;
  nodeFilter.innerHTML = `<option value="todos">Todos</option>${nodes.map((node) => `<option>${node}</option>`).join("")}`;
  const nodeSelect = eventForm.elements.node;
  nodeSelect.innerHTML = nodes.map((node) => `<option>${node}</option>`).join("");
  eventForm.elements.date.value = dates.at(-1) || new Date().toISOString().slice(0, 10);
  initSurgeryControls();
}

function initSurgeryControls() {
  const dateFilter = document.querySelector("#surgeryDateFilter");
  const specialtyFilter = document.querySelector("#surgerySpecialtyFilter");
  const statusFilter = document.querySelector("#surgeryStatusFilter");
  const categoryFilter = document.querySelector("#surgeryCategoryFilter");
  if (!dateFilter || !specialtyFilter || !statusFilter || !categoryFilter) return;
  const currentDate = dateFilter.value || "todos";
  const currentSpecialty = specialtyFilter.value || "todos";
  const currentStatus = statusFilter.value || "todos";
  const currentCategory = categoryFilter.value || "todos";
  const dates = [...new Set(surgeryRecords.map((item) => item.date || "Sin fecha"))].sort();
  const specialties = [...new Set(surgeryRecords.map((item) => item.specialty || "Sin especialidad"))].sort();
  const statuses = [...new Set(surgeryRecords.map((item) => item.status || "Sin estado"))].sort();
  const categories = [...new Set(surgeryRecords.map((item) => item.category || "Sin categoria"))].sort();
  dateFilter.innerHTML = `<option value="todos">Todos los días</option>${dates.map((item) => `<option>${item}</option>`).join("")}`;
  specialtyFilter.innerHTML = `<option value="todos">Todas</option>${specialties.map((item) => `<option>${item}</option>`).join("")}`;
  statusFilter.innerHTML = `<option value="todos">Todos</option>${statuses.map((item) => `<option>${item}</option>`).join("")}`;
  categoryFilter.innerHTML = `<option value="todos">Todas</option>${categories.map((item) => `<option>${item}</option>`).join("")}`;
  dateFilter.value = dates.includes(currentDate) ? currentDate : "todos";
  specialtyFilter.value = specialties.includes(currentSpecialty) ? currentSpecialty : "todos";
  statusFilter.value = statuses.includes(currentStatus) ? currentStatus : "todos";
  categoryFilter.value = categories.includes(currentCategory) ? currentCategory : "todos";
}

function renderKpis() {
  const daily = filteredDaily();
  const totals = daily.reduce(
    (acc, item) => {
      acc.programmed += Number(item.programmed) || 0;
      acc.operated += Number(item.operated) || 0;
      acc.suspended += Number(item.suspended) || 0;
      acc.delay += Number(item.firstDelay) || 0;
      acc.days += 1;
      return acc;
    },
    { programmed: 0, operated: 0, suspended: 0, delay: 0, days: 0 },
  );
  const suspensionRate = totals.programmed ? totals.suspended / totals.programmed : 0;
  const avgDelay = totals.days ? totals.delay / totals.days : 0;
  document.querySelector("#kpiProgramados").textContent = totals.programmed;
  document.querySelector("#kpiOperados").textContent = totals.operated;
  document.querySelector("#kpiSuspensiones").textContent = percent(suspensionRate);
  document.querySelector("#kpiSuspensionesDetail").textContent = `${totals.suspended} casos`;
  document.querySelector("#kpiDemora").textContent = `${avgDelay.toLocaleString("es-CL", { maximumFractionDigits: 1 })} min`;
  renderExecutiveStatus(totals, suspensionRate, avgDelay);
}

function renderExecutiveStatus(totals, suspensionRate, avgDelay) {
  const card = document.querySelector("#operationalStatus");
  const title = document.querySelector("#statusTitle");
  const detail = document.querySelector("#statusDetail");
  const focus = document.querySelector("#priorityFocus");
  const focusDetail = document.querySelector("#priorityDetail");
  const action = document.querySelector("#suggestedAction");
  if (!card || !title || !detail || !focus || !focusDetail || !action) return;
  card.classList.remove("warning", "critical");
  const events = filteredEvents();
  const nodeTally = tallyBy(events, (item) => item.node);
  const topNode = Object.entries(nodeTally).sort((a, b) => b[1] - a[1])[0];
  if (!totals.programmed) {
    title.textContent = "Sin jornada seleccionada";
    detail.textContent = "Seleccione fecha o cargue datos para ver el estado.";
    focus.textContent = "Cargar datos";
    focusDetail.textContent = "Sin datos no es posible priorizar.";
    action.textContent = "Actualizar desde Excel o registrar evento";
    return;
  }
  if (suspensionRate > 0.1 || avgDelay > 30) {
    card.classList.add("critical");
    title.textContent = "Jornada crítica";
    action.textContent = "Escalar causa principal hoy";
  } else if (suspensionRate > 0.06 || avgDelay > 15) {
    card.classList.add("warning");
    title.textContent = "Jornada en observación";
    action.textContent = "Asignar responsable del nodo principal";
  } else {
    title.textContent = "Jornada dentro de meta";
    action.textContent = "Mantener seguimiento diario";
  }
  detail.textContent = `${totals.operated}/${totals.programmed} operados, ${totals.suspended} suspendidos, demora promedio ${avgDelay.toLocaleString("es-CL", { maximumFractionDigits: 1 })} min.`;
  focus.textContent = topNode ? topNode[0] : "Sin nodo dominante";
  focusDetail.textContent = topNode ? `${topNode[1]} eventos registrados en el nodo.` : "Registre causas para identificar el punto crítico.";
}

function renderDailyList() {
  const daily = filteredDaily();
  selectedDayLabel.textContent = formatDate(dateFilter.value);
  dailyList.innerHTML = daily
    .map((item) => {
      const rate = item.programmed ? item.suspended / item.programmed : 0;
      const delayClass = item.firstDelay > 15 ? "bad" : "warn";
      return `<div class="day-card">
        <div>
          <strong>${formatDate(item.date)}</strong>
          <p>${item.operated}/${item.programmed} operados</p>
        </div>
        <div>
          <strong>${item.note}</strong>
          <p>Ocupacion ${percent(item.occupancy)} · recambio ${item.turnover} min · suspension ${percent(rate)}</p>
        </div>
        <span class="badge ${delayClass}">${item.firstDelay} min</span>
      </div>`;
    })
    .join("");
}

function tallyBy(items, keyFn, countFn = (item) => item.count || 1) {
  return items.reduce((acc, item) => {
    const key = keyFn(item);
    acc[key] = (acc[key] || 0) + Number(countFn(item));
    return acc;
  }, {});
}

function renderNodeMap() {
  const tally = tallyBy(filteredEvents(), (item) => item.node);
  const rows = Object.entries(tally).sort((a, b) => b[1] - a[1]);
  const max = Math.max(...rows.map(([, value]) => value), 1);
  nodeMap.innerHTML = rows.length
    ? rows
        .map(([node, value], index) => `<div class="node-row">
          <strong>${node}</strong><span>${value}</span>
          <div class="track"><div class="fill blue" style="width:${(value / max) * 100}%"></div></div>
        </div>`)
        .join("")
    : `<p>No hay eventos para este filtro.</p>`;
}

function renderSuspensions() {
  const suspensions = filteredEvents().filter((item) => item.type === "Suspensión");
  const tally = tallyBy(suspensions, (item) => item.cause);
  const rows = Object.entries(tally).sort((a, b) => b[1] - a[1]);
  const max = Math.max(...rows.map(([, value]) => value), 1);
  suspensionBars.innerHTML = rows.length
    ? rows
        .map(([cause, value], index) => `<div class="bar-row">
          <strong>${cause}</strong><span>${value}</span>
          <div class="track"><div class="fill blue" style="width:${(value / max) * 100}%"></div></div>
        </div>`)
        .join("")
    : `<p>No hay suspensiones para este filtro.</p>`;
}

function renderRows() {
  eventRows.innerHTML = filteredEvents()
    .sort((a, b) => b.date.localeCompare(a.date))
    .map((item) => `<tr>
      <td>${formatDate(item.date)}</td>
      <td>${item.type}</td>
      <td>${item.specialty}</td>
      <td>${item.room}</td>
      <td>${item.node}</td>
      <td>${item.minutes}</td>
      <td>${item.count}</td>
      <td>${item.cause}</td>
    </tr>`)
    .join("");
}

function renderDelayChart() {
  const canvas = document.querySelector("#delayChart");
  const ctx = canvas.getContext("2d");
  const data = filteredDaily();
  const pad = 42;
  const width = canvas.width;
  const height = canvas.height;
  const chartWidth = width - pad * 2;
  const chartHeight = height - pad * 2;
  const max = Math.max(60, ...data.map((item) => item.firstDelay));
  ctx.clearRect(0, 0, width, height);
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, width, height);
  if (!data.length) {
    ctx.fillStyle = "#60768b";
    ctx.font = "14px Inter, sans-serif";
    ctx.fillText("Sin datos para el filtro seleccionado", 42, height / 2);
    return;
  }
  ctx.strokeStyle = "#d8e0e8";
  ctx.lineWidth = 1;
  for (let i = 0; i <= 4; i += 1) {
    const y = pad + (chartHeight / 4) * i;
    ctx.beginPath();
    ctx.moveTo(pad, y);
    ctx.lineTo(width - pad, y);
    ctx.stroke();
  }
  data.forEach((item, index) => {
    const slot = chartWidth / data.length;
    const barWidth = Math.min(28, Math.max(12, slot * 0.36));
    const x = pad + index * slot + (slot - barWidth) / 2;
    const barHeight = (item.firstDelay / max) * chartHeight;
    const y = height - pad - barHeight;
    ctx.fillStyle = "#16639f";
    ctx.fillRect(x, y, barWidth, barHeight);
    ctx.fillStyle = "#102033";
    ctx.font = "13px Inter, sans-serif";
    ctx.fillText(`${item.firstDelay}`, x, y - 8);
    ctx.fillStyle = "#64748b";
    ctx.font = "12px Inter, sans-serif";
    ctx.fillText(item.date.slice(5), x - 2, height - 16);
  });
  const targetY = height - pad - (15 / max) * chartHeight;
  ctx.strokeStyle = "#d97706";
  ctx.setLineDash([6, 5]);
  ctx.beginPath();
  ctx.moveTo(pad, targetY);
  ctx.lineTo(width - pad, targetY);
  ctx.stroke();
  ctx.setLineDash([]);
  ctx.fillStyle = "#d97706";
  ctx.fillText("Meta 15 min", width - pad - 80, targetY - 7);
}

function surgeryFilteredRecords() {
  const date = document.querySelector("#surgeryDateFilter")?.value || "todos";
  const specialty = document.querySelector("#surgerySpecialtyFilter")?.value || "todos";
  const status = document.querySelector("#surgeryStatusFilter")?.value || "todos";
  const category = document.querySelector("#surgeryCategoryFilter")?.value || "todos";
  return surgeryRecords.filter((item) => {
    const dateOk = date === "todos" || item.date === date;
    const specialtyOk = specialty === "todos" || item.specialty === specialty;
    const statusOk = status === "todos" || item.status === status;
    const categoryOk = category === "todos" || (item.category || "Sin categoria") === category;
    return dateOk && specialtyOk && statusOk && categoryOk;
  });
}

function classifyAsertividad(difference) {
  if (difference === "" || Number.isNaN(difference)) return { label: "Pendiente", cls: "sobre" };
  if (difference > 15) return { label: "Subestimacion", cls: "sub" };
  if (difference < -15) return { label: "Sobrestimacion", cls: "sobre" };
  return { label: "Asertivo", cls: "ok" };
}

function renderSurgeryAnalysis() {
  const body = document.querySelector("#surgeryRows");
  if (!body) return;
  const records = surgeryFilteredRecords();
  const completed = records.filter((item) => Number.isFinite(Number(item.actualSurgeryMin)));
  const avgDuration = completed.length ? completed.reduce((sum, item) => sum + Number(item.actualSurgeryMin), 0) / completed.length : 0;
  let underCount = 0;
  document.querySelector("#surgeryTotal").textContent = records.length;
  document.querySelector("#surgeryDelayed").textContent = records.filter((item) => Number(item.entryDelayMin) > 15).length;
  document.querySelector("#surgeryAvgDuration").textContent = `${avgDuration.toLocaleString("es-CL", { maximumFractionDigits: 1 })} min`;
  const rows = records.map((item) => {
    const programmed = programmedTimes[item.id] ?? "";
    const difference = programmed === "" || item.actualSurgeryMin === "" ? "" : Number(item.actualSurgeryMin) - Number(programmed);
    const asertividad = classifyAsertividad(difference);
    if (asertividad.label === "Subestimacion") underCount += 1;
    return `<tr>
      <td>${item.patient}</td>
      <td>${item.age}</td>
      <td>${item.surgery}</td>
      <td>${item.code1 || ""}</td>
      <td>${item.code2 || ""}</td>
      <td>${item.category || ""}</td>
      <td>${item.specialty}</td>
      <td>${item.anesthetist || ""}</td>
      <td>${item.surgeon}</td>
      <td>${item.entryTime || ""}</td>
      <td>${item.surgeryStart || ""}</td>
      <td>${item.surgeryEnd || ""}</td>
      <td>${item.exitTime || ""}</td>
      <td>${item.actualSurgeryMin || ""}</td>
      <td><input class="programmed-input" data-surgery-id="${item.id}" type="number" min="0" step="1" value="${programmed}" aria-label="Tiempo programado ${item.patient}" /></td>
      <td>${difference === "" ? "" : difference}</td>
      <td><span class="asertividad ${asertividad.cls}">${asertividad.label}</span></td>
    </tr>`;
  });
  document.querySelector("#surgeryUnder").textContent = underCount;
  body.innerHTML = rows.join("");
  renderRoomSummary(records);
  renderLeadershipSummary(records);
  renderSurgeryDonuts(records);
}

function renderSurgeryDonuts(records) {
  renderDonutChart({
    canvas: document.querySelector("#categoryDonut"),
    legend: document.querySelector("#categoryLegend"),
    data: tallyBy(records, (item) => item.category || "Sin categoria"),
    colors: ["#16639f", "#e8743b", "#177e55", "#9ca3af"],
  });
  renderDonutChart({
    canvas: document.querySelector("#specialtyDonut"),
    legend: document.querySelector("#specialtyLegend"),
    data: tallyBy(records, (item) => item.specialty || "Sin especialidad"),
    colors: ["#16639f", "#e8743b", "#177e55", "#7c3aed", "#0f766e", "#b91c1c", "#64748b"],
    maxLegendItems: 7,
  });
}

function renderDonutChart({ canvas, legend, data, colors, maxLegendItems = 6 }) {
  if (!canvas || !legend) return;
  const ctx = canvas.getContext("2d");
  const entries = Object.entries(data)
    .filter(([, value]) => Number(value) > 0)
    .sort((a, b) => b[1] - a[1]);
  const width = canvas.width;
  const height = canvas.height;
  const centerX = width / 2;
  const centerY = height / 2 - 10;
  const radius = Math.min(width, height) * 0.33;
  const innerRadius = radius * 0.58;
  const total = entries.reduce((sum, [, value]) => sum + Number(value), 0);
  ctx.clearRect(0, 0, width, height);
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, width, height);
  if (!total) {
    ctx.fillStyle = "#60768b";
    ctx.font = "14px Inter, sans-serif";
    ctx.textAlign = "center";
    ctx.fillText("Sin datos", centerX, centerY);
    legend.innerHTML = `<span class="legend-empty">Sin registros para este filtro.</span>`;
    return;
  }
  let start = -Math.PI / 2;
  entries.forEach(([label, value], index) => {
    const angle = (Number(value) / total) * Math.PI * 2;
    ctx.beginPath();
    ctx.moveTo(centerX, centerY);
    ctx.arc(centerX, centerY, radius, start, start + angle);
    ctx.closePath();
    ctx.fillStyle = colors[index % colors.length];
    ctx.fill();
    start += angle;
  });
  ctx.globalCompositeOperation = "destination-out";
  ctx.beginPath();
  ctx.arc(centerX, centerY, innerRadius, 0, Math.PI * 2);
  ctx.fill();
  ctx.globalCompositeOperation = "source-over";
  ctx.fillStyle = "#10365a";
  ctx.font = "700 28px Inter, sans-serif";
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.fillText(total, centerX, centerY - 6);
  ctx.fillStyle = "#60768b";
  ctx.font = "12px Inter, sans-serif";
  ctx.fillText("cirugías", centerX, centerY + 18);
  legend.innerHTML = entries
    .slice(0, maxLegendItems)
    .map(([label, value], index) => {
      const share = total ? Number(value) / total : 0;
      return `<div class="legend-row">
        <span class="legend-dot" style="background:${colors[index % colors.length]}"></span>
        <strong>${label}</strong>
        <span>${value} · ${percent(share)}</span>
      </div>`;
    })
    .join("");
}

function renderRoomSummary(records) {
  const body = document.querySelector("#roomSummaryRows");
  if (!body) return;
  const grouped = new Map();
  records.forEach((item) => {
    const key = `${item.date}|${item.room || "Sin pabellon"}`;
    if (!grouped.has(key)) {
      grouped.set(key, {
        date: item.date,
        room: item.room || "Sin pabellon",
        total: 0,
        asertivo: 0,
        subestimacion: 0,
        sobrestimacion: 0,
        pendiente: 0,
        suspensiones: 0,
        causas: [],
      });
    }
    const row = grouped.get(key);
    row.total += 1;
    const programmed = programmedTimes[item.id] ?? "";
    const difference = programmed === "" || item.actualSurgeryMin === "" ? "" : Number(item.actualSurgeryMin) - Number(programmed);
    const asertividad = classifyAsertividad(difference).label;
    if (asertividad === "Subestimacion") row.subestimacion += 1;
    if (asertividad === "Sobrestimacion") row.sobrestimacion += 1;
    if (asertividad === "Asertivo") row.asertivo += 1;
    if (asertividad === "Pendiente") row.pendiente += 1;
    if ((item.status || "").toLowerCase() === "suspendido") {
      row.suspensiones += 1;
      row.causas.push(item.surgery || "Sin causa registrada");
    }
  });
  const rows = [...grouped.values()].sort((a, b) => a.date.localeCompare(b.date) || a.room.localeCompare(b.room));
  body.innerHTML = rows
    .map((row) => `<tr>
      <td>${formatDate(row.date)}</td>
      <td><strong>${row.room}</strong></td>
      <td>${row.total}</td>
      <td><span class="count-badge ok">${row.asertivo}</span></td>
      <td><span class="count-badge sub">${row.subestimacion}</span></td>
      <td><span class="count-badge sobre">${row.sobrestimacion}</span></td>
      <td><span class="count-badge pending">${row.pendiente}</span></td>
      <td><span class="count-badge ${row.suspensiones ? "susp" : "ok"}">${row.suspensiones}</span></td>
      <td>${row.causas.length ? [...new Set(row.causas)].join("; ") : "Sin suspensiones"}</td>
    </tr>`)
    .join("");
}

function getRoomSummary(records) {
  const grouped = new Map();
  records.forEach((item) => {
    const key = `${item.date}|${item.room || "Sin pabellon"}`;
    if (!grouped.has(key)) {
      grouped.set(key, {
        date: item.date,
        room: item.room || "Sin pabellon",
        total: 0,
        asertivo: 0,
        subestimacion: 0,
        sobrestimacion: 0,
        pendiente: 0,
        suspensiones: 0,
      });
    }
    const row = grouped.get(key);
    row.total += 1;
    const programmed = programmedTimes[item.id] ?? "";
    const difference = programmed === "" || item.actualSurgeryMin === "" ? "" : Number(item.actualSurgeryMin) - Number(programmed);
    const asertividad = classifyAsertividad(difference).label;
    if (asertividad === "Subestimacion") row.subestimacion += 1;
    if (asertividad === "Sobrestimacion") row.sobrestimacion += 1;
    if (asertividad === "Asertivo") row.asertivo += 1;
    if (asertividad === "Pendiente") row.pendiente += 1;
    if ((item.status || "").toLowerCase() === "suspendido") row.suspensiones += 1;
  });
  return [...grouped.values()];
}

function renderLeadershipSummary(records) {
  const risk = document.querySelector("#leadRisk");
  const riskDetail = document.querySelector("#leadRiskDetail");
  const room = document.querySelector("#leadRoom");
  const roomDetail = document.querySelector("#leadRoomDetail");
  const pending = document.querySelector("#leadPending");
  if (!risk || !riskDetail || !room || !roomDetail || !pending) return;
  const summary = getRoomSummary(records);
  const pendingCount = summary.reduce((sum, item) => sum + item.pendiente, 0);
  const riskiest = summary
    .map((item) => ({ ...item, score: item.subestimacion * 3 + item.suspensiones * 4 + item.sobrestimacion + item.pendiente * 0.5 }))
    .sort((a, b) => b.score - a.score)[0];
  pending.textContent = pendingCount;
  if (!riskiest) {
    risk.textContent = "Sin datos";
    riskDetail.textContent = "Actualice desde Excel o revise filtros.";
    room.textContent = "Sin datos";
    roomDetail.textContent = "No hay pabellones para priorizar.";
    return;
  }
  risk.textContent = riskiest.suspensiones ? "Suspensiones" : riskiest.subestimacion ? "Subestimación" : riskiest.pendiente ? "Datos incompletos" : "Sin alerta mayor";
  riskDetail.textContent = `${riskiest.room}: ${riskiest.subestimacion} subestimaciones, ${riskiest.suspensiones} suspensiones.`;
  room.textContent = riskiest.room;
  roomDetail.textContent = `Prioridad por score operativo ${riskiest.score.toLocaleString("es-CL", { maximumFractionDigits: 1 })}.`;
}

function normalizeExcelDate(value) {
  if (!value) return "";
  if (typeof value === "number") {
    const date = new Date(Date.UTC(1899, 11, 30 + value));
    return date.toISOString().slice(0, 10);
  }
  const text = String(value).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;
  const match = text.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/);
  if (match) {
    const [, day, month, year] = match;
    return `${year}-${month.padStart(2, "0")}-${day.padStart(2, "0")}`;
  }
  return text;
}

function normalizeExcelTime(value) {
  if (value === null || value === undefined || value === "") return "";
  if (typeof value === "number") {
    const totalMinutes = Math.round(value * 24 * 60);
    const hours = Math.floor(totalMinutes / 60) % 24;
    const minutes = totalMinutes % 60;
    return `${String(hours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}`;
  }
  return String(value).trim();
}

function minutesBetween(start, end) {
  const startMinutes = timeToMinutes(start);
  const endMinutes = timeToMinutes(end);
  if (startMinutes === null || endMinutes === null) return "";
  const raw = endMinutes - startMinutes;
  return raw >= 0 ? raw : raw + 1440;
}

async function loadSurgeryFromExcel() {
  const button = document.querySelector("#refreshExcelData");
  if (!window.XLSX) {
    alert("No se pudo cargar el lector de Excel. Revise conexión a internet o exporte como CSV.");
    return;
  }
  if (button) button.textContent = "Actualizando...";
  try {
    const response = await fetch(`${EXCEL_TEMPLATE_URL}?t=${Date.now()}`);
    if (!response.ok) throw new Error("No se pudo leer el archivo Excel");
    const buffer = await response.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array", cellDates: false });
    const sheet = workbook.Sheets["Carga Web Cirugias"] || workbook.Sheets[workbook.SheetNames[0]];
    const catalogSheet = workbook.Sheets["Catalogos"];
    const codeCategory = {};
    if (catalogSheet) {
      XLSX.utils.sheet_to_json(catalogSheet, { defval: "", range: 2 }).forEach((row) => {
        const code = String(row["Codigo"] || "").trim();
        const category = String(row["Categoria"] || "").trim();
        if (code && category) codeCategory[code] = category;
      });
    }
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "", range: 2 });
    surgeryRecords = rows
      .filter((row) => row["Fecha Quirofano"] || row["Paciente"] || row["Episodio"])
      .map((row, index) => {
        const id = String(row["Episodio"] || `excel-${index + 1}`).trim();
        const entryTime = normalizeExcelTime(row["Ingreso Pabellon"]);
        const surgeryStart = normalizeExcelTime(row["Inicio Cirugia"]);
        const surgeryEnd = normalizeExcelTime(row["Termino Cirugia"]);
        const exitTime = normalizeExcelTime(row["Salida Pabellon"]);
        const actualSurgeryMin = Number(row["Duracion Cirugia Formula Min"]) || minutesBetween(surgeryStart, surgeryEnd);
        const programmed = row["Tiempo programado de cirugía en tabla (Min)"];
        if (programmed !== "" && programmed !== null && programmed !== undefined) {
          programmedTimes[id] = Number(programmed);
        }
        const code1 = String(row["Codigo 1"] || row["Código 1"] || "").trim();
        const code2 = String(row["Codigo 2"] || row["Código 2"] || "").trim();
        const firstName = String(row["Paciente"] || "").trim();
        const lastName = String(row["Apellido"] || "").trim();
        return {
          id,
          date: normalizeExcelDate(row["Fecha Quirofano"]),
          patient: [firstName, lastName].filter(Boolean).join(" "),
          lastName,
          age: String(row["Edad"] || "").trim(),
          diagnosis: String(row["Diagnostico"] || "").trim(),
          surgery: String(row["Descripción Intervención"] || row["Cirugia"] || row["Causa Suspension"] || "").trim(),
          code1,
          code2,
          category: String(row["Categoria"] || codeCategory[code1] || codeCategory[code2] || "").trim(),
          specialty: String(row["Especialidad"] || "Sin especialidad").trim(),
          anesthetist: String(row["Anestesista"] || "").trim(),
          surgeon: String(row["Cirujano"] || "").trim(),
          room: String(row["Pabellon/Quirofano"] || "Sin pabellon").trim(),
          type: String(row["Tipo"] || "").trim(),
          status: String(row["Estado"] || "").trim(),
          entryTime,
          surgeryStart,
          surgeryEnd,
          exitTime,
          entryDelayMin: Number(row["Demora Ingreso-Inicio Min"]) || minutesBetween(entryTime, surgeryStart) || 0,
          actualSurgeryMin,
          roomDurationMin: Number(row["Duracion Quirofano Min"]) || "",
        };
      });
    saveProgrammedTimes();
    saveSurgeryRecords();
    initSurgeryControls();
    renderSurgeryAnalysis();
  } catch (error) {
    alert(`No se pudo actualizar desde Excel: ${error.message}`);
  } finally {
    if (button) button.textContent = "Actualizar desde Excel";
  }
}

function renderAll() {
  renderKpis();
  renderDailyList();
  renderNodeMap();
  renderSuspensions();
  renderRows();
  renderDelayChart();
  renderSurgeryAnalysis();
}

eventForm.addEventListener("submit", (submitEvent) => {
  submitEvent.preventDefault();
  const form = new FormData(eventForm);
  const item = event(
    form.get("date"),
    form.get("type"),
    form.get("specialty").trim(),
    form.get("room").trim(),
    form.get("node"),
    Number(form.get("minutes")) || 0,
    Number(form.get("count")) || 1,
    form.get("cause").trim(),
  );
  state.events.push(item);
  const day = state.daily.find((entry) => entry.date === item.date);
  if (!day) {
    state.daily.push({
      date: item.date,
      programmed: item.type === "Suspensión" ? item.count : 0,
      operated: 0,
      suspended: item.type === "Suspensión" ? item.count : 0,
      firstDelay: item.type.includes("Demora") ? item.minutes : 0,
      occupancy: 0,
      turnover: item.type.includes("Recambio") ? item.minutes : 0,
      note: item.cause,
    });
  } else {
    if (item.type === "Suspensión") day.suspended += item.count;
    if (item.type.includes("Demora") && item.minutes > day.firstDelay) day.firstDelay = item.minutes;
    if (item.type.includes("Recambio") && item.minutes > day.turnover) day.turnover = item.minutes;
    day.note = item.cause;
  }
  state.daily.sort((a, b) => a.date.localeCompare(b.date));
  saveState();
  initControls();
  dateFilter.value = item.date;
  nodeFilter.value = "todos";
  renderAll();
  eventForm.reset();
  eventForm.elements.date.value = item.date;
});

document.querySelector("#resetData").addEventListener("click", () => {
  state = { daily: initialDaily, events: initialEvents };
  saveState();
  initControls();
  renderAll();
});

document.querySelector("#exportCsv").addEventListener("click", () => {
  const rows = [["Fecha", "Tipo", "Especialidad", "Pabellon", "Nodo", "Minutos", "Cantidad", "Causa"]];
  filteredEvents().forEach((item) => rows.push([item.date, item.type, item.specialty, item.room, item.node, item.minutes, item.count, item.cause]));
  const csv = rows.map((row) => row.map((cell) => `"${String(cell).replaceAll('"', '""')}"`).join(",")).join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "crr_eventos_filtrados.csv";
  link.click();
  URL.revokeObjectURL(url);
});

dateFilter.addEventListener("change", renderAll);
nodeFilter.addEventListener("change", renderAll);
document.querySelector("#surgeryDateFilter")?.addEventListener("change", renderAll);
document.querySelector("#surgerySpecialtyFilter")?.addEventListener("change", renderAll);
document.querySelector("#surgeryStatusFilter")?.addEventListener("change", renderAll);
document.querySelector("#surgeryCategoryFilter")?.addEventListener("change", renderAll);
eventTypeFilter?.addEventListener("change", renderAll);
document.querySelector("#refreshExcelData")?.addEventListener("click", loadSurgeryFromExcel);
document.querySelector("#surgeryRows")?.addEventListener("change", (inputEvent) => {
  if (!inputEvent.target.matches(".programmed-input")) return;
  const id = inputEvent.target.dataset.surgeryId;
  const value = inputEvent.target.value;
  if (value === "") {
    delete programmedTimes[id];
  } else {
    programmedTimes[id] = Number(value);
  }
  saveProgrammedTimes();
  renderSurgeryAnalysis();
});

initControls();
renderAll();
