const EXCEL_URL = "./outputs/crr_dashboard/plantilla_carga_web_crr.xlsx";
const colors = ["#16639f", "#e8743b", "#177e55", "#7c3aed", "#0f766e", "#b91c1c", "#64748b"];

let records = [];

const els = {
  day: document.querySelector("#dayFilter"),
  room: document.querySelector("#roomFilter"),
  status: document.querySelector("#statusFilter"),
  reload: document.querySelector("#reloadExcel"),
};

function clean(value) {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

function normalizeDate(value) {
  if (!value) return "";
  if (typeof value === "number") {
    const date = new Date(Date.UTC(1899, 11, 30 + value));
    return date.toISOString().slice(0, 10);
  }
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
    const hours = Math.floor(total / 60) % 24;
    const minutes = total % 60;
    return `${String(hours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}`;
  }
  const date = value instanceof Date ? value : null;
  if (date) return `${String(date.getHours()).padStart(2, "0")}:${String(date.getMinutes()).padStart(2, "0")}`;
  return clean(value);
}

function formatDate(value) {
  if (!value || value === "todos") return "Todos";
  const [year, month, day] = value.split("-");
  return `${day}-${month}-${year}`;
}

function percent(value) {
  return `${(value * 100).toLocaleString("es-CL", { maximumFractionDigits: 1 })}%`;
}

function isSuspended(record) {
  return /suspend/i.test(record.status);
}

function isDone(record) {
  return /atendido|realizado/i.test(record.status);
}

function isLate(record) {
  return /atraso/i.test(record.entryClass);
}

function tally(items, keyFn) {
  return items.reduce((acc, item) => {
    const key = keyFn(item) || "Sin dato";
    acc[key] = (acc[key] || 0) + 1;
    return acc;
  }, {});
}

async function loadExcel() {
  if (!window.XLSX) {
    alert("No se pudo cargar el lector de Excel.");
    return;
  }
  els.reload.textContent = "Actualizando...";
  const response = await fetch(`${EXCEL_URL}?t=${Date.now()}`);
  const buffer = await response.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  const sheet = workbook.Sheets["Carga Web Cirugias"] || workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "", range: 2 });
  records = rows
    .filter((row) => row["Fecha Quirofano"] || row["Paciente"] || row["Episodio"])
    .map((row, index) => {
      const firstName = clean(row["Paciente"]);
      const lastName = clean(row["Apellido"]);
      return {
        id: clean(row["Episodio"]) || `row-${index}`,
        date: normalizeDate(row["Fecha Quirofano"]),
        patient: [firstName, lastName].filter(Boolean).join(" "),
        room: clean(row["Pabellon/Quirofano"]) || "Sin pabellón",
        status: clean(row["Estado"]) || "Sin estado",
        entryTime: normalizeTime(row["Ingreso Pabellon"]),
        entryClass: clean(row["Clasificación ingreso 08:00-08:15"]),
        category: clean(row["Categoria"]) || "Sin categoría",
        specialty: clean(row["Especialidad"]) || "Sin especialidad",
        surgery: clean(row["Descripción Intervención"]) || clean(row["Cirugia"]),
        suspensionCause: clean(row["Causa Suspension"]),
        anesthetist: clean(row["Anestesista"]),
        surgeon: clean(row["Cirujano"]),
      };
    });
  initFilters();
  render();
  els.reload.textContent = "Actualizar Excel";
}

function initFilters() {
  const currentDay = els.day.value || "latest";
  const currentRoom = els.room.value || "todos";
  const currentStatus = els.status.value || "todos";
  const dates = [...new Set(records.map((item) => item.date).filter(Boolean))].sort();
  const rooms = [...new Set(records.map((item) => item.room))].sort();
  const statuses = [...new Set(records.map((item) => item.status))].sort();
  els.day.innerHTML = `<option value="todos">Todos los días</option>${dates.map((date) => `<option value="${date}">${formatDate(date)}</option>`).join("")}`;
  els.room.innerHTML = `<option value="todos">Todos</option>${rooms.map((room) => `<option>${room}</option>`).join("")}`;
  els.status.innerHTML = `<option value="todos">Todos</option>${statuses.map((status) => `<option>${status}</option>`).join("")}`;
  els.day.value = currentDay === "latest" ? dates.at(-1) || "todos" : dates.includes(currentDay) ? currentDay : "todos";
  els.room.value = rooms.includes(currentRoom) ? currentRoom : "todos";
  els.status.value = statuses.includes(currentStatus) ? currentStatus : "todos";
}

function filteredRecords() {
  return records.filter((record) => {
    const dayOk = els.day.value === "todos" || record.date === els.day.value;
    const roomOk = els.room.value === "todos" || record.room === els.room.value;
    const statusOk = els.status.value === "todos" || record.status === els.status.value;
    return dayOk && roomOk && statusOk;
  });
}

function render() {
  const data = filteredRecords();
  const suspended = data.filter(isSuspended);
  const done = data.filter(isDone);
  const late = data.filter(isLate);
  document.querySelector("#totalCases").textContent = data.length;
  document.querySelector("#doneCases").textContent = done.length;
  document.querySelector("#lateCases").textContent = late.length;
  document.querySelector("#suspendedCases").textContent = suspended.length;
  document.querySelector("#suspendedDetail").textContent = data.length ? `${percent(suspended.length / data.length)} del día` : "0% del día";
  renderPriority(data, late, suspended);
  renderRooms(data);
  renderRows(data);
  renderDonut("categoryChart", "categoryLegend", tally(data, (item) => item.category));
  renderDonut("specialtyChart", "specialtyLegend", tally(data, (item) => item.specialty));
}

function renderPriority(data, late, suspended) {
  const box = document.querySelector("#priorityBox");
  const title = document.querySelector("#priorityTitle");
  const text = document.querySelector("#priorityText");
  const actions = document.querySelector("#actionList");
  box.classList.remove("warning", "danger");
  if (!data.length) {
    title.textContent = "Sin registros para el filtro";
    text.textContent = "Seleccione otro día o actualice el Excel.";
    actions.innerHTML = "";
    return;
  }
  if (suspended.length || late.length >= 3) {
    box.classList.add("danger");
    title.textContent = suspended.length ? "Revisar suspensiones del día" : "Atrasos concentrados";
    text.textContent = `${late.length} ingresos con atraso y ${suspended.length} suspensiones registradas.`;
  } else if (late.length) {
    box.classList.add("warning");
    title.textContent = "Jornada en observación";
    text.textContent = `${late.length} ingreso(s) posterior(es) a las 08:15.`;
  } else {
    title.textContent = "Jornada sin alerta mayor";
    text.textContent = "No hay suspensiones ni atrasos de ingreso en el filtro actual.";
  }
  const topRoom = Object.entries(tally(late, (item) => item.room)).sort((a, b) => b[1] - a[1])[0];
  const topCause = Object.entries(tally(suspended, (item) => item.suspensionCause)).sort((a, b) => b[1] - a[1])[0];
  actions.innerHTML = [
    topRoom ? ["Pabellón a mirar", `${topRoom[0]} concentra ${topRoom[1]} atraso(s).`] : ["Ingreso", "Mantener control de primera hora."],
    topCause ? ["Causa suspensión", `${topCause[0]}: ${topCause[1]} caso(s).`] : ["Suspensiones", "Sin causa dominante registrada."],
  ]
    .map(([head, detail]) => `<div class="action-item"><strong>${head}</strong><span>${detail}</span></div>`)
    .join("");
}

function renderRooms(data) {
  const grouped = new Map();
  data.forEach((item) => {
    if (!grouped.has(item.room)) grouped.set(item.room, { total: 0, late: 0, suspended: 0, done: 0 });
    const row = grouped.get(item.room);
    row.total += 1;
    if (isLate(item)) row.late += 1;
    if (isSuspended(item)) row.suspended += 1;
    if (isDone(item)) row.done += 1;
  });
  const rows = [...grouped.entries()].sort((a, b) => a[0].localeCompare(b[0]));
  document.querySelector("#roomBoard").innerHTML = rows.length
    ? rows
        .map(([room, item]) => {
          const cls = item.suspended ? "danger" : item.late ? "warning" : "";
          return `<article class="room-card ${cls}">
            <strong>${room}</strong>
            <span>${item.done}/${item.total} resueltos</span>
            <div class="room-metrics">
              <span class="chip">${item.total} casos</span>
              <span class="chip warning">${item.late} atrasos</span>
              <span class="chip danger">${item.suspended} susp.</span>
            </div>
          </article>`;
        })
        .join("")
    : `<p>Sin pabellones para este filtro.</p>`;
}

function renderRows(data) {
  document.querySelector("#dailyRows").innerHTML = data
    .map((item) => {
      const entryClass = isLate(item) ? "warning" : item.entryClass ? "" : "warning";
      const statusClass = isSuspended(item) ? "danger" : isDone(item) ? "" : "warning";
      return `<tr>
        <td><strong>${item.patient || "Sin paciente"}</strong></td>
        <td>${item.room}</td>
        <td><span class="badge ${statusClass}">${item.status}</span></td>
        <td>${item.entryTime}</td>
        <td><span class="badge ${entryClass}">${item.entryClass || "Pendiente"}</span></td>
        <td>${item.category}</td>
        <td>${item.specialty}</td>
        <td>${item.surgery}</td>
        <td>${item.suspensionCause || "Sin suspensión"}</td>
      </tr>`;
    })
    .join("");
}

function renderDonut(canvasId, legendId, data) {
  const canvas = document.querySelector(`#${canvasId}`);
  const legend = document.querySelector(`#${legendId}`);
  const ctx = canvas.getContext("2d");
  const entries = Object.entries(data).filter(([, value]) => value > 0).sort((a, b) => b[1] - a[1]);
  const total = entries.reduce((sum, [, value]) => sum + value, 0);
  const cx = canvas.width / 2;
  const cy = canvas.height / 2 - 6;
  const radius = 88;
  ctx.clearRect(0, 0, canvas.width, canvas.height);
  ctx.fillStyle = "#fff";
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  if (!total) {
    ctx.fillStyle = "#60768b";
    ctx.textAlign = "center";
    ctx.font = "14px Inter, sans-serif";
    ctx.fillText("Sin datos", cx, cy);
    legend.innerHTML = "";
    return;
  }
  let start = -Math.PI / 2;
  entries.forEach(([, value], index) => {
    const angle = (value / total) * Math.PI * 2;
    ctx.beginPath();
    ctx.moveTo(cx, cy);
    ctx.arc(cx, cy, radius, start, start + angle);
    ctx.closePath();
    ctx.fillStyle = colors[index % colors.length];
    ctx.fill();
    start += angle;
  });
  ctx.globalCompositeOperation = "destination-out";
  ctx.beginPath();
  ctx.arc(cx, cy, 50, 0, Math.PI * 2);
  ctx.fill();
  ctx.globalCompositeOperation = "source-over";
  ctx.fillStyle = "#10365a";
  ctx.font = "800 28px Inter, sans-serif";
  ctx.textAlign = "center";
  ctx.fillText(total, cx, cy - 3);
  ctx.fillStyle = "#60768b";
  ctx.font = "12px Inter, sans-serif";
  ctx.fillText("casos", cx, cy + 20);
  legend.innerHTML = entries
    .slice(0, 6)
    .map(([label, value], index) => `<div class="legend-row">
      <span class="legend-dot" style="background:${colors[index % colors.length]}"></span>
      <strong>${label}</strong>
      <span>${value} · ${percent(value / total)}</span>
    </div>`)
    .join("");
}

els.reload.addEventListener("click", loadExcel);
els.day.addEventListener("change", render);
els.room.addEventListener("change", render);
els.status.addEventListener("change", render);
loadExcel().catch((error) => {
  console.error(error);
  document.querySelector("#priorityTitle").textContent = "No se pudo leer el Excel";
  document.querySelector("#priorityText").textContent = error.message;
});
