const baseColumns = [
  {key: "clients_served", label: "Total Clients Served", goal: "15-20", color: "green"},
  {key: "services", label: "Total Services Provided", goal: "45-60", color: "peach"},
  {key: "referrals_to_shelter", label: "Referral to Shelter", goal: "", color: "salmon"},
  {key: "currently_enrolled", label: "Total Currently Enrolled", note: "All enrolled clients during reporting period", goal: "20+", color: "blue"},
  {key: "engaged", label: "Total Engaged", goal: "15-20", color: "gold"},
  {key: "exits", label: "Total Exits (All Clients)", goal: "-", color: "blue"},
  {key: "positive_exits", label: "Positive Exits", goal: "-", color: "blue"},
  {key: "ph_exits", label: "PH Exits", goal: "-", color: "blue"},
  {key: "no_exit_interview", label: "No Exit Interview Exits", goal: "0", color: "blue"},
];

const mhRtOnlyColumns = [
  {key: "chronically_homeless", label: "Chronically Homeless", goal: "", color: "sage"},
  {key: "first_enrollment", label: "First Enrollment in System", goal: "-", color: "sage"},
];

const endingColumns = [
  {key: "cls_assessments", label: "Instances of Service (CLS Assessments)", goal: "45-60", color: "green"},
  {key: "case_notes", label: "Case Notes (Enrollment Level)", goal: "45-60", color: "pink"},
  {key: "average_days", label: "Average # of Days between Service Provided and Case Note Entered", goal: "1-3 days", color: "pink", average: true},
];

const programSelect = document.querySelector("#program-select");
const managerList = document.querySelector("#manager-list");
const managerTemplate = document.querySelector("#manager-template");
const staffTemplate = document.querySelector("#staff-template");
let currentMetrics = {};
let currentServiceUsers = {};
let currentActive = {};
let currentPeriodStaff = new Set();
let currentSharedMetrics = {};
let clsFile = null;
let servicesFile = null;
let enrollmentFile = null;
let caseNotesFile = null;
let teamFile = null;
const pendingAutoProcesses = new Set();

function program() { return programSelect.value; }
function columns() {
  return [...baseColumns, ...(program() === "MHRT" ? mhRtOnlyColumns : []), ...endingColumns];
}

function updateProgramControls() {
  const mockButton = document.querySelector("#load-mock");
  mockButton.hidden = program() !== "MHRT";
}

function resetProgramUploads() {
  clsFile = null;
  servicesFile = null;
  enrollmentFile = null;
  caseNotesFile = null;
  teamFile = null;
  pendingAutoProcesses.clear();
  document.querySelector("#cls-file").value = "";
  document.querySelector("#services-file").value = "";
  document.querySelector("#enrollment-file").value = "";
  document.querySelector("#case-notes-file").value = "";
  document.querySelector("#team-file").value = "";
  document.querySelector("#cls-file-name").textContent = "No file selected";
  document.querySelector("#services-file-name").textContent = "No file selected";
  document.querySelector("#enrollment-file-name").textContent = "No file selected";
  document.querySelector("#case-notes-file-name").textContent = "No file selected";
  document.querySelector("#team-file-name").textContent = "No file selected";
  document.querySelector("#cls-drop-zone").classList.remove("has-file");
  document.querySelector("#services-drop-zone").classList.remove("has-file");
  document.querySelector("#enrollment-drop-zone").classList.remove("has-file");
  document.querySelector("#case-notes-drop-zone").classList.remove("has-file");
  document.querySelector("#team-drop-zone").classList.remove("has-file");
  showStatus("#upload-status", "");
}

function updateDownloadVisibility() {
  const list = document.querySelector("#download-list");
  list.querySelectorAll(".download-link").forEach(link => {
    link.hidden = link.dataset.program !== program();
  });
  const hasVisibleDownload = [...list.querySelectorAll(".download-link")].some(link => !link.hidden);
  let empty = list.querySelector(".empty-download");
  if (!hasVisibleDownload && !empty) {
    empty = document.createElement("span");
    empty.className = "empty-download";
    empty.textContent = `No ${program()} downloads created in this session.`;
    list.append(empty);
  } else if (hasVisibleDownload) {
    empty?.remove();
  } else if (empty) {
    empty.textContent = `No ${program()} downloads created in this session.`;
  }
}

const statusTimers = new Map();
function showStatus(id, message, isError = false, clearAfter = 3000) {
  const element = document.querySelector(id);
  window.clearTimeout(statusTimers.get(id));
  element.textContent = message;
  element.classList.toggle("error", isError);
  if (clearAfter > 0) {
    statusTimers.set(id, window.setTimeout(() => { element.textContent = ""; }, clearAfter));
  }
}

function addStaff(card, name = "", serviceUser = "", active = true) {
  const row = staffTemplate.content.firstElementChild.cloneNode(true);
  row.querySelector(".staff-name").value = name;
  row.querySelector(".service-user").value = serviceUser;
  row.querySelector(".staff-active").value = active ? "yes" : "no";
  row.querySelector(".staff-name").addEventListener("input", renderTable);
  row.querySelector(".staff-active").addEventListener("change", renderTable);
  row.querySelector(".remove-staff").addEventListener("click", () => { row.remove(); renderTable(); });
  card.querySelector(".staff-list").append(row);
}

function addManager(manager = {name: "", staff: []}) {
  const card = managerTemplate.content.firstElementChild.cloneNode(true);
  card.querySelector(".manager-name").value = manager.name || "";
  card.querySelector(".manager-name").addEventListener("input", renderTable);
  card.querySelector(".add-staff").addEventListener("click", () => addStaff(card));
  card.querySelector(".remove-manager").addEventListener("click", () => { card.remove(); renderTable(); });
  managerList.append(card);
  (manager.staff || []).forEach(name => addStaff(
    card,
    name,
    currentServiceUsers[name] || "",
    currentActive[name] !== false
  ));
}

function collectServiceUsers() {
  const mappings = {};
  document.querySelectorAll(".staff-row").forEach(row => {
    const staff = row.querySelector(".staff-name").value.trim();
    const serviceUser = row.querySelector(".service-user").value.trim();
    if (staff && serviceUser) mappings[staff] = serviceUser;
  });
  return mappings;
}

function collectStructure() {
  const active = {};
  document.querySelectorAll(".staff-row").forEach(row => {
    const staff = row.querySelector(".staff-name").value.trim();
    if (staff) active[staff] = row.querySelector(".staff-active").value === "yes";
  });
  return {
    program: program(),
    service_users: collectServiceUsers(),
    active,
    managers: [...document.querySelectorAll(".manager-card")].map(card => ({
      name: card.querySelector(".manager-name").value.trim(),
      staff: [...card.querySelectorAll(".staff-name")].map(input => input.value.trim()).filter(Boolean)
    }))
  };
}

function captureMetrics() {
  document.querySelectorAll("#report-body tr.staff-data-row").forEach(row => {
    const staff = row.dataset.staff;
    currentMetrics[staff] = {};
    row.querySelectorAll("input[data-metric]").forEach(input => {
      const rawValue = input.value.trim();
      const nullable = input.dataset.metric === "average_days";
      currentMetrics[staff][input.dataset.metric] = nullable && (rawValue === "" || rawValue === "-")
        ? null
        : Number(rawValue || 0);
    });
  });
}

function buildHeader() {
  const head = document.querySelector("#report-head");
  head.innerHTML = "";
  const titleRow = document.createElement("tr");
  titleRow.className = "metric-header-row";
  const nameHeader = document.createElement("th");
  nameHeader.textContent = "Outreach Outcomes";
  nameHeader.className = "staff-heading";
  titleRow.append(nameHeader);
  columns().forEach((column, index) => {
    if (index === baseColumns.length + (program() === "MHRT" ? mhRtOnlyColumns.length : 0)) {
      const spacer = document.createElement("th");
      spacer.className = "spacer-column";
      spacer.rowSpan = 2;
      titleRow.append(spacer);
    }
    const th = document.createElement("th");
    th.className = `metric-heading ${column.color}`;
    th.innerHTML = `<span>${column.label}</span>${column.note ? `<small>${column.note}</small>` : ""}`;
    titleRow.append(th);
  });

  const goalRow = document.createElement("tr");
  goalRow.className = "goal-row";
  const goalLabel = document.createElement("th");
  goalLabel.textContent = "CSO Monthly Goals";
  goalRow.append(goalLabel);
  columns().forEach(column => {
    const th = document.createElement("th");
    th.textContent = column.goal;
    goalRow.append(th);
  });
  head.append(titleRow, goalRow);
}

function metricCell(staff, column) {
  const cell = document.createElement("td");
  const input = document.createElement("input");
  const nullable = column.key === "average_days";
  input.type = nullable ? "text" : "number";
  if (nullable) input.inputMode = "decimal";
  else input.min = "0";
  input.step = column.average ? "0.01" : "1";
  input.dataset.metric = column.key;
  const value = currentMetrics[staff]?.[column.key];
  input.value = nullable && (value === null || value === undefined) ? "-" : (value ?? 0);
  if ((currentSharedMetrics[staff] || []).includes(column.key)) {
    cell.classList.add("shared-metric");
    input.title = "Shared assignment counted for multiple staff";
  }
  cell.append(input);
  return cell;
}

function valueFor(staff, key) {
  const value = currentMetrics[staff]?.[key];
  if (value === null || value === undefined || value === "" || value === "-") return null;
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric : null;
}
function aggregate(staffNames, column) {
  const nullable = column.key === "average_days";
  if (!staffNames.length) return nullable ? "-" : 0;
  const values = staffNames.map(staff => valueFor(staff, column.key)).filter(value => value !== null);
  if (!values.length) return nullable ? "-" : 0;
  if (column.average) return (values.reduce((a, b) => a + b, 0) / values.length).toFixed(6).replace(/0+$/, "").replace(/\.$/, "");
  return values.reduce((a, b) => a + b, 0);
}

function addSpacer(row, className = "") {
  const spacer = document.createElement("td");
  spacer.className = `spacer-column ${className}`;
  row.append(spacer);
}

function subtotalRow(label, staffNames, className) {
  const row = document.createElement("tr");
  row.className = className;
  const labelCell = document.createElement("th");
  labelCell.textContent = label;
  row.append(labelCell);
  const activeColumns = columns();
  activeColumns.forEach((column, index) => {
    if (index === activeColumns.length - endingColumns.length) addSpacer(row);
    const cell = document.createElement("td");
    cell.textContent = aggregate(staffNames, column);
    row.append(cell);
  });
  return row;
}

function renderTable() {
  captureMetrics();
  buildHeader();
  const tbody = document.querySelector("#report-body");
  const tfoot = document.querySelector("#report-foot");
  tbody.innerHTML = "";
  tfoot.innerHTML = "";
  const structure = collectStructure();
  const allStaff = [];

  let renderedManagers = 0;
  structure.managers.forEach(manager => {
    const visibleStaff = manager.staff.filter(staff => (
      structure.active[staff] !== false || currentPeriodStaff.has(staff)
    ));
    if (!visibleStaff.length) return;
    if (renderedManagers > 0) {
      const divider = document.createElement("tr");
      divider.className = "manager-divider";
      const cell = document.createElement("td");
      cell.colSpan = columns().length + 2;
      divider.append(cell);
      tbody.append(divider);
    }
    visibleStaff.forEach(staff => {
      const row = document.createElement("tr");
      row.className = "staff-data-row";
      row.dataset.staff = staff;
      row.dataset.manager = manager.name;
      const nameCell = document.createElement("th");
      nameCell.textContent = staff;
      row.append(nameCell);
      columns().forEach((column, index) => {
        if (index === columns().length - endingColumns.length) addSpacer(row);
        row.append(metricCell(staff, column));
      });
      tbody.append(row);
      allStaff.push(staff);
    });
    if (manager.name) {
      tbody.append(subtotalRow(`${manager.name} Total`, visibleStaff, "manager-total"));
    }
    renderedManagers += 1;
  });

  if (allStaff.length) {
    tfoot.append(subtotalRow(`${program()} Total`, allStaff, "program-total"));
    const dateRow = document.createElement("tr");
    dateRow.className = "date-row";
    const label = document.createElement("th");
    label.textContent = "Date range:";
    const range = document.createElement("td");
    range.colSpan = columns().length + 1;
    range.textContent = formatDateRange();
    dateRow.append(label, range);
    tfoot.append(dateRow);
  }
  document.querySelector("#empty-table").hidden = allStaff.length > 0;
  applyFilter();
}

function formatDate(value) {
  if (!value) return "";
  const [year, month, day] = value.split("-");
  return `${month}/${day}/${year}`;
}
function formatDateRange() {
  const start = formatDate(document.querySelector("#start-date").value);
  const end = formatDate(document.querySelector("#end-date").value);
  return start && end ? `${start}-${end}` : "Select start and end dates";
}

function applyFilter() {
  const query = document.querySelector("#staff-filter").value.trim().toLowerCase();
  document.querySelectorAll("#report-body tr.staff-data-row").forEach(row => {
    const searchable = `${row.dataset.staff} ${row.dataset.manager}`.toLowerCase();
    row.hidden = Boolean(query) && !searchable.includes(query);
  });
}

async function loadProgram() {
  const selectedProgram = program();
  managerList.innerHTML = "";
  document.querySelector("#report-body").innerHTML = "";
  updateProgramControls();
  document.querySelector("#table-program").textContent = program();
  const [structureResponse, metricsResponse, serviceUsersResponse] = await Promise.all([
    fetch(`/api/programs/${selectedProgram}/structure`, {cache: "no-store"}),
    fetch(`/api/programs/${selectedProgram}/metrics`, {cache: "no-store"}),
    fetch(`/api/programs/${selectedProgram}/service-users`, {cache: "no-store"})
  ]);
  const structure = await structureResponse.json();
  const metrics = await metricsResponse.json();
  const serviceUsers = await serviceUsersResponse.json();
  if (program() !== selectedProgram) return;
  currentMetrics = metrics;
  currentServiceUsers = serviceUsers;
  currentActive = structure.active || {};
  currentPeriodStaff = new Set();
  currentSharedMetrics = {};
  structure.managers.forEach(addManager);
  renderTable();
  await loadPresetNames();
}

async function loadPresetNames(selectedName = "") {
  const response = await fetch(`/api/programs/${program()}/presets`);
  const result = await response.json();
  const select = document.querySelector("#preset-select");
  select.innerHTML = '<option value="">Choose saved preset</option>';
  (result.presets || []).forEach(name => {
    const option = document.createElement("option");
    option.value = name;
    option.textContent = name;
    option.selected = name === selectedName;
    select.append(option);
  });
}

async function savePreset() {
  const name = document.querySelector("#preset-name").value.trim();
  const response = await fetch(`/api/programs/${program()}/presets`, {
    method: "POST",
    headers: {"Content-Type": "application/json"},
    body: JSON.stringify({name, structure: collectStructure()})
  });
  const result = await response.json();
  if (!response.ok) return showStatus("#structure-status", result.error, true);
  currentServiceUsers = result.service_users || currentServiceUsers;
  currentActive = result.structure.active || collectStructure().active;
  await loadPresetNames(result.name);
  showStatus("#structure-status", `Preset “${result.name}” saved.`);
}

async function loadPreset() {
  const name = document.querySelector("#preset-select").value;
  const response = await fetch(`/api/programs/${program()}/presets/load`, {
    method: "POST",
    headers: {"Content-Type": "application/json"},
    body: JSON.stringify({name})
  });
  const result = await response.json();
  if (!response.ok) return showStatus("#structure-status", result.error, true);
  managerList.innerHTML = "";
  currentServiceUsers = result.service_users || {};
  currentActive = result.structure.active || {};
  currentPeriodStaff = new Set();
  result.structure.managers.forEach(addManager);
  renderTable();
  showStatus("#structure-status", `Preset “${result.name}” loaded.`);
}

function datesAreReady() {
  return Boolean(
    document.querySelector("#start-date").value
    && document.querySelector("#end-date").value
  );
}

async function autoProcessWhenReady(key, processor, label) {
  if (!datesAreReady()) {
    pendingAutoProcesses.add(key);
    return showStatus(
      "#upload-status",
      `${label} selected. Choose the start and end dates to process it automatically.`,
      false,
      0
    );
  }
  pendingAutoProcesses.delete(key);
  await processor();
}

async function processPendingUploads() {
  if (!datesAreReady() || !pendingAutoProcesses.size) return;
  const processors = [
    ["cls", processCls],
    ["services", processServices],
    ["enrollment", processEnrollment],
    ["timeliness", processTimeliness],
  ];
  for (const [key, processor] of processors) {
    if (!pendingAutoProcesses.has(key)) continue;
    pendingAutoProcesses.delete(key);
    await processor();
  }
}

async function chooseClsFile(file) {
  if (!file) return;
  const extension = file.name.split(".").pop().toLowerCase();
  if (!["csv", "xlsx", "xls"].includes(extension)) {
    clsFile = null;
    return showStatus("#upload-status", "Choose a CSV, XLSX, or XLS file.", true);
  }
  clsFile = file;
  document.querySelector("#cls-file-name").textContent = file.name;
  document.querySelector("#cls-drop-zone").classList.add("has-file");
  await autoProcessWhenReady("cls", processCls, "CLS report");
}

async function chooseServicesFile(file) {
  if (!file) return;
  const extension = file.name.split(".").pop().toLowerCase();
  if (!["csv", "xlsx", "xls"].includes(extension)) {
    servicesFile = null;
    return showStatus("#upload-status", "Choose a Services CSV, XLSX, or XLS file.", true);
  }
  servicesFile = file;
  document.querySelector("#services-file-name").textContent = file.name;
  document.querySelector("#services-drop-zone").classList.add("has-file");
  await autoProcessWhenReady("services", processServices, "Services report");
}

async function chooseEnrollmentFile(file) {
  if (!file) return;
  const extension = file.name.split(".").pop().toLowerCase();
  if (!["csv", "xlsx", "xls"].includes(extension)) {
    enrollmentFile = null;
    return showStatus("#upload-status", "Choose an Enrollment CSV, XLSX, or XLS file.", true);
  }
  enrollmentFile = file;
  document.querySelector("#enrollment-file-name").textContent = file.name;
  document.querySelector("#enrollment-drop-zone").classList.add("has-file");
  await autoProcessWhenReady("enrollment", processEnrollment, "Enrollment/Outcomes report");
}

async function chooseCaseNotesFile(file) {
  if (!file) return;
  const extension = file.name.split(".").pop().toLowerCase();
  if (!["csv", "xlsx", "xls"].includes(extension)) {
    caseNotesFile = null;
    return showStatus("#upload-status", "Choose a Case Notes CSV, XLSX, or XLS file.", true);
  }
  caseNotesFile = file;
  document.querySelector("#case-notes-file-name").textContent = file.name;
  document.querySelector("#case-notes-drop-zone").classList.add("has-file");
  await autoProcessWhenReady("timeliness", processTimeliness, "Case Notes/Timeliness report");
}

async function chooseTeamFile(file) {
  if (!file) return;
  const extension = file.name.split(".").pop().toLowerCase();
  if (!["csv", "xlsx", "xls"].includes(extension)) {
    teamFile = null;
    return showStatus("#upload-status", "Choose a Team CSV, XLSX, or XLS file.", true);
  }
  teamFile = file;
  const selectedProgram = program();
  document.querySelector("#team-file-name").textContent = file.name;
  document.querySelector("#team-drop-zone").classList.add("has-file");
  const form = new FormData();
  form.append("team_file", teamFile);
  showStatus("#upload-status", "Loading Team mapping into Step 1…");
  const response = await fetch(`/api/programs/${selectedProgram}/team/import`, {
    method: "POST",
    body: form,
  });
  const result = await response.json();
  if (program() !== selectedProgram) return;
  if (!response.ok) {
    teamFile = null;
    document.querySelector("#team-drop-zone").classList.remove("has-file");
    return showStatus("#upload-status", result.error, true);
  }
  await loadProgram();
  showStatus(
    "#structure-status",
    `${result.staff_count} staff loaded into ${result.manager_count} manager groups from the Team mapping.`
  );
}

async function processCls() {
  if (!clsFile) return showStatus("#upload-status", "Drop a CLS report first.", true);
  const startDate = document.querySelector("#start-date").value;
  const endDate = document.querySelector("#end-date").value;
  if (!startDate || !endDate) {
    return showStatus("#upload-status", "Select the table start and end dates first.", true);
  }
  const form = new FormData();
  const selectedProgram = program();
  form.append("cls_file", clsFile);
  form.append("start_date", startDate);
  form.append("end_date", endDate);
  showStatus("#upload-status", "Processing CLS report…", false, 0);
  let response;
  try {
    response = await fetch(`/api/programs/${selectedProgram}/cls/process`, {method: "POST", body: form});
  } catch (error) {
    return showStatus(
      "#upload-status",
      "The Flask server stopped responding. Keep the server Terminal open and reload this page.",
      true,
      0
    );
  }
  const result = await response.json();
  if (program() !== selectedProgram) return;
  if (!response.ok) return showStatus("#upload-status", result.error, true);
  managerList.innerHTML = "";
  document.querySelector("#report-body").innerHTML = "";
  document.querySelector("#report-foot").innerHTML = "";
  currentMetrics = result.metrics || {};
  currentServiceUsers = result.service_users || currentServiceUsers;
  currentActive = result.structure.active || {};
  (result.report_staff || []).forEach(staff => currentPeriodStaff.add(staff));
  result.structure.managers.forEach(addManager);
  renderTable();
  await loadPresetNames();
  addDownload(result.download_name, result.download_url);
  showStatus(
    "#upload-status",
    `${result.rows_in_date_range} CLS rows processed. Updated Total Clients Served and CLS Assessments; team assignments were unchanged.`
      + (result.unmatched_staff?.length ? ` ${result.unmatched_staff.length} unmatched staff must be added in Step 1.` : ""),
    false,
    10000
  );
  document.querySelector(".table-panel").scrollIntoView({behavior: "smooth"});
}

async function processServices() {
  if (!servicesFile) return showStatus("#upload-status", "Drop a Services report first.", true);
  const startDate = document.querySelector("#start-date").value;
  const endDate = document.querySelector("#end-date").value;
  if (!startDate || !endDate) {
    return showStatus("#upload-status", "Select the table start and end dates first.", true);
  }
  const selectedProgram = program();
  const form = new FormData();
  form.append("services_file", servicesFile);
  form.append("start_date", startDate);
  form.append("end_date", endDate);
  showStatus("#upload-status", "Processing Services report…", false, 0);
  let response;
  try {
    response = await fetch(`/api/programs/${selectedProgram}/services/process`, {
      method: "POST",
      body: form,
    });
  } catch (error) {
    return showStatus(
      "#upload-status",
      "The Flask server stopped responding. Keep the server Terminal open and reload this page.",
      true,
      0
    );
  }
  const result = await response.json();
  if (program() !== selectedProgram) return;
  if (!response.ok) return showStatus("#upload-status", result.error, true, 0);
  managerList.innerHTML = "";
  document.querySelector("#report-body").innerHTML = "";
  document.querySelector("#report-foot").innerHTML = "";
  currentMetrics = result.metrics || {};
  currentServiceUsers = result.service_users || currentServiceUsers;
  currentActive = result.structure.active || {};
  (result.report_staff || []).forEach(staff => currentPeriodStaff.add(staff));
  result.structure.managers.forEach(addManager);
  renderTable();
  addDownload(result.download_name, result.download_url);
  showStatus(
    "#upload-status",
    `${result.rows_in_date_range} deduplicated services processed; ${result.shelter_referrals} shelter referrals. Team assignments were unchanged.`
      + (result.unmatched_service_users?.length
        ? ` ${result.unmatched_service_users.length} unmatched username(s) were added under No Team; enter their full names, save Step 1, then process Services again.`
        : " All service usernames matched.")
      + (result.unmatched_staff?.length ? ` ${result.unmatched_staff.length} resolved staff must be added in Step 1.` : ""),
    false,
    8000
  );
  document.querySelector(".table-panel").scrollIntoView({behavior: "smooth"});
}

async function processEnrollment() {
  if (!enrollmentFile) return showStatus("#upload-status", "Drop an Enrollment/Outcomes report first.", true);
  const startDate = document.querySelector("#start-date").value;
  const endDate = document.querySelector("#end-date").value;
  if (!startDate || !endDate) {
    return showStatus("#upload-status", "Select the table start and end dates first.", true);
  }
  const selectedProgram = program();
  const form = new FormData();
  form.append("enrollment_file", enrollmentFile);
  if (teamFile) form.append("team_file", teamFile);
  if (clsFile) form.append("cls_file", clsFile);
  if (servicesFile) form.append("services_file", servicesFile);
  form.append("start_date", startDate);
  form.append("end_date", endDate);
  showStatus("#upload-status", "Processing Enrollment/Outcomes report…", false, 0);
  let response;
  try {
    response = await fetch(`/api/programs/${selectedProgram}/enrollment/process`, {
      method: "POST",
      body: form,
    });
  } catch (error) {
    return showStatus(
      "#upload-status",
      "The Flask server stopped responding. Keep the server Terminal open and reload this page.",
      true,
      0
    );
  }
  const result = await response.json();
  if (program() !== selectedProgram) return;
  if (!response.ok) return showStatus("#upload-status", result.error, true, 0);
  managerList.innerHTML = "";
  document.querySelector("#report-body").innerHTML = "";
  document.querySelector("#report-foot").innerHTML = "";
  currentMetrics = result.metrics || {};
  currentActive = result.structure.active || {};
  currentSharedMetrics = result.shared_metrics || {};
  (result.report_staff || []).forEach(staff => currentPeriodStaff.add(staff));
  result.structure.managers.forEach(addManager);
  renderTable();
  addDownload(result.download_name, result.download_url);
  showStatus(
    "#upload-status",
    `${result.deduplicated_rows} enrollment records processed for ${result.staff_count} staff. Team assignments were unchanged.`
      + (result.shared_assignments ? ` ${result.shared_assignments} shared assignment row(s) are marked purple.` : ""),
    false,
    10000
  );
  document.querySelector(".table-panel").scrollIntoView({behavior: "smooth"});
}

async function processTimeliness() {
  if (!caseNotesFile) return showStatus("#upload-status", "Drop a Case Notes report first.", true);
  const startDate = document.querySelector("#start-date").value;
  const endDate = document.querySelector("#end-date").value;
  if (!startDate || !endDate) {
    return showStatus("#upload-status", "Select the table start and end dates first.", true);
  }
  const selectedProgram = program();
  const form = new FormData();
  form.append("case_notes_file", caseNotesFile);
  form.append("start_date", startDate);
  form.append("end_date", endDate);
  showStatus("#upload-status", "Processing Case Notes and Timeliness…", false, 0);
  let response;
  try {
    response = await fetch(`/api/programs/${selectedProgram}/timeliness/process`, {
      method: "POST",
      body: form,
    });
  } catch (error) {
    return showStatus(
      "#upload-status",
      "The Flask server stopped responding. Keep the server Terminal open and reload this page.",
      true,
      0
    );
  }
  const result = await response.json();
  if (program() !== selectedProgram) return;
  if (!response.ok) return showStatus("#upload-status", result.error, true, 0);
  document.querySelector("#report-body").innerHTML = "";
  document.querySelector("#report-foot").innerHTML = "";
  currentMetrics = result.metrics || {};
  (result.report_staff || []).forEach(staff => currentPeriodStaff.add(staff));
  renderTable();
  addDownload(result.download_name, result.download_url);
  showStatus(
    "#upload-status",
    `${result.deduplicated_rows} Case Notes processed for ${result.staff_count} staff; ${result.duplicates_removed} duplicate(s) removed and ${result.negative_tim_rows} negative TIM value(s) left blank.`
      + (result.unmatched_staff?.length ? ` ${result.unmatched_staff.length} staff did not match the saved Team roster.` : ""),
    false,
    12000
  );
  document.querySelector(".table-panel").scrollIntoView({behavior: "smooth"});
}

async function createCompleteExcel() {
  if (!enrollmentFile) {
    return showStatus("#metrics-status", "Keep the Enrollment/Outcomes file selected first.", true, 8000);
  }
  if (!caseNotesFile) {
    return showStatus("#metrics-status", "Keep the Case Notes/Timeliness file selected first.", true, 8000);
  }
  const startDate = document.querySelector("#start-date").value;
  const endDate = document.querySelector("#end-date").value;
  if (!startDate || !endDate) {
    return showStatus("#metrics-status", "Select the table start and end dates first.", true);
  }
  captureMetrics();
  const selectedProgram = program();
  const form = new FormData();
  form.append("enrollment_file", enrollmentFile);
  form.append("case_notes_file", caseNotesFile);
  if (teamFile) form.append("team_file", teamFile);
  if (clsFile) form.append("cls_file", clsFile);
  if (servicesFile) form.append("services_file", servicesFile);
  form.append("start_date", startDate);
  form.append("end_date", endDate);
  form.append("metrics_json", JSON.stringify(currentMetrics));
  showStatus("#metrics-status", "Creating complete Excel workbook…", false, 0);
  let response;
  try {
    response = await fetch(`/api/programs/${selectedProgram}/finalize`, {
      method: "POST",
      body: form,
    });
  } catch (error) {
    return showStatus(
      "#metrics-status",
      "The Flask server stopped responding. Keep the server Terminal open and reload this page.",
      true,
      0
    );
  }
  const result = await response.json();
  if (program() !== selectedProgram) return;
  if (!response.ok) return showStatus("#metrics-status", result.error, true, 0);
  document.querySelector("#report-body").innerHTML = "";
  document.querySelector("#report-foot").innerHTML = "";
  currentMetrics = result.metrics || currentMetrics;
  currentSharedMetrics = result.shared_metrics || currentSharedMetrics;
  renderTable();
  addDownload(result.download_name, result.download_url);
  const included = Object.entries(result.included_sources || {})
    .filter(([, value]) => value)
    .map(([name]) => name.replace("case_notes", "case notes"));
  showStatus(
    "#metrics-status",
    `Complete Excel created with Weekly Table, all pivots, and selected source tabs: ${included.join(", ")}.`,
    false,
    12000
  );
}

function addDownload(name, url) {
  const list = document.querySelector("#download-list");
  list.querySelector(".empty-download")?.remove();
  const link = document.createElement("a");
  link.href = url;
  link.textContent = name;
  link.className = "download-link";
  link.dataset.program = program();
  list.prepend(link);
  updateDownloadVisibility();
}

async function saveStructure() {
  const payload = collectStructure();
  const response = await fetch(`/api/programs/${program()}/structure`, {
    method: "PUT", headers: {"Content-Type": "application/json"}, body: JSON.stringify(payload)
  });
  const result = await response.json();
  if (!response.ok) return showStatus("#structure-status", result.error, true);
  currentServiceUsers = payload.service_users;
  currentActive = result.active || payload.active;
  managerList.innerHTML = "";
  result.managers.forEach(addManager);
  renderTable();
  showStatus("#structure-status", `${program()} structure saved.`);
}

async function exportTeamMapping() {
  const payload = collectStructure();
  const response = await fetch(`/api/programs/${program()}/team/export`, {
    method: "POST",
    headers: {"Content-Type": "application/json"},
    body: JSON.stringify(payload),
  });
  const result = await response.json();
  if (!response.ok) return showStatus("#structure-status", result.error, true);
  currentServiceUsers = result.service_users || payload.service_users;
  currentActive = result.structure.active || payload.active;
  addDownload(result.download_name, result.download_url);
  showStatus("#structure-status", `${program()} Team mapping saved and added to Downloads.`, false, 8000);
}

async function saveMetrics() {
  captureMetrics();
  renderTable();
  const response = await fetch(`/api/programs/${program()}/metrics`, {
    method: "PUT", headers: {"Content-Type": "application/json"}, body: JSON.stringify(currentMetrics)
  });
  const result = await response.json();
  if (!response.ok) return showStatus("#metrics-status", result.error, true);
  currentMetrics = result;
  showStatus("#metrics-status", `${program()} table values saved.`);
}

async function loadMockData() {
  if (program() !== "MHRT") {
    return showStatus("#structure-status", "Select MHRT before loading the MHRT preview.", true);
  }
  const approved = window.confirm("Load the synthetic MHRT preview? This replaces currently saved MHRT staff and table values.");
  if (!approved) return;
  const response = await fetch("/api/programs/MHRT/mock", {method: "POST"});
  const result = await response.json();
  if (!response.ok) return showStatus("#structure-status", result.error, true);
  await loadProgram();
  const today = new Date();
  const end = today.toISOString().slice(0, 10);
  const startDate = new Date(today);
  startDate.setDate(today.getDate() - 6);
  document.querySelector("#start-date").value = startDate.toISOString().slice(0, 10);
  document.querySelector("#end-date").value = end;
  renderTable();
  document.querySelector(".table-panel").scrollIntoView({behavior: "smooth"});
  showStatus("#structure-status", "Synthetic MHRT preview loaded.");
}

document.querySelector("#add-manager").addEventListener("click", () => addManager());
document.querySelector("#load-mock").addEventListener("click", loadMockData);
document.querySelector("#save-structure").addEventListener("click", saveStructure);
document.querySelector("#export-team").addEventListener("click", exportTeamMapping);
document.querySelector("#save-metrics").addEventListener("click", saveMetrics);
document.querySelector("#create-complete-excel").addEventListener("click", createCompleteExcel);
document.querySelector("#save-preset").addEventListener("click", savePreset);
document.querySelector("#load-preset").addEventListener("click", loadPreset);
const clsInput = document.querySelector("#cls-file");
const clsDropZone = document.querySelector("#cls-drop-zone");
clsDropZone.addEventListener("click", () => clsInput.click());
clsDropZone.addEventListener("keydown", event => {
  if (event.key === "Enter" || event.key === " ") clsInput.click();
});
clsInput.addEventListener("change", () => chooseClsFile(clsInput.files[0]));
clsDropZone.addEventListener("dragover", event => {
  event.preventDefault();
  clsDropZone.classList.add("dragging");
});
clsDropZone.addEventListener("dragleave", () => clsDropZone.classList.remove("dragging"));
clsDropZone.addEventListener("drop", event => {
  event.preventDefault();
  clsDropZone.classList.remove("dragging");
  chooseClsFile(event.dataTransfer.files[0]);
});
const servicesInput = document.querySelector("#services-file");
const servicesDropZone = document.querySelector("#services-drop-zone");
servicesDropZone.addEventListener("click", () => servicesInput.click());
servicesDropZone.addEventListener("keydown", event => {
  if (event.key === "Enter" || event.key === " ") servicesInput.click();
});
servicesInput.addEventListener("change", () => chooseServicesFile(servicesInput.files[0]));
servicesDropZone.addEventListener("dragover", event => {
  event.preventDefault();
  servicesDropZone.classList.add("dragging");
});
servicesDropZone.addEventListener("dragleave", () => servicesDropZone.classList.remove("dragging"));
servicesDropZone.addEventListener("drop", event => {
  event.preventDefault();
  servicesDropZone.classList.remove("dragging");
  chooseServicesFile(event.dataTransfer.files[0]);
});
const enrollmentInput = document.querySelector("#enrollment-file");
const enrollmentDropZone = document.querySelector("#enrollment-drop-zone");
enrollmentDropZone.addEventListener("click", () => enrollmentInput.click());
enrollmentDropZone.addEventListener("keydown", event => {
  if (event.key === "Enter" || event.key === " ") enrollmentInput.click();
});
enrollmentInput.addEventListener("change", () => chooseEnrollmentFile(enrollmentInput.files[0]));
enrollmentDropZone.addEventListener("dragover", event => {
  event.preventDefault();
  enrollmentDropZone.classList.add("dragging");
});
enrollmentDropZone.addEventListener("dragleave", () => enrollmentDropZone.classList.remove("dragging"));
enrollmentDropZone.addEventListener("drop", event => {
  event.preventDefault();
  enrollmentDropZone.classList.remove("dragging");
  chooseEnrollmentFile(event.dataTransfer.files[0]);
});
const caseNotesInput = document.querySelector("#case-notes-file");
const caseNotesDropZone = document.querySelector("#case-notes-drop-zone");
caseNotesDropZone.addEventListener("click", () => caseNotesInput.click());
caseNotesDropZone.addEventListener("keydown", event => {
  if (event.key === "Enter" || event.key === " ") caseNotesInput.click();
});
caseNotesInput.addEventListener("change", () => chooseCaseNotesFile(caseNotesInput.files[0]));
caseNotesDropZone.addEventListener("dragover", event => {
  event.preventDefault();
  caseNotesDropZone.classList.add("dragging");
});
caseNotesDropZone.addEventListener("dragleave", () => caseNotesDropZone.classList.remove("dragging"));
caseNotesDropZone.addEventListener("drop", event => {
  event.preventDefault();
  caseNotesDropZone.classList.remove("dragging");
  chooseCaseNotesFile(event.dataTransfer.files[0]);
});
const teamInput = document.querySelector("#team-file");
const teamDropZone = document.querySelector("#team-drop-zone");
teamDropZone.addEventListener("click", () => teamInput.click());
teamDropZone.addEventListener("keydown", event => {
  if (event.key === "Enter" || event.key === " ") teamInput.click();
});
teamInput.addEventListener("change", () => chooseTeamFile(teamInput.files[0]));
teamDropZone.addEventListener("dragover", event => {
  event.preventDefault();
  teamDropZone.classList.add("dragging");
});
teamDropZone.addEventListener("dragleave", () => teamDropZone.classList.remove("dragging"));
teamDropZone.addEventListener("drop", event => {
  event.preventDefault();
  teamDropZone.classList.remove("dragging");
  chooseTeamFile(event.dataTransfer.files[0]);
});
document.querySelector("#staff-filter").addEventListener("input", applyFilter);
document.querySelector("#start-date").addEventListener("change", () => {
  renderTable();
  processPendingUploads();
});
document.querySelector("#end-date").addEventListener("change", () => {
  renderTable();
  processPendingUploads();
});
programSelect.addEventListener("change", async () => {
  resetProgramUploads();
  updateDownloadVisibility();
  await loadProgram();
});
updateProgramControls();
loadProgram();
