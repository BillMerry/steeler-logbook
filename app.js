// --- Constants & state ---------------------------------------------

const STORAGE_KEY = "steeler_logbook_passages_v5";
const THEME_KEY   = "steeler_logbook_theme_v1";
const PORTS_KEY   = "steeler_logbook_ports_v1";

let passages = [];
let currentPassageId = null;
let knownPorts = [];

// --- Storage helpers -----------------------------------------------

function loadPassages() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    passages = raw ? JSON.parse(raw) : [];
  } catch (e) {
    console.error("Failed to load passages", e);
    passages = [];
  }
}

function savePassages() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(passages));
  } catch (e) {
    console.error("Failed to save passages", e);
  }
}

function loadPorts() {
  try {
    const raw = localStorage.getItem(PORTS_KEY);
    knownPorts = raw ? JSON.parse(raw) : [];
  } catch {
    knownPorts = [];
  }
}

function savePorts() {
  try {
    localStorage.setItem(PORTS_KEY, JSON.stringify(knownPorts));
  } catch (e) {
    console.warn("Failed to save ports", e);
  }
}

function rememberPort(name) {
  const trimmed = (name || "").trim();
  if (!trimmed) return;
  if (!knownPorts.includes(trimmed)) {
    knownPorts.push(trimmed);
    knownPorts.sort((a, b) => a.localeCompare(b));
    savePorts();
    renderPortsDatalist();
  }
}

// --- Small helpers -------------------------------------------------

function getCurrentPassage() {
  return passages.find(p => p.id === currentPassageId) || null;
}

function escapeHtml(str) {
  if (str == null) return "";
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function quote(value) {
  if (value == null) return '""';
  const s = String(value).replace(/"/g, '""');
  return `"${s}"`;
}

function timeOnlyFromIso(iso) {
  if (!iso || iso.length < 16) return iso || "";
  return iso.slice(11, 16);
}

function switchToTab(tabId) {
  tabButtons.forEach(b => b.classList.toggle("active", b.dataset.tab === tabId));
  tabs.forEach(t => t.classList.toggle("active", t.id === tabId));
}

// Position formatting helpers: decimal degrees -> dºmm.mmm'H
function formatLatFromDecimal(decimal) {
  if (isNaN(decimal)) return "";
  const hemi = decimal >= 0 ? "N" : "S";
  const dAbs = Math.abs(decimal);
  const deg = Math.floor(dAbs);
  const minutes = (dAbs - deg) * 60;
  const minutesStr = minutes.toFixed(3).padStart(6, "0");
  return `${deg}º${minutesStr}'${hemi}`;
}
function formatLonFromDecimal(decimal) {
  if (isNaN(decimal)) return "";
  const hemi = decimal >= 0 ? "E" : "W";
  const dAbs = Math.abs(decimal);
  const deg = Math.floor(dAbs);
  const minutes = (dAbs - deg) * 60;
  const minutesStr = minutes.toFixed(3).padStart(6, "0");
  return `${deg}º${minutesStr}'${hemi}`;
}
function parseAndFormatPositionInput(val, currentLat, currentLon) {
  if (!val) return { lat: "", lon: "" };

  if (/[º°NnSsEeWw]/.test(val)) {
    const parts = val.split(",").map(s => s.trim());
    return { lat: parts[0] || currentLat || "", lon: parts[1] || currentLon || "" };
  }

  const parts = val.split(",").map(s => s.trim());
  const latNum = parseFloat(parts[0]);
  const lonNum = parseFloat(parts[1]);
  if (isNaN(latNum) || isNaN(lonNum)) return { lat: val, lon: currentLon || "" };

  return { lat: formatLatFromDecimal(latNum), lon: formatLonFromDecimal(lonNum) };
}

function isLocalDestination(val) {
  const s = (val || "").trim().toLowerCase();
  return !s || s === "local";
}

// --- DOM references ------------------------------------------------

const headerPassageMain = document.getElementById("headerPassageMain");
const headerSunrise     = document.getElementById("headerSunrise");
const headerCrew        = document.getElementById("headerCrew");
const themeToggleBtn    = document.getElementById("themeToggleBtn");

const tabButtons = document.querySelectorAll(".tab-btn");
const tabs       = document.querySelectorAll(".tab");

const homeNewPassageBtn = document.getElementById("homeNewPassageBtn");
const homePassageList   = document.getElementById("homePassageList");

const exportBackupBtn = document.getElementById("exportBackupBtn");
const importBackupBtn = document.getElementById("importBackupBtn");
const importFileInput = document.getElementById("importFileInput");

const planForm = document.getElementById("planForm");
const planDate = document.getElementById("planDate");
const planFrom = document.getElementById("planFrom");
const planTo   = document.getElementById("planTo");
const planVessel = document.getElementById("planVessel");
const planSkipper = document.getElementById("planSkipper");
const planCrew = document.getElementById("planCrew");
const planSunriseSet = document.getElementById("planSunriseSet");
const planTidalCoeff = document.getElementById("planTidalCoeff");
const planCurrents = document.getElementById("planCurrents");
const planWeather = document.getElementById("planWeather");
const planComms = document.getElementById("planComms");
const tideStationsContainer = document.getElementById("tideStationsContainer");
const addTideStationBtn = document.getElementById("addTideStationBtn");
const dailySummariesContainer = document.getElementById("dailySummariesContainer");
const addDailySummaryBtn = document.getElementById("addDailySummaryBtn");

const addEntryBtn = document.getElementById("addEntryBtn");
const logEntriesContainer = document.getElementById("logEntriesContainer");
const logEmptyMessage = document.getElementById("logEmptyMessage");
const planSummaryPanel = document.getElementById("planSummaryPanel");
const logLayout = document.getElementById("logLayout");
const splitViewBtn = document.getElementById("splitViewBtn");
const expandPlanBtn = document.getElementById("expandPlanBtn");
const expandLogBtn = document.getElementById("expandLogBtn");
const engineStartBtn = document.getElementById("engineStartBtn");
const slipLinesBtn = document.getElementById("slipLinesBtn");
const dockLinesBtn = document.getElementById("dockLinesBtn");
const shutdownBtn = document.getElementById("shutdownBtn");
const exportCsvBtn = document.getElementById("exportCsvBtn");
const logSummaryPanel = document.getElementById("logSummaryPanel");

const modalOverlay = document.getElementById("modalOverlay");
const modalTitle   = document.getElementById("modalTitle");
const modalBody    = document.getElementById("modalBody");
const modalCancelBtn = document.getElementById("modalCancelBtn");
const modalOkBtn     = document.getElementById("modalOkBtn");

// --- Theme handling -----------------------------------------------

function applyTheme(theme) {
  document.body.dataset.theme = theme;
  localStorage.setItem(THEME_KEY, theme);
  if (themeToggleBtn) themeToggleBtn.textContent = theme === "night" ? "Day" : "Night";
}

themeToggleBtn?.addEventListener("click", () => {
  const current = document.body.dataset.theme || "day";
  applyTheme(current === "night" ? "day" : "night");
});

// --- Tabs ----------------------------------------------------------

tabButtons.forEach(btn => {
  btn.addEventListener("click", () => switchToTab(btn.dataset.tab));
});

// --- Header info ---------------------------------------------------

function updatePassageHeader() {
  const p = getCurrentPassage();
  if (!p) {
    headerPassageMain.textContent = "";
    headerSunrise.textContent = "";
    headerCrew.textContent = "";
    return;
  }

  const date = p.plan.date || p.createdAt.slice(0, 10);
  const from = p.plan.from || "?";
  const to   = p.plan.to   || "?";

  headerPassageMain.textContent = `${date} – ${from} → ${to}`;
  headerSunrise.textContent = p.plan.sunriseSet ? `Sunrise–Set: ${p.plan.sunriseSet}` : "";

  const crewParts = [];
  if (p.plan.skipper) crewParts.push(`Skipper: ${p.plan.skipper}`);
  if (p.plan.crew)    crewParts.push(`Crew: ${p.plan.crew}`);
  headerCrew.textContent = crewParts.join("  |  ");
}

// --- Ports datalist -----------------------------------------------

function renderPortsDatalist() {
  const dl = document.getElementById("portsList");
  if (!dl) return;
  dl.innerHTML = "";
  knownPorts.forEach(p => {
    const opt = document.createElement("option");
    opt.value = p;
    dl.appendChild(opt);
  });
}

// --- Modal ---------------------------------------------------------

function showModal({ title, bodyHtml, onOk }) {
  modalTitle.textContent = title;
  modalBody.innerHTML = bodyHtml;
  modalOverlay.classList.remove("hidden");

  const cleanup = () => {
    modalOverlay.classList.add("hidden");
    modalBody.innerHTML = "";
    modalOkBtn.onclick = null;
    modalCancelBtn.onclick = null;
  };

  modalCancelBtn.onclick = () => cleanup();
  modalOkBtn.onclick = () => {
    const res = onOk?.();
    if (res !== false) cleanup();
  };
}

// --- Backup / Restore ----------------------------------------------

function exportBackup() {
  const payload = {
    format: "steeler-logbook-backup",
    version: 1,
    exportedAt: new Date().toISOString(),
    data: {
      passages,
      knownPorts,
      theme: localStorage.getItem(THEME_KEY) || "day"
    }
  };

  const json = JSON.stringify(payload, null, 2);
  const blob = new Blob([json], { type: "application/json" });
  const url = URL.createObjectURL(blob);

  const d = new Date();
  const y = d.getFullYear();
  const mo = String(d.getMonth() + 1).padStart(2, "0");
  const da = String(d.getDate()).padStart(2, "0");
  const hh = String(d.getHours()).padStart(2, "0");
  const mm = String(d.getMinutes()).padStart(2, "0");
  const filename = `STEELER-Logbook-backup-${y}${mo}${da}${hh}${mm}.json`;

  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function importBackupFile(file) {
  const reader = new FileReader();
  reader.onload = () => {
    try {
      const obj = JSON.parse(reader.result);
      if (!obj || obj.format !== "steeler-logbook-backup" || !obj.data) {
        alert("That file doesn’t look like a STEELER Logbook backup.");
        return;
      }
      if (!Array.isArray(obj.data.passages) || !Array.isArray(obj.data.knownPorts)) {
        alert("Backup file is missing expected data.");
        return;
      }
      const ok = confirm("Restore backup? This will REPLACE the current logbook data on this device.");
      if (!ok) return;

      passages = obj.data.passages;
      knownPorts = obj.data.knownPorts;

      localStorage.setItem(STORAGE_KEY, JSON.stringify(passages));
      localStorage.setItem(PORTS_KEY, JSON.stringify(knownPorts));

      applyTheme(obj.data.theme || "day");

      refreshHomePassageList();
      currentPassageId = passages[0]?.id || null;
      loadPassageIntoUI();
      alert("Backup restored successfully.");
    } catch (e) {
      console.error(e);
      alert("Could not restore that file (invalid JSON).");
    }
  };
  reader.readAsText(file);
}

exportBackupBtn?.addEventListener("click", exportBackup);
importBackupBtn?.addEventListener("click", () => importFileInput?.click());
importFileInput?.addEventListener("change", (e) => {
  const file = e.target.files?.[0];
  if (!file) return;
  importBackupFile(file);
  e.target.value = "";
});

// --- HOME: passage list + delete + swipe ---------------------------

function deletePassageById(id) {
  const idx = passages.findIndex(p => p.id === id);
  if (idx < 0) return;
  const p = passages[idx];
  const label = `${p.plan.date || p.createdAt.slice(0,10)} – ${(p.plan.from||"?")} → ${(p.plan.to||"?")}`;
  const ok = confirm(`Delete this passage?\n\n${label}\n\nThis cannot be undone (unless you’ve got a backup).`);
  if (!ok) return;

  passages.splice(idx, 1);
  savePassages();

  if (currentPassageId === id) currentPassageId = passages[0]?.id || null;

  refreshHomePassageList();
  loadPassageIntoUI();
}

function attachSwipeToCard(card, passageId) {
  let startX = 0;
  card.addEventListener("touchstart", (e) => { startX = e.changedTouches[0].screenX; }, { passive: true });
  card.addEventListener("touchend", (e) => {
    const dx = e.changedTouches[0].screenX - startX;
    if (dx < -90) deletePassageById(passageId);
  }, { passive: true });
}

function refreshHomePassageList() {
  homePassageList.innerHTML = "";

  if (passages.length === 0) {
    const p = document.createElement("p");
    p.textContent = "No passages yet. Tap “+ New Passage” to get started.";
    p.style.opacity = "0.8";
    p.style.fontSize = "0.85rem";
    homePassageList.appendChild(p);
    return;
  }

  passages.forEach(passage => {
    const card = document.createElement("div");
    card.className = "passage-card";

    const date = passage.plan.date || passage.createdAt.slice(0, 10);
    const from = passage.plan.from || "?";
    const to = passage.plan.to || "?";
    const status = passage.finish?.shutdownLogged ? "Completed" : "In progress";
    const entriesCount = passage.entries?.length || 0;

    const left = document.createElement("div");
    left.className = "passage-card-left";
    left.innerHTML = `
      <div class="passage-card-title">${escapeHtml(`${date} – ${from} → ${to}`)}</div>
      <div class="passage-card-meta"><span>${entriesCount} entries</span><span>${status}</span></div>
    `;

    const actions = document.createElement("div");
    actions.className = "passage-card-actions";

    const del = document.createElement("button");
    del.className = "passage-delete-btn";
    del.textContent = "Delete";
    del.addEventListener("click", (e) => { e.stopPropagation(); deletePassageById(passage.id); });

    actions.appendChild(del);
    card.appendChild(left);
    card.appendChild(actions);

    card.addEventListener("click", () => {
      currentPassageId = passage.id;
      loadPassageIntoUI();
      switchToTab("logTab");
    });

    attachSwipeToCard(card, passage.id);
    homePassageList.appendChild(card);
  });
}

// --- Layout mode controls (Log tab) -------------------------------

function setActiveViewButton(btn) {
  document.querySelectorAll(".view-btn").forEach(b => b.classList.remove("active"));
  btn.classList.add("active");
}

function setLogLayoutMode(mode, btn) {
  logLayout.classList.remove("split", "plan-only", "log-only");
  logLayout.classList.add(mode === "plan-only" ? "plan-only" : mode === "log-only" ? "log-only" : "split");
  if (btn) setActiveViewButton(btn);
}

splitViewBtn.addEventListener("click", () => setLogLayoutMode("split", splitViewBtn));
expandPlanBtn.addEventListener("click", () => setLogLayoutMode("plan-only", expandPlanBtn));
expandLogBtn.addEventListener("click", () => setLogLayoutMode("log-only", expandLogBtn));

// --- Plan tab logic -----------------------------------------------

function ensureFlags(p) {
  if (!p.flags) p.flags = { engineStart: false, slip: false, dock: false };
  if (typeof p.flags.engineStart !== "boolean") p.flags.engineStart = false;
  if (typeof p.flags.slip !== "boolean") p.flags.slip = false;
  if (typeof p.flags.dock !== "boolean") p.flags.dock = false;
}

function ensureAutoTideStations(p) {
  if (!p) return;
  if (!p.plan.tideStations) p.plan.tideStations = [];

  const origin = (p.plan.from || "").trim();
  const dest = (p.plan.to || "").trim();

  const want = [];
  if (origin) want.push(origin);
  if (!isLocalDestination(dest) && dest && dest !== origin) want.push(dest);

  // If nothing to prepopulate, do nothing
  if (want.length === 0) return;

  // If no stations at all -> create auto stations for want
  if (p.plan.tideStations.length === 0) {
    p.plan.tideStations = want.map((name, i) => ({
      id: `ts_${Date.now()}_${i}`,
      name,
      hw1: "", hw2: "", lw1: "", lw2: "",
      auto: true
    }));
    return;
  }

  // Update only existing auto stations; never clobber manual stations
  const autos = p.plan.tideStations.filter(ts => ts.auto);
  const manuals = p.plan.tideStations.filter(ts => !ts.auto);

  // If there are no autos, assume user fully manual; do nothing
  if (autos.length === 0) return;

  const newAutos = want.map((name, i) => {
    const prev = autos[i] || {};
    return {
      id: prev.id || `ts_${Date.now()}_${i}`,
      name,
      hw1: prev.hw1 || "",
      hw2: prev.hw2 || "",
      lw1: prev.lw1 || "",
      lw2: prev.lw2 || "",
      auto: true
    };
  });

  p.plan.tideStations = [...newAutos, ...manuals];
}

function createPassage() {
  const id = "p_" + Date.now();
  const today = new Date().toISOString().slice(0, 10);

  const passage = {
    id,
    flags: { engineStart: false, slip: false, dock: false },
    plan: {
      date: today,
      from: "",
      to: "",
      vessel: "STEELER",
      skipper: "",
      crew: "",
      sunriseSet: "",
      tidalCoeff: "",
      tideStations: [],
      currents: "",
      weather: "",
      comms: "",
      engineHoursStart: "",
      fuelStartPercent: "",
      dailySummaries: [
        { id: "ds_" + Date.now(), date: today, fee: "", notes: "" }
      ]
    },
    entries: [],
    finish: {
      engineHoursEnd: "",
      fuelEndPercent: "",
      notes: "",
      shutdownLogged: false
    },
    createdAt: new Date().toISOString()
  };

  passages.unshift(passage);
  currentPassageId = id;
  savePassages();
  refreshHomePassageList();
  loadPassageIntoUI();
}

function loadPlanIntoForm(p) {
  planDate.value = p.plan.date || "";
  planFrom.value = p.plan.from || "";
  planTo.value   = p.plan.to   || "";
  planVessel.value = p.plan.vessel || "STEELER";
  planSkipper.value = p.plan.skipper || "";
  planCrew.value = p.plan.crew || "";
  planSunriseSet.value = p.plan.sunriseSet || "";
  planTidalCoeff.value = p.plan.tidalCoeff || "";
  planCurrents.value = p.plan.currents || "";
  planWeather.value = p.plan.weather || "";
  planComms.value = p.plan.comms || "";

  renderTideStations(p);
  renderDailySummaries(p);
}

function renderTideStations(p) {
  tideStationsContainer.innerHTML = "";
  const stations = p.plan.tideStations || [];
  stations.forEach((st, index) => {
    const row = document.createElement("div");
    row.className = "tide-station-row";
    row.dataset.index = index;
    row.dataset.auto = st.auto ? "true" : "false";
    row.dataset.id = st.id || "";

    row.innerHTML = `
      <div class="row">
        <label>
          Tide station
          <input type="text" class="ts-name" value="${escapeHtml(st.name || "")}" list="portsList">
        </label>
        <button type="button" class="btn btn-secondary btn-small remove-tide-station">Remove</button>
      </div>
      <div class="row">
        <label>HW 1 <input type="time" class="ts-hw1" value="${st.hw1 || ""}"></label>
        <label>HW 2 <input type="time" class="ts-hw2" value="${st.hw2 || ""}"></label>
      </div>
      <div class="row">
        <label>LW 1 <input type="time" class="ts-lw1" value="${st.lw1 || ""}"></label>
        <label>LW 2 <input type="time" class="ts-lw2" value="${st.lw2 || ""}"></label>
      </div>
      <div class="row">
        <button type="button" class="btn btn-secondary btn-small move-up">↑</button>
        <button type="button" class="btn btn-secondary btn-small move-down">↓</button>
      </div>
    `;

    const nameInput = row.querySelector(".ts-name");
    nameInput.addEventListener("input", () => { row.dataset.auto = "false"; });

    row.querySelector(".remove-tide-station").addEventListener("click", () => {
      p.plan.tideStations = readTideStationsFromForm();
      p.plan.tideStations.splice(index, 1);
      renderTideStations(p);
    });

    row.querySelector(".move-up").addEventListener("click", () => moveTideStation(index, -1));
    row.querySelector(".move-down").addEventListener("click", () => moveTideStation(index, 1));

    tideStationsContainer.appendChild(row);
  });
}

function readTideStationsFromForm() {
  const stations = [];
  const rows = tideStationsContainer.querySelectorAll(".tide-station-row");
  rows.forEach(row => {
    stations.push({
      id: row.dataset.id || ("ts_" + Date.now() + "_" + Math.random().toString(36).slice(2)),
      name: row.querySelector(".ts-name").value.trim(),
      hw1: row.querySelector(".ts-hw1").value,
      hw2: row.querySelector(".ts-hw2").value,
      lw1: row.querySelector(".ts-lw1").value,
      lw2: row.querySelector(".ts-lw2").value,
      auto: row.dataset.auto === "true"
    });
  });
  return stations;
}

function moveTideStation(index, delta) {
  const p = getCurrentPassage();
  if (!p) return;
  p.plan.tideStations = readTideStationsFromForm();
  const stations = p.plan.tideStations;
  const newIndex = index + delta;
  if (newIndex < 0 || newIndex >= stations.length) return;
  const [item] = stations.splice(index, 1);
  stations.splice(newIndex, 0, item);
  renderTideStations(p);
}

addTideStationBtn.addEventListener("click", () => {
  const p = getCurrentPassage();
  if (!p) return;
  p.plan.tideStations = readTideStationsFromForm();
  p.plan.tideStations.push({
    id: "ts_" + Date.now(),
    name: "",
    hw1: "", hw2: "", lw1: "", lw2: "",
    auto: false
  });
  renderTideStations(p);
});

function renderDailySummaries(p) {
  dailySummariesContainer.innerHTML = "";
  const days = p.plan.dailySummaries || [];
  days.forEach((d, index) => {
    const row = document.createElement("div");
    row.className = "daily-summary-row";
    row.dataset.index = index;

    row.innerHTML = `
      <div class="row ds-row">
        <label>
          Date
          <input type="date" class="ds-date" value="${d.date || ""}">
        </label>
        <label>
          Mooring fee
          <input type="text" class="ds-fee" value="${escapeHtml(d.fee || "")}" placeholder="e.g. £35.00">
        </label>
      </div>
      <label>
        Notes
        <textarea class="ds-notes" rows="2">${escapeHtml(d.notes || "")}</textarea>
      </label>
      <button type="button" class="btn btn-secondary btn-small remove-daily-summary" style="margin-top:0.3rem;">
        Remove day
      </button>
    `;

    row.querySelector(".remove-daily-summary").addEventListener("click", () => {
      p.plan.dailySummaries = readDailySummariesFromForm();
      p.plan.dailySummaries.splice(index, 1);
      renderDailySummaries(p);
    });

    dailySummariesContainer.appendChild(row);
  });
}

function readDailySummariesFromForm() {
  const days = [];
  const rows = dailySummariesContainer.querySelectorAll(".daily-summary-row");
  rows.forEach(row => {
    days.push({
      id: "ds_" + Date.now() + "_" + Math.random().toString(36).slice(2),
      date: row.querySelector(".ds-date").value,
      fee: row.querySelector(".ds-fee").value.trim(),
      notes: row.querySelector(".ds-notes").value.trim()
    });
  });
  return days;
}

addDailySummaryBtn.addEventListener("click", () => {
  const p = getCurrentPassage();
  if (!p) return;
  p.plan.dailySummaries = readDailySummariesFromForm();
  p.plan.dailySummaries.push({ id: "ds_" + Date.now(), date: "", fee: "", notes: "" });
  renderDailySummaries(p);
});

// Sync auto tide stations on input (not just change)
let tideSyncTimer = null;
function scheduleAutoTideSync() {
  clearTimeout(tideSyncTimer);
  tideSyncTimer = setTimeout(() => {
    const p = getCurrentPassage();
    if (!p) return;
    p.plan.from = planFrom.value.trim();
    p.plan.to   = planTo.value.trim();
    ensureAutoTideStations(p);
    renderTideStations(p);
    updatePlanSummaryPanel();
    updatePassageHeader();
  }, 120);
}

planFrom.addEventListener("input", scheduleAutoTideSync);
planTo.addEventListener("input", scheduleAutoTideSync);

// Save plan -> remember ports, ensure tide stations, then jump to Log
planForm.addEventListener("submit", (e) => {
  e.preventDefault();
  const p = getCurrentPassage();
  if (!p) return;

  p.plan.date = planDate.value;
  p.plan.from = planFrom.value.trim();
  p.plan.to   = planTo.value.trim();
  p.plan.vessel = planVessel.value.trim();
  p.plan.skipper = planSkipper.value.trim();
  p.plan.crew = planCrew.value.trim();
  p.plan.sunriseSet = planSunriseSet.value.trim();
  p.plan.tidalCoeff = planTidalCoeff.value.trim();
  p.plan.currents = planCurrents.value.trim();
  p.plan.weather = planWeather.value.trim();
  p.plan.comms = planComms.value.trim();

  p.plan.tideStations = readTideStationsFromForm();
  ensureAutoTideStations(p);

  p.plan.dailySummaries = readDailySummariesFromForm();

  savePassages();
  rememberPort(p.plan.from);
  rememberPort(p.plan.to);

  refreshHomePassageList();
  updatePassageHeader();
  updatePlanSummaryPanel();

  switchToTab("logTab");
});

// --- Plan summary panel (no START block) ---------------------------

function updatePlanSummaryPanel() {
  const p = getCurrentPassage();
  if (!p) {
    planSummaryPanel.innerHTML = "<p>No passage selected.</p>";
    return;
  }

  const tidalCoeff = p.plan.tidalCoeff || "";
  const currents = p.plan.currents || "";
  const weather  = p.plan.weather || "";
  const comms    = p.plan.comms || "";
  const tideStations = p.plan.tideStations || [];
  const dailySummaries = p.plan.dailySummaries || [];

  const tideStationsHtml = tideStations.length
    ? tideStations.map(ts => {
        const parts = [];
        if (ts.hw1 || ts.hw2) parts.push(`HW: ${[ts.hw1, ts.hw2].filter(Boolean).join(", ")}`);
        if (ts.lw1 || ts.lw2) parts.push(`LW: ${[ts.lw1, ts.lw2].filter(Boolean).join(", ")}`);
        const detail = parts.length ? " – " + parts.join(" | ") : "";
        return `<div class="tide-row">${escapeHtml(ts.name || "Station")}${detail}</div>`;
      }).join("")
    : "<p><em>–</em></p>";

  const dailySummaryHtml = dailySummaries.length
    ? dailySummaries.map(ds => {
        const dateLabel = ds.date || "No date";
        const feeLabel  = ds.fee  ? ` – ${escapeHtml(ds.fee)}` : "";
        const notesLabel = ds.notes ? ` – ${escapeHtml(ds.notes)}` : "";
        return `<div class="daily-summary-item plan-link" data-goto="dailySummariesContainer">${escapeHtml(dateLabel)}${feeLabel}${notesLabel}</div>`;
      }).join("")
    : "<p class=\"plan-link\" data-goto=\"dailySummariesContainer\"><em>–</em></p>";

  planSummaryPanel.innerHTML = `
    <div class="plan-summary-grid">
      <div class="col">
        <div class="block plan-link" data-goto="planTidalCoeff">
          <p class="section-title">TIDES</p>
          <p>${tidalCoeff ? `<strong>Coeff:</strong> ${escapeHtml(tidalCoeff)}` : "<strong>Coeff:</strong> –"}</p>
          <p><strong>Tide stations:</strong></p>
          ${tideStationsHtml}
        </div>

        <div class="block plan-link" data-goto="planCurrents">
          <p class="section-title">TIDAL CURRENTS / FLOWS</p>
          <p>${currents ? escapeHtml(currents).replace(/\n/g, "<br>") : "<em>–</em>"}</p>
        </div>

        <div class="block plan-link" data-goto="planComms">
          <p class="section-title">COMMS / PILOTAGE</p>
          <p>${comms ? escapeHtml(comms).replace(/\n/g, "<br>") : "<em>–</em>"}</p>
        </div>
      </div>

      <div class="col">
        <div class="block plan-link" data-goto="planWeather">
          <p class="section-title">WEATHER</p>
          <p>${weather ? escapeHtml(weather).replace(/\n/g, "<br>") : "<em>–</em>"}</p>
        </div>

        <div class="block">
          <p class="section-title">DAILY SUMMARY</p>
          ${dailySummaryHtml}
        </div>
      </div>
    </div>
  `;
}

planSummaryPanel.addEventListener("click", (e) => {
  const target = e.target.closest(".plan-link");
  if (!target) return;
  const fieldId = target.dataset.goto;
  if (!fieldId) return;

  switchToTab("planTab");
  const el = document.getElementById(fieldId);
  if (el) setTimeout(() => el.scrollIntoView({ behavior: "smooth", block: "start" }), 50);
});

// --- Log entries ----------------------------------------------------

function passageIsShutdown(p) {
  return p?.finish?.shutdownLogged === true;
}

function addLogEntry() {
  const p = getCurrentPassage();
  if (!p) return alert("No passage selected.");
  ensureFlags(p);
  if (passageIsShutdown(p)) return alert("Shutdown already recorded – no further log entries allowed.");

  const now = new Date();
  const timeStr = now.toISOString().slice(0, 16);
  const previous = p.entries[0] || null;

  const entry = {
    id: "e_" + Date.now(),
    time: timeStr,
    lat: "",
    lon: "",
    course: previous ? previous.course : "",
    speed: previous ? previous.speed : "",
    rpm: previous ? previous.rpm : "",
    engTP: previous ? previous.engTP : "",
    waterLog: previous ? (previous.waterLog || "") : "",
    groundLog: previous ? previous.groundLog : "",
    fuelUsed: previous ? previous.fuelUsed : "",
    notes: ""
  };

  p.entries.unshift(entry);
  savePassages();
  renderLogEntries();
  refreshHomePassageList();
}

function addSpecialEntry(noteText) {
  const p = getCurrentPassage();
  if (!p) return alert("No passage selected.");
  ensureFlags(p);
  if (passageIsShutdown(p)) return alert("Shutdown already recorded – no further log entries allowed.");

  const now = new Date();
  const timeStr = now.toISOString().slice(0, 16);
  const previous = p.entries[0] || null;

  const entry = {
    id: "e_" + Date.now(),
    time: timeStr,
    lat: "",
    lon: "",
    course: previous ? previous.course : "",
    speed: previous ? previous.speed : "",
    rpm: previous ? previous.rpm : "",
    engTP: previous ? previous.engTP : "",
    waterLog: previous ? (previous.waterLog || "") : "",
    groundLog: previous ? previous.groundLog : "",
    fuelUsed: previous ? previous.fuelUsed : "",
    notes: noteText || ""
  };

  p.entries.unshift(entry);
  savePassages();
  renderLogEntries();
  refreshHomePassageList();
}

function addDockEntry() {
  const p = getCurrentPassage();
  if (!p) return alert("No passage selected.");
  ensureFlags(p);
  if (passageIsShutdown(p)) return alert("Shutdown already recorded – no further log entries allowed.");

  const now = new Date();
  const timeStr = now.toISOString().slice(0, 16);
  const previous = p.entries[0] || null;

  const entry = {
    id: "e_" + Date.now(),
    time: timeStr,
    lat: "",
    lon: "",
    course: "",
    speed: "0",
    rpm: "",
    engTP: "",
    waterLog: previous ? (previous.waterLog || "") : "",
    groundLog: previous ? previous.groundLog : "",
    fuelUsed: previous ? previous.fuelUsed : "",
    notes: "Alongside / docked"
  };

  p.entries.unshift(entry);
  savePassages();
  renderLogEntries();
  refreshHomePassageList();
}

function attachSwipeToRow(tr, entryId) {
  let startX = 0;
  tr.addEventListener("touchstart", (e) => { startX = e.changedTouches[0].screenX; }, { passive: true });
  tr.addEventListener("touchend", (e) => {
    const dx = e.changedTouches[0].screenX - startX;
    if (dx < -90) deleteLogEntryById(entryId);
  }, { passive: true });
}

function deleteLogEntryById(entryId) {
  const p = getCurrentPassage();
  if (!p) return;
  const idx = p.entries.findIndex(e => e.id === entryId);
  if (idx < 0) return;

  const ok = confirm("Delete this log entry?");
  if (!ok) return;

  p.entries.splice(idx, 1);
  savePassages();
  renderLogEntries();
  refreshHomePassageList();
}

function handlePositionEdit(entry) {
  function manualPosition() {
    const current = (entry.lat || "") + (entry.lon ? `, ${entry.lon}` : "");
    const val = prompt("Position (decimal \"lat, lon\" or formatted):", current);
    if (val === null) return;
    const result = parseAndFormatPositionInput(val.trim(), entry.lat, entry.lon);
    entry.lat = result.lat;
    entry.lon = result.lon;
    savePassages();
    renderLogEntries();
  }

  if (!navigator.geolocation) return manualPosition();

  const useGps = confirm("Use current GPS position? Press Cancel to enter manually.");
  if (!useGps) return manualPosition();

  navigator.geolocation.getCurrentPosition(
    (pos) => {
      entry.lat = formatLatFromDecimal(pos.coords.latitude);
      entry.lon = formatLonFromDecimal(pos.coords.longitude);
      savePassages();
      renderLogEntries();
    },
    (err) => alert("Unable to get GPS position: " + err.message),
    { enableHighAccuracy: true, maximumAge: 30000, timeout: 15000 }
  );
}

// Engine start: numeric-friendly modal + only once
engineStartBtn.addEventListener("click", () => {
  const p = getCurrentPassage();
  if (!p) return alert("No passage selected.");
  ensureFlags(p);
  if (passageIsShutdown(p)) return alert("Shutdown already recorded – no further log entries allowed.");
  if (p.flags.engineStart) return alert("Engine Start already recorded for this passage.");

  showModal({
    title: "Engine Start",
    bodyHtml: `
      <label style="display:flex;flex-direction:column;gap:0.25rem;margin-bottom:0.5rem;">
        Engine hours at start
        <input id="ehStart" type="number" inputmode="decimal" step="0.1" value="${escapeHtml(p.plan.engineHoursStart || "")}">
      </label>
      <label style="display:flex;flex-direction:column;gap:0.25rem;">
        Fuel % at start
        <input id="fuelStart" type="number" inputmode="numeric" step="1" value="${escapeHtml(p.plan.fuelStartPercent || "")}">
      </label>
    `,
    onOk: () => {
      const eh = document.getElementById("ehStart").value.trim();
      const fu = document.getElementById("fuelStart").value.trim();
      p.plan.engineHoursStart = eh;
      p.plan.fuelStartPercent = fu;

      addSpecialEntry("Engine start");
      p.flags.engineStart = true;

      savePassages();
      updatePlanSummaryPanel();
      updateLogSummary();
    }
  });
});

// Slip: only once
slipLinesBtn.addEventListener("click", () => {
  const p = getCurrentPassage();
  if (!p) return;
  ensureFlags(p);
  if (passageIsShutdown(p)) return alert("Shutdown already recorded – no further log entries allowed.");
  if (p.flags.slip) return alert("Slip already recorded for this passage.");
  addSpecialEntry("Slipped lines / underway");
  p.flags.slip = true;
  savePassages();
});

// Dock: only once
dockLinesBtn.addEventListener("click", () => {
  const p = getCurrentPassage();
  if (!p) return;
  ensureFlags(p);
  if (passageIsShutdown(p)) return alert("Shutdown already recorded – no further log entries allowed.");
  if (p.flags.dock) return alert("Dock already recorded for this passage.");
  addDockEntry();
  p.flags.dock = true;
  savePassages();
});

// Shutdown: one only; keep summary below, keep notes clean
shutdownBtn.addEventListener("click", () => {
  const p = getCurrentPassage();
  if (!p) return alert("No passage selected.");
  if (p.finish.shutdownLogged) return alert("Shutdown has already been recorded for this passage.");

  showModal({
    title: "Shutdown",
    bodyHtml: `
      <label style="display:flex;flex-direction:column;gap:0.25rem;margin-bottom:0.5rem;">
        Engine hours (end)
        <input id="ehEnd" type="number" inputmode="decimal" step="0.1" value="${escapeHtml(p.finish.engineHoursEnd || "")}">
      </label>
      <label style="display:flex;flex-direction:column;gap:0.25rem;margin-bottom:0.5rem;">
        Fuel % at end
        <input id="fuelEnd" type="number" inputmode="numeric" step="1" value="${escapeHtml(p.finish.fuelEndPercent || "")}">
      </label>
      <label style="display:flex;flex-direction:column;gap:0.25rem;">
        Notes / defects
        <input id="shNotes" type="text" value="${escapeHtml(p.finish.notes || "")}">
      </label>
    `,
    onOk: () => {
      p.finish.engineHoursEnd = document.getElementById("ehEnd").value.trim();
      p.finish.fuelEndPercent = document.getElementById("fuelEnd").value.trim();
      p.finish.notes = document.getElementById("shNotes").value.trim();
      p.finish.shutdownLogged = true;

      // Final entry note kept clean
      const note = p.finish.notes ? `Shutdown / alongside – ${p.finish.notes}` : "Shutdown / alongside";

      p.entries.unshift({
        id: "e_" + Date.now(),
        time: new Date().toISOString().slice(0, 16),
        lat: "",
        lon: "",
        course: "",
        speed: "0",
        rpm: "",
        engTP: "",
        waterLog: "",
        groundLog: "",
        fuelUsed: "",
        notes: note
      });

      savePassages();
      renderLogEntries();
      refreshHomePassageList();
      updatePassageHeader();
      updateLogSummary();
    }
  });
});

function renderLogEntries() {
  const p = getCurrentPassage();
  logEntriesContainer.innerHTML = "";

  if (!p || (p.entries?.length || 0) === 0) {
    logEmptyMessage.style.display = "block";
    logSummaryPanel.textContent = "";
    return;
  }
  logEmptyMessage.style.display = "none";

  const entries = p.entries.slice().sort((a, b) => (a.time > b.time ? 1 : -1));

  entries.forEach(entry => {
    const tr = document.createElement("tr");
    attachSwipeToRow(tr, entry.id);

    const tdTime = document.createElement("td");
    tdTime.textContent = entry.time ? timeOnlyFromIso(entry.time) : "";
    tdTime.classList.add("editable-cell");
    tdTime.addEventListener("click", () => {
      const val = prompt("Time (YYYY-MM-DD HH:MM or HH:MM):", entry.time || "");
      if (val === null) return;
      entry.time = val.trim();
      savePassages();
      renderLogEntries();
    });
    tr.appendChild(tdTime);

    function addInputCell(value, opts) {
      const td = document.createElement("td");
      const inp = document.createElement("input");
      inp.className = opts.className;
      inp.type = opts.type || "text";
      if (opts.inputMode) inp.inputMode = opts.inputMode;
      if (opts.step) inp.step = opts.step;
      inp.value = value || "";
      inp.addEventListener("change", () => {
        entry[opts.field] = inp.value.trim();
        savePassages();
        updateLogSummary();
      });
      td.appendChild(inp);
      tr.appendChild(td);
    }

    addInputCell(entry.course, { field: "course", className: "log-input log-input-cog", type: "text", inputMode: "numeric" });
    addInputCell(entry.speed,  { field: "speed",  className: "log-input log-input-num", type: "number", inputMode: "decimal", step: "0.1" });
    addInputCell(entry.rpm,    { field: "rpm",    className: "log-input log-input-num", type: "number", inputMode: "numeric", step: "10" });
    addInputCell(entry.engTP,  { field: "engTP",  className: "log-input log-input-num", type: "text",   inputMode: "decimal" });

    addInputCell(entry.waterLog || "", { field: "waterLog", className: "log-input log-input-num", type: "number", inputMode: "decimal", step: "0.1" });
    addInputCell(entry.groundLog,      { field: "groundLog", className: "log-input log-input-num", type: "number", inputMode: "decimal", step: "0.1" });

    addInputCell(entry.fuelUsed, { field: "fuelUsed", className: "log-input log-input-num", type: "number", inputMode: "decimal", step: "0.1" });

    const tdNotes = document.createElement("td");

    const notesText = document.createElement("div");
    notesText.textContent = entry.notes || "";
    notesText.classList.add("editable-cell");
    notesText.addEventListener("click", () => {
      const val = prompt("Notes:", entry.notes || "");
      if (val === null) return;
      entry.notes = val.trim();
      savePassages();
      renderLogEntries();
    });
    tdNotes.appendChild(notesText);

    const actions = document.createElement("div");
    actions.className = "entry-actions";

    const hasPos = (entry.lat && entry.lat.trim()) || (entry.lon && entry.lon.trim());
    if (!hasPos) {
      const posBtn = document.createElement("button");
      posBtn.className = "btn btn-secondary btn-small";
      posBtn.textContent = "Position";
      posBtn.addEventListener("click", () => handlePositionEdit(entry));
      actions.appendChild(posBtn);
    } else {
      const posSpan = document.createElement("span");
      posSpan.className = "pos-field";
      posSpan.textContent = entry.lat && entry.lon ? `${entry.lat}, ${entry.lon}` : (entry.lat || entry.lon);
      posSpan.addEventListener("click", () => handlePositionEdit(entry));
      actions.appendChild(posSpan);
    }

    const delBtn = document.createElement("button");
    delBtn.className = "entry-del-btn";
    delBtn.textContent = "Del";
    delBtn.addEventListener("click", () => deleteLogEntryById(entry.id));
    actions.appendChild(delBtn);

    tdNotes.appendChild(actions);
    tr.appendChild(tdNotes);

    logEntriesContainer.appendChild(tr);
  });

  updateLogSummary();
}

function updateLogSummary() {
  const p = getCurrentPassage();
  if (!p || !p.finish.shutdownLogged) {
    logSummaryPanel.textContent = "";
    return;
  }

  let ehText = "–";
  const start = parseFloat(p.plan.engineHoursStart || "NaN");
  const endVal = parseFloat(p.finish.engineHoursEnd || "NaN");
  if (!isNaN(start) && !isNaN(endVal)) ehText = `${(endVal - start).toFixed(1)} h (from ${start} to ${endVal})`;

  let fuelUsed = "–";
  const sorted = p.entries.slice().sort((a, b) => (a.time > b.time ? 1 : -1));
  for (let i = sorted.length - 1; i >= 0; i--) {
    const fu = parseFloat(sorted[i].fuelUsed || "NaN");
    if (!isNaN(fu)) { fuelUsed = `${fu}`; break; }
  }

  let gLog = "–";
  for (let i = sorted.length - 1; i >= 0; i--) {
    if (sorted[i].groundLog) { gLog = sorted[i].groundLog; break; }
  }

  let durationText = "–";
  const times = sorted.map(e => e.time).filter(Boolean).map(t => new Date(t));
  if (times.length >= 2) {
    const min = times.reduce((a, b) => (a < b ? a : b));
    const max = times.reduce((a, b) => (a > b ? a : b));
    const ms = max - min;
    if (!isNaN(ms) && ms > 0) {
      const minutes = Math.round(ms / 60000);
      durationText = `${Math.floor(minutes / 60)}h ${minutes % 60}m`;
    }
  }

  logSummaryPanel.innerHTML = `
    <strong>Summary:</strong>
    Engine hours this passage: ${ehText} |
    Fuel used: ${fuelUsed} |
    Fuel start: ${p.plan.fuelStartPercent || "–"}% |
    Fuel end: ${p.finish.fuelEndPercent || "–"}% |
    Final GLog: ${gLog} |
    Passage duration: ${durationText}
  `;
}

// CSV Export
function exportCurrentPassageToCsv() {
  const p = getCurrentPassage();
  if (!p) return alert("No passage selected.");

  const date = p.plan.date || p.createdAt.slice(0, 10);
  const from = p.plan.from || "UnknownFrom";
  const to = p.plan.to || "UnknownTo";
  const filename = `${date} ${from} - ${to}.csv`.replace(/[/\\?%*:|"<>]/g, "-");

  const lines = [];
  lines.push("Passage Plan");
  lines.push(`Date,${quote(date)}`);
  lines.push(`Origin,${quote(p.plan.from)}`);
  lines.push(`Intended Destination,${quote(p.plan.to)}`);
  lines.push(`Vessel,${quote(p.plan.vessel)}`);
  lines.push(`Skipper,${quote(p.plan.skipper)}`);
  lines.push(`Crew,${quote(p.plan.crew)}`);
  lines.push("");
  lines.push(`Sunrise/Set,${quote(p.plan.sunriseSet)}`);
  lines.push(`Tidal Coefficient,${quote(p.plan.tidalCoeff)}`);
  lines.push("");

  lines.push("Tide Stations");
  lines.push("Station,HW1,HW2,LW1,LW2");
  (p.plan.tideStations || []).forEach(ts => {
    lines.push([ts.name || "", ts.hw1 || "", ts.hw2 || "", ts.lw1 || "", ts.lw2 || ""].map(quote).join(","));
  });
  lines.push("");

  lines.push("Tidal Currents / Flows");
  lines.push(quote(p.plan.currents));
  lines.push("");

  lines.push("Weather");
  lines.push(quote(p.plan.weather));
  lines.push("");

  lines.push("Comms / Pilotage");
  lines.push(quote(p.plan.comms));
  lines.push("");

  lines.push("Daily Summary");
  lines.push("Date,Mooring fee,Notes");
  (p.plan.dailySummaries || []).forEach(ds => {
    lines.push([ds.date || "", ds.fee || "", ds.notes || ""].map(quote).join(","));
  });
  lines.push("");

  lines.push(`Engine hours start,${quote(p.plan.engineHoursStart)}`);
  lines.push(`Fuel start %,${quote(p.plan.fuelStartPercent)}`);
  lines.push("");

  lines.push("Log Entries");
  lines.push(["Time","Lat","Lon","COG/Heading","Speed (kn)","RPM","Eng T/P","WLog (NM)","GLog (NM)","Fuel used","Notes"].map(quote).join(","));

  p.entries.slice().sort((a, b) => (a.time > b.time ? 1 : -1)).forEach(e => {
    lines.push([
      e.time ? e.time.replace("T", " ") : "",
      e.lat, e.lon, e.course, e.speed, e.rpm, e.engTP, e.waterLog, e.groundLog, e.fuelUsed, e.notes
    ].map(quote).join(","));
  });

  lines.push("");
  lines.push("End of Passage");
  lines.push(`Engine hours end,${quote(p.finish.engineHoursEnd)}`);
  lines.push(`Fuel end %,${quote(p.finish.fuelEndPercent)}`);
  lines.push(`Summary notes,${quote(p.finish.notes)}`);

  const csvContent = lines.join("\r\n");
  const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);

  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

exportCsvBtn.addEventListener("click", exportCurrentPassageToCsv);
addEntryBtn.addEventListener("click", () => addLogEntry());

// --- Load passage into UI -----------------------------------------

function loadPassageIntoUI() {
  const p = getCurrentPassage();
  if (!p) {
    planForm?.reset();
    logEntriesContainer.innerHTML = "";
    logEmptyMessage.style.display = "block";
    planSummaryPanel.innerHTML = "<p>No passage selected.</p>";
    logSummaryPanel.textContent = "";
    updatePassageHeader();
    return;
  }

  ensureFlags(p);
  ensureAutoTideStations(p);

  updatePassageHeader();
  loadPlanIntoForm(p);
  updatePlanSummaryPanel();
  renderLogEntries();
  updateLogSummary();
}

// --- Create new passage -------------------------------------------

homeNewPassageBtn.addEventListener("click", () => {
  if (passages.length > 0) {
    const ok = confirm("Start a new passage? (Existing ones will remain in history.)");
    if (!ok) return;
  }
  createPassage();
  switchToTab("planTab");
});

// --- Initial load --------------------------------------------------

loadPassages();
loadPorts();
renderPortsDatalist();
applyTheme(localStorage.getItem(THEME_KEY) || "day");

refreshHomePassageList();

if (!currentPassageId && passages.length > 0) currentPassageId = passages[0].id;

loadPassageIntoUI();
setLogLayoutMode("split", splitViewBtn);

// Service worker registration
if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => {
    navigator.serviceWorker.register("service-worker.js").catch(err => {
      console.warn("Service worker registration failed", err);
    });
  });
}
