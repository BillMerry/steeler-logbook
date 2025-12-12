// --- Constants & state ---------------------------------------------

const STORAGE_KEY = "steeler_logbook_passages_v4";
const THEME_KEY = "steeler_logbook_theme_v1";
const PORTS_KEY = "steeler_logbook_ports_v1";

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

function formatDateForCard(dateStr) {
  if (!dateStr) return "Unknown date";
  return dateStr;
}

function timeOnlyFromIso(iso) {
  if (!iso || iso.length < 16) return iso || "";
  return iso.slice(11, 16);
}

function switchToTab(tabId) {
  tabButtons.forEach(b => {
    if (b.dataset.tab === tabId) {
      b.classList.add("active");
    } else {
      b.classList.remove("active");
    }
  });
  tabs.forEach(t => {
    if (t.id === tabId) {
      t.classList.add("active");
    } else {
      t.classList.remove("active");
    }
  });
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
  if (!val) {
    return { lat: "", lon: "" };
  }

  // If it looks like DMS already (has º or N/S/E/W), just split & store
  if (/[º°NnSsEeWw]/.test(val)) {
    const parts = val.split(",").map(s => s.trim());
    return {
      lat: parts[0] || currentLat || "",
      lon: parts[1] || currentLon || ""
    };
  }

  // Otherwise assume decimal degrees "lat, lon"
  const parts = val.split(",").map(s => s.trim());
  const latNum = parseFloat(parts[0]);
  const lonNum = parseFloat(parts[1]);

  if (isNaN(latNum) || isNaN(lonNum)) {
    return { lat: val, lon: currentLon || "" };
  }

  return {
    lat: formatLatFromDecimal(latNum),
    lon: formatLonFromDecimal(lonNum)
  };
}

// --- DOM references ------------------------------------------------

// Header
const headerPassageMain = document.getElementById("headerPassageMain");
const headerSunrise = document.getElementById("headerSunrise");
const headerCrew = document.getElementById("headerCrew");
const themeToggleBtn = document.getElementById("themeToggleBtn");

// Tabs
const tabButtons = document.querySelectorAll(".tab-btn");
const tabs = document.querySelectorAll(".tab");

// Home tab
const homeNewPassageBtn = document.getElementById("homeNewPassageBtn");
const homePassageList = document.getElementById("homePassageList");

// Plan form fields
const planForm = document.getElementById("planForm");
const planDate = document.getElementById("planDate");
const planFrom = document.getElementById("planFrom");
const planTo = document.getElementById("planTo");
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

// Log tab
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

// --- Theme handling -----------------------------------------------

function applyTheme(theme) {
  document.body.dataset.theme = theme;
  localStorage.setItem(THEME_KEY, theme);
  if (themeToggleBtn) {
    themeToggleBtn.textContent = theme === "night" ? "Day" : "Night";
  }
}

if (themeToggleBtn) {
  themeToggleBtn.addEventListener("click", () => {
    const current = document.body.dataset.theme || "day";
    const next = current === "night" ? "day" : "night";
    applyTheme(next);
  });
}

// --- Tabs ----------------------------------------------------------

tabButtons.forEach(btn => {
  btn.addEventListener("click", () => {
    const target = btn.dataset.tab;

    tabButtons.forEach(b => b.classList.remove("active"));
    tabs.forEach(t => t.classList.remove("active"));

    btn.classList.add("active");
    document.getElementById(target).classList.add("active");
  });
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
  const to = p.plan.to || "?";

  const sunriseSet = p.plan.sunriseSet || "";
  const skipper = p.plan.skipper || "";
  const crew = p.plan.crew || "";

  headerPassageMain.textContent = `${date} – ${from} → ${to}`;
  headerSunrise.textContent = sunriseSet ? `Sunrise–Set: ${sunriseSet}` : "";

  const crewParts = [];
  if (skipper) crewParts.push(`Skipper: ${skipper}`);
  if (crew) crewParts.push(`Crew: ${crew}`);
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

// --- HOME: passage list --------------------------------------------

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

    const date = formatDateForCard(passage.plan.date || passage.createdAt.slice(0, 10));
    const from = passage.plan.from || "?";
    const to = passage.plan.to || "?";

    const status = passage.finish && passage.finish.shutdownLogged ? "Completed" : "In progress";
    const entriesCount = passage.entries.length;

    const title = document.createElement("div");
    title.className = "passage-card-title";
    title.textContent = `${date} – ${from} → ${to}`;

    const meta = document.createElement("div");
    meta.className = "passage-card-meta";
    meta.innerHTML = `<span>${entriesCount} entries</span><span>${status}</span>`;

    card.appendChild(title);
    card.appendChild(meta);

    card.addEventListener("click", () => {
      currentPassageId = passage.id;
      loadPassageIntoUI();
      switchToTab("logTab");
    });

    homePassageList.appendChild(card);
  });
}

// --- Layout mode controls (Log tab) -------------------------------

function setLogLayoutMode(mode) {
  logLayout.classList.remove("split", "plan-only", "log-only");
  if (mode === "plan-only") {
    logLayout.classList.add("plan-only");
  } else if (mode === "log-only") {
    logLayout.classList.add("log-only");
  } else {
    logLayout.classList.add("split");
  }
}

splitViewBtn.addEventListener("click", () => setLogLayoutMode("split"));
expandPlanBtn.addEventListener("click", () => setLogLayoutMode("plan-only"));
expandLogBtn.addEventListener("click", () => setLogLayoutMode("log-only"));

// --- Plan tab logic -----------------------------------------------

function ensureDefaultTideStations(p) {
  if (!p) return;
  if (p.plan.tideStations && p.plan.tideStations.length > 0) return;

  const origin = (p.plan.from || "").trim();
  const dest = (p.plan.to || "").trim();

  const stations = [];
  if (origin) {
    stations.push({
      id: "ts_" + Date.now() + "_orig",
      name: origin,
      hw1: "",
      hw2: "",
      lw1: "",
      lw2: ""
    });
  }

  if (dest && dest.toLowerCase() !== "local" && dest !== origin) {
    stations.push({
      id: "ts_" + Date.now() + "_dest",
      name: dest,
      hw1: "",
      hw2: "",
      lw1: "",
      lw2: ""
    });
  }

  if (stations.length > 0) {
    p.plan.tideStations = stations;
  }
}

function createPassage() {
  const id = "p_" + Date.now();
  const today = new Date().toISOString().slice(0, 10);

  const passage = {
    id,
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
        {
          id: "ds_" + Date.now(),
          date: today,
          fee: "",
          notes: ""
        }
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

  ensureDefaultTideStations(passage);
  passages.unshift(passage);
  currentPassageId = id;
  savePassages();
  refreshHomePassageList();
  loadPassageIntoUI();
}

function loadPlanIntoForm(p) {
  planDate.value = p.plan.date || "";
  planFrom.value = p.plan.from || "";
  planTo.value = p.plan.to || "";
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

    row.innerHTML = `
      <div class="row">
        <label>
          Tide station
          <input type="text" class="ts-name" value="${escapeHtml(st.name || "")}">
        </label>
        <button type="button" class="secondary-btn small remove-tide-station">Remove</button>
      </div>
      <div class="row">
        <label>
          HW 1
          <input type="time" class="ts-hw1" value="${st.hw1 || ""}">
        </label>
        <label>
          HW 2
          <input type="time" class="ts-hw2" value="${st.hw2 || ""}">
        </label>
      </div>
      <div class="row">
        <label>
          LW 1
          <input type="time" class="ts-lw1" value="${st.lw1 || ""}">
        </label>
        <label>
          LW 2
          <input type="time" class="ts-lw2" value="${st.lw2 || ""}">
        </label>
      </div>
    `;

    const headerRow = row.querySelector(".row");
    const removeBtn = row.querySelector(".remove-tide-station");
    removeBtn.addEventListener("click", () => {
      p.plan.tideStations = readTideStationsFromForm();
      p.plan.tideStations.splice(index, 1);
      renderTideStations(p);
    });

    // Reorder buttons
    const moveUpBtn = document.createElement("button");
    moveUpBtn.type = "button";
    moveUpBtn.textContent = "↑";
    moveUpBtn.className = "secondary-btn small";
    moveUpBtn.addEventListener("click", () => moveTideStation(index, -1));

    const moveDownBtn = document.createElement("button");
    moveDownBtn.type = "button";
    moveDownBtn.textContent = "↓";
    moveDownBtn.className = "secondary-btn small";
    moveDownBtn.addEventListener("click", () => moveTideStation(index, 1));

    headerRow.appendChild(moveUpBtn);
    headerRow.appendChild(moveDownBtn);

    tideStationsContainer.appendChild(row);
  });
}

function readTideStationsFromForm() {
  const stations = [];
  const rows = tideStationsContainer.querySelectorAll(".tide-station-row");
  rows.forEach(row => {
    const name = row.querySelector(".ts-name").value.trim();
    const hw1 = row.querySelector(".ts-hw1").value;
    const hw2 = row.querySelector(".ts-hw2").value;
    const lw1 = row.querySelector(".ts-lw1").value;
    const lw2 = row.querySelector(".ts-lw2").value;
    stations.push({
      id: "ts_" + Date.now() + "_" + Math.random().toString(36).slice(2),
      name,
      hw1,
      hw2,
      lw1,
      lw2
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
  if (!p.plan.tideStations) p.plan.tideStations = [];
  p.plan.tideStations.push({
    id: "ts_" + Date.now(),
    name: "",
    hw1: "",
    hw2: "",
    lw1: "",
    lw2: ""
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
        <label class="ds-date-field">
          Date
          <input type="date" class="ds-date" value="${d.date || ""}">
        </label>
        <label class="ds-fee-field">
          Mooring fee
          <input type="text" class="ds-fee" value="${escapeHtml(d.fee || "")}" placeholder="e.g. £35.00">
        </label>
      </div>
      <label>
        Notes
        <textarea class="ds-notes" rows="2">${escapeHtml(d.notes || "")}</textarea>
      </label>
      <button type="button" class="secondary-btn small remove-daily-summary" style="margin-top:0.3rem;">
        Remove day
      </button>
    `;

    const removeBtn = row.querySelector(".remove-daily-summary");
    removeBtn.addEventListener("click", () => {
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
    const date = row.querySelector(".ds-date").value;
    const fee = row.querySelector(".ds-fee").value.trim();
    const notes = row.querySelector(".ds-notes").value.trim();
    days.push({
      id: "ds_" + Date.now() + "_" + Math.random().toString(36).slice(2),
      date,
      fee,
      notes
    });
  });
  return days;
}

addDailySummaryBtn.addEventListener("click", () => {
  const p = getCurrentPassage();
  if (!p) return;
  p.plan.dailySummaries = readDailySummariesFromForm();
  if (!p.plan.dailySummaries) p.plan.dailySummaries = [];
  p.plan.dailySummaries.push({
    id: "ds_" + Date.now(),
    date: "",
    fee: "",
    notes: ""
  });
  renderDailySummaries(p);
});

planForm.addEventListener("submit", (e) => {
  e.preventDefault();
  const p = getCurrentPassage();
  if (!p) return;

  p.plan.date = planDate.value;
  p.plan.from = planFrom.value.trim();
  p.plan.to = planTo.value.trim();
  p.plan.vessel = planVessel.value.trim();
  p.plan.skipper = planSkipper.value.trim();
  p.plan.crew = planCrew.value.trim();
  p.plan.sunriseSet = planSunriseSet.value.trim();
  p.plan.tidalCoeff = planTidalCoeff.value.trim();
  p.plan.currents = planCurrents.value.trim();
  p.plan.weather = planWeather.value.trim();
  p.plan.comms = planComms.value.trim();

  p.plan.tideStations = readTideStationsFromForm();
  ensureDefaultTideStations(p);
  p.plan.dailySummaries = readDailySummariesFromForm();

  savePassages();
  rememberPort(p.plan.from);
  rememberPort(p.plan.to);
  refreshHomePassageList();
  updatePassageHeader();
  updatePlanSummaryPanel();

  switchToTab("logTab");
});

// --- Plan summary panel -------------------------------------------

function updatePlanSummaryPanel() {
  if (!planSummaryPanel) return;
  const p = getCurrentPassage();
  if (!p) {
    planSummaryPanel.innerHTML = "<p>No passage selected.</p>";
    return;
  }

  const tidalCoeff = p.plan.tidalCoeff || "";
  const currents = p.plan.currents || "";
  const weather = p.plan.weather || "";
  const comms = p.plan.comms || "";
  const engStart = p.plan.engineHoursStart || "";
  const fuelStart = p.plan.fuelStartPercent || "";
  const tideStations = p.plan.tideStations || [];
  const dailySummaries = p.plan.dailySummaries || [];

  let tideStationsHtml = "";
  if (tideStations.length > 0) {
    tideStationsHtml = tideStations
      .map(ts => {
        const parts = [];
        if (ts.hw1 || ts.hw2) {
          parts.push(`HW: ${[ts.hw1, ts.hw2].filter(Boolean).join(", ")}`);
        }
        if (ts.lw1 || ts.lw2) {
          parts.push(`LW: ${[ts.lw1, ts.lw2].filter(Boolean).join(", ")}`);
        }
        const detail = parts.length ? " – " + parts.join(" | ") : "";
        return `<div class="tide-row">${escapeHtml(ts.name || "Station")}${detail}</div>`;
      })
      .join("");
  } else {
    tideStationsHtml = "<p><em>–</em></p>";
  }

  let dailySummaryHtml = "";
  if (dailySummaries.length > 0) {
    dailySummaryHtml = dailySummaries
      .map(ds => {
        const dateLabel = ds.date || "No date";
        const feeLabel = ds.fee ? ` – ${escapeHtml(ds.fee)}` : "";
        const notesLabel = ds.notes ? ` – ${escapeHtml(ds.notes)}` : "";
        return `<div class="daily-summary-item plan-link" data-goto="dailySummariesContainer">${escapeHtml(dateLabel)}${feeLabel}${notesLabel}</div>`;
      })
      .join("");
  } else {
    dailySummaryHtml = "<p class=\"plan-link\" data-goto=\"dailySummariesContainer\"><em>–</em></p>";
  }

  planSummaryPanel.innerHTML = `
    <div class="plan-summary-grid">
      <div class="col col-1">
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

        <div class="block">
          <p class="section-title">START</p>
          <p>Engine hours: ${engStart || "–"}</p>
          <p>Fuel start: ${fuelStart ? fuelStart + "%" : "–"}</p>
        </div>
      </div>

      <div class="col col-2">
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

// Clicking plan summary → jump to Plan field
planSummaryPanel.addEventListener("click", (e) => {
  const target = e.target.closest(".plan-link");
  if (!target) return;
  const fieldId = target.dataset.goto;
  if (!fieldId) return;

  switchToTab("planTab");
  const el = document.getElementById(fieldId);
  if (el) {
    setTimeout(() => {
      el.scrollIntoView({ behavior: "smooth", block: "start" });
      if (el.focus) {
        el.focus({ preventScroll: true });
      }
    }, 50);
  }
});

// --- Log entries ---------------------------------------------------

function makeEditableCell(td, entryId, field, label) {
  td.classList.add("editable-cell");
  td.addEventListener("click", () => {
    const p = getCurrentPassage();
    if (!p) return;
    const entry = p.entries.find(e => e.id === entryId);
    if (!entry) return;

    const currentVal = entry[field] || "";
    const newVal = prompt(label, currentVal);
    if (newVal === null) return;

    entry[field] = newVal.trim();
    savePassages();
    renderLogEntries();
  });
}

function passageIsShutdown(p) {
  return p && p.finish && p.finish.shutdownLogged;
}

function preventIfShutdown(actionName) {
  const p = getCurrentPassage();
  if (!p) {
    alert("No passage selected.");
    return true;
  }
  if (passageIsShutdown(p)) {
    alert(`Shutdown already recorded – no further ${actionName} entries allowed.`);
    return true;
  }
  return false;
}

function addLogEntry() {
  const p = getCurrentPassage();
  if (!p || passageIsShutdown(p)) {
    if (!p) alert("No passage selected.");
    else alert("Shutdown already recorded – no further log entries allowed.");
    return;
  }

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
    waterLog: previous ? (previous.waterLog || previous.logReading || "") : "",
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
  if (!p || passageIsShutdown(p)) {
    if (!p) alert("No passage selected.");
    else alert("Shutdown already recorded – no further log entries allowed.");
    return;
  }

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
    waterLog: previous ? (previous.waterLog || previous.logReading || "") : "",
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
  if (!p || passageIsShutdown(p)) {
    if (!p) alert("No passage selected.");
    else alert("Shutdown already recorded – no further log entries allowed.");
    return;
  }

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
    waterLog: previous ? (previous.waterLog || previous.logReading || "") : "",
    groundLog: previous ? previous.groundLog : "",
    fuelUsed: previous ? previous.fuelUsed : "",
    notes: "Alongside / docked"
  };

  p.entries.unshift(entry);
  savePassages();
  renderLogEntries();
  refreshHomePassageList();
}

// Engine start now captures engine hours & fuel%

engineStartBtn.addEventListener("click", () => {
  const p = getCurrentPassage();
  if (!p) {
    alert("No passage selected.");
    return;
  }
  if (passageIsShutdown(p)) {
    alert("Shutdown already recorded – no further log entries allowed.");
    return;
  }

  if (!p.plan.engineHoursStart) {
    let eh = prompt("Engine hours at start:", p.plan.engineHoursStart || "");
    if (eh === null) return;
    eh = eh.trim();
    p.plan.engineHoursStart = eh;
  }

  if (!p.plan.fuelStartPercent) {
    let fp = prompt("Fuel % at start:", p.plan.fuelStartPercent || "");
    if (fp === null) return;
    fp = fp.trim();
    p.plan.fuelStartPercent = fp;
  }

  const now = new Date();
  const timeStr = now.toISOString().slice(0, 16);

  const previous = p.entries[0] || null;

  const noteParts = ["Engine start"];
  if (p.plan.engineHoursStart) noteParts.push(`Engine hours: ${p.plan.engineHoursStart}`);
  if (p.plan.fuelStartPercent) noteParts.push(`Fuel: ${p.plan.fuelStartPercent}%`);

  const entry = {
    id: "e_" + Date.now(),
    time: timeStr,
    lat: "",
    lon: "",
    course: previous ? previous.course : "",
    speed: previous ? previous.speed : "",
    rpm: previous ? previous.rpm : "",
    engTP: previous ? previous.engTP : "",
    waterLog: previous ? (previous.waterLog || previous.logReading || "") : "",
    groundLog: previous ? previous.groundLog : "",
    fuelUsed: previous ? previous.fuelUsed : "",
    notes: noteParts.join(" | ")
  };

  p.entries.unshift(entry);
  savePassages();
  renderLogEntries();
  refreshHomePassageList();
  updatePlanSummaryPanel();
});

slipLinesBtn.addEventListener("click", () => {
  addSpecialEntry("Slipped lines / underway");
});

dockLinesBtn.addEventListener("click", addDockEntry);

// Render log entries ------------------------------------------------

function renderLogEntries() {
  const p = getCurrentPassage();
  logEntriesContainer.innerHTML = "";

  if (!p || p.entries.length === 0) {
    if (logEmptyMessage) logEmptyMessage.style.display = "block";
    if (logSummaryPanel) logSummaryPanel.textContent = "";
    return;
  } else {
    if (logEmptyMessage) logEmptyMessage.style.display = "none";
  }

  const entries = p.entries.slice().sort((a, b) => (a.time > b.time ? 1 : -1));

  entries.forEach(entry => {
    const tr = document.createElement("tr");

    // Time (prompt)
    const tdTime = document.createElement("td");
    tdTime.textContent = entry.time ? timeOnlyFromIso(entry.time) : "";
    makeEditableCell(tdTime, entry.id, "time", "Time (YYYY-MM-DD HH:MM or HH:MM)");
    tr.appendChild(tdTime);

    // COG
    const tdCog = document.createElement("td");
    const inputCog = document.createElement("input");
    inputCog.type = "text";
    inputCog.inputMode = "numeric";
    inputCog.className = "log-input log-input-cog";
    inputCog.value = entry.course || "";
    inputCog.addEventListener("change", () => {
      const p2 = getCurrentPassage();
      if (!p2) return;
      const e2 = p2.entries.find(e => e.id === entry.id);
      if (!e2) return;
      e2.course = inputCog.value.trim();
      savePassages();
    });
    tdCog.appendChild(inputCog);
    tr.appendChild(tdCog);

    // Speed
    const tdSpeed = document.createElement("td");
    const inputSpeed = document.createElement("input");
    inputSpeed.type = "number";
    inputSpeed.inputMode = "decimal";
    inputSpeed.className = "log-input log-input-num";
    inputSpeed.value = entry.speed || "";
    inputSpeed.addEventListener("change", () => {
      const p2 = getCurrentPassage();
      if (!p2) return;
      const e2 = p2.entries.find(e => e.id === entry.id);
      if (!e2) return;
      e2.speed = inputSpeed.value.trim();
      savePassages();
    });
    tdSpeed.appendChild(inputSpeed);
    tr.appendChild(tdSpeed);

    // RPM
    const tdRpm = document.createElement("td");
    const inputRpm = document.createElement("input");
    inputRpm.type = "number";
    inputRpm.inputMode = "decimal";
    inputRpm.className = "log-input log-input-num";
    inputRpm.value = entry.rpm || "";
    inputRpm.addEventListener("change", () => {
      const p2 = getCurrentPassage();
      if (!p2) return;
      const e2 = p2.entries.find(e => e.id === entry.id);
      if (!e2) return;
      e2.rpm = inputRpm.value.trim();
      savePassages();
    });
    tdRpm.appendChild(inputRpm);
    tr.appendChild(tdRpm);

    // Eng T/P
    const tdEng = document.createElement("td");
    const inputEng = document.createElement("input");
    inputEng.type = "text";
    inputEng.inputMode = "decimal";
    inputEng.className = "log-input log-input-num";
    inputEng.value = entry.engTP || "";
    inputEng.addEventListener("change", () => {
      const p2 = getCurrentPassage();
      if (!p2) return;
      const e2 = p2.entries.find(e => e.id === entry.id);
      if (!e2) return;
      e2.engTP = inputEng.value.trim();
      savePassages();
    });
    tdEng.appendChild(inputEng);
    tr.appendChild(tdEng);

    // WLog
    const tdWaterLog = document.createElement("td");
    const waterVal = entry.waterLog || entry.logReading || "";
    const inputWLog = document.createElement("input");
    inputWLog.type = "number";
    inputWLog.inputMode = "decimal";
    inputWLog.className = "log-input log-input-num";
    inputWLog.value = waterVal;
    inputWLog.addEventListener("change", () => {
      const p2 = getCurrentPassage();
      if (!p2) return;
      const e2 = p2.entries.find(e => e.id === entry.id);
      if (!e2) return;
      e2.waterLog = inputWLog.value.trim();
      savePassages();
      updateLogSummary();
    });
    tdWaterLog.appendChild(inputWLog);
    tr.appendChild(tdWaterLog);

    // GLog
    const tdGroundLog = document.createElement("td");
    const inputGLog = document.createElement("input");
    inputGLog.type = "number";
    inputGLog.inputMode = "decimal";
    inputGLog.className = "log-input log-input-num";
    inputGLog.value = entry.groundLog || "";
    inputGLog.addEventListener("change", () => {
      const p2 = getCurrentPassage();
      if (!p2) return;
      const e2 = p2.entries.find(e => e.id === entry.id);
      if (!e2) return;
      e2.groundLog = inputGLog.value.trim();
      savePassages();
      updateLogSummary();
    });
    tdGroundLog.appendChild(inputGLog);
    tr.appendChild(tdGroundLog);

    // Fuel Used
    const tdFuel = document.createElement("td");
    const inputFuel = document.createElement("input");
    inputFuel.type = "number";
    inputFuel.inputMode = "decimal";
    inputFuel.className = "log-input log-input-num";
    inputFuel.value = entry.fuelUsed || "";
    inputFuel.addEventListener("change", () => {
      const p2 = getCurrentPassage();
      if (!p2) return;
      const e2 = p2.entries.find(e => e.id === entry.id);
      if (!e2) return;
      e2.fuelUsed = inputFuel.value.trim();
      savePassages();
      updateLogSummary();
    });
    tdFuel.appendChild(inputFuel);
    tr.appendChild(tdFuel);

    // Notes + Position
    const tdNotes = document.createElement("td");

    const notesText = document.createElement("div");
    notesText.textContent = entry.notes || "";
    notesText.style.marginBottom = "0.2rem";
    notesText.classList.add("editable-cell");
    notesText.addEventListener("click", () => {
      const p2 = getCurrentPassage();
      if (!p2) return;
      const e2 = p2.entries.find(e => e.id === entry.id);
      if (!e2) return;
      const val = prompt("Notes:", e2.notes || "");
      if (val === null) return;
      e2.notes = val.trim();
      savePassages();
      renderLogEntries();
    });
    tdNotes.appendChild(notesText);

    const actions = document.createElement("div");

    const hasPos = (entry.lat && entry.lat.trim()) || (entry.lon && entry.lon.trim());
    if (!hasPos) {
      const posBtn = document.createElement("button");
      posBtn.className = "secondary-btn small";
      posBtn.textContent = "Position";
      posBtn.addEventListener("click", () => handlePositionEdit(entry.id));
      actions.appendChild(posBtn);
    } else {
      const posSpan = document.createElement("span");
      posSpan.className = "pos-field";
      const latText = entry.lat || "";
      const lonText = entry.lon || "";
      posSpan.textContent = latText && lonText ? `${latText}, ${lonText}` : (latText || lonText);
      posSpan.addEventListener("click", () => handlePositionEdit(entry.id));
      actions.appendChild(posSpan);
    }

    tdNotes.appendChild(actions);
    tr.appendChild(tdNotes);

    logEntriesContainer.appendChild(tr);
  });

  updateLogSummary();
}

// --- Position edit / GPS -------------------------------------------

function handlePositionEdit(entryId) {
  const p = getCurrentPassage();
  if (!p) return;
  const entry = p.entries.find(e => e.id === entryId);
  if (!entry) return;

  function manualPosition() {
    const current = (entry.lat || "") + (entry.lon ? `, ${entry.lon}` : "");
    const val = prompt(
      "Position (decimal \"lat, lon\" or formatted with º and N/S/E/W):",
      current
    );
    if (val === null) return;
    const result = parseAndFormatPositionInput(val.trim(), entry.lat, entry.lon);
    entry.lat = result.lat;
    entry.lon = result.lon;
    savePassages();
    renderLogEntries();
  }

  if (!navigator.geolocation) {
    manualPosition();
    return;
  }

  const useGps = confirm("Use current GPS position? Press Cancel to enter manually.");
  if (useGps) {
    getGpsForEntry(entryId);
  } else {
    manualPosition();
  }
}

function getGpsForEntry(entryId) {
  if (!navigator.geolocation) {
    alert("Geolocation not supported by this device/browser.");
    return;
  }

  navigator.geolocation.getCurrentPosition(
    (pos) => {
      const p = getCurrentPassage();
      if (!p) return;
      const entry = p.entries.find(e => e.id === entryId);
      if (!entry) return;

      const { latitude, longitude } = pos.coords;
      entry.lat = formatLatFromDecimal(latitude);
      entry.lon = formatLonFromDecimal(longitude);
      savePassages();
      renderLogEntries();
    },
    (err) => {
      console.error("GPS error", err);
      alert("Unable to get GPS position: " + err.message);
    },
    { enableHighAccuracy: true, maximumAge: 30000, timeout: 15000 }
  );
}

// --- Shutdown ------------------------------------------------------

shutdownBtn.addEventListener("click", () => {
  const p = getCurrentPassage();
  if (!p) {
    alert("No passage selected.");
    return;
  }

  if (p.finish.shutdownLogged) {
    alert("Shutdown has already been recorded for this passage.");
    return;
  }

  const now = new Date();
  const timeIso = now.toISOString().slice(0, 16);

  let engineEnd = prompt(
    "Engine hours (end):",
    p.finish.engineHoursEnd || p.plan.engineHoursStart || ""
  );
  if (engineEnd === null) return;
  engineEnd = engineEnd.trim();

  let fuelPct = prompt(
    "Fuel % at end:",
    p.finish.fuelEndPercent || ""
  );
  if (fuelPct === null) return;
  fuelPct = fuelPct.trim();

  let notes = prompt("Summary notes / defects:", p.finish.notes || "");
  if (notes === null) return;
  notes = notes.trim();

  p.finish.engineHoursEnd = engineEnd;
  p.finish.fuelEndPercent = fuelPct;
  p.finish.notes = notes;
  p.finish.shutdownLogged = true;

  const parts = [];
  parts.push("Shutdown / alongside");
  parts.push(`Time: ${timeIso.replace("T", " ")}`);

  const start = parseFloat(p.plan.engineHoursStart || "NaN");
  const endVal = parseFloat(engineEnd || "NaN");
  if (!isNaN(start) && !isNaN(endVal)) {
    const diff = (endVal - start).toFixed(1);
    parts.push(`Engine hours this passage: ${diff} h (from ${start} to ${endVal})`);
  } else if (engineEnd) {
    parts.push(`Engine hours end: ${engineEnd}`);
  }

  if (fuelPct) {
    parts.push(`Fuel end: ${fuelPct}%`);
  }
  if (notes) {
    parts.push(`Notes: ${notes}`);
  }

  const finalEntry = {
    id: "e_" + Date.now(),
    time: timeIso,
    lat: "",
    lon: "",
    course: "",
    speed: "0",
    rpm: "",
    engTP: "",
    waterLog: "",
    groundLog: "",
    fuelUsed: "",
    notes: parts.join(" | ")
  };

  p.entries.unshift(finalEntry);

  savePassages();
  renderLogEntries();
  refreshHomePassageList();
  updatePassageHeader();
  alert("Shutdown recorded and final log entry added.");
});

// --- Log summary panel ---------------------------------------------

function updateLogSummary() {
  if (!logSummaryPanel) return;
  const p = getCurrentPassage();
  if (!p || !p.entries.length || !passageIsShutdown(p)) {
    logSummaryPanel.textContent = "";
    return;
  }

  // Engine hours
  let ehText = "–";
  const start = parseFloat(p.plan.engineHoursStart || "NaN");
  const endVal = parseFloat(p.finish.engineHoursEnd || "NaN");
  if (!isNaN(start) && !isNaN(endVal)) {
    const diff = (endVal - start).toFixed(1);
    ehText = `${diff} h (from ${start} to ${endVal})`;
  }

  // Fuel used: last non-empty
  let fuelUsed = "–";
  const sorted = p.entries.slice().sort((a, b) => (a.time > b.time ? 1 : -1));
  for (let i = sorted.length - 1; i >= 0; i--) {
    const fu = parseFloat(sorted[i].fuelUsed || "NaN");
    if (!isNaN(fu)) {
      fuelUsed = `${fu}`;
      break;
    }
  }

  const fuelStartPct = p.plan.fuelStartPercent || "–";
  const fuelEndPct = p.finish.fuelEndPercent || "–";

  // Final GLog
  let gLog = "–";
  for (let i = sorted.length - 1; i >= 0; i--) {
    if (sorted[i].groundLog) {
      gLog = sorted[i].groundLog;
      break;
    }
  }

  // Duration
  let durationText = "–";
  const times = sorted
    .map(e => e.time)
    .filter(Boolean)
    .map(t => new Date(t));
  if (times.length >= 2) {
    const min = times.reduce((a, b) => (a < b ? a : b));
    const max = times.reduce((a, b) => (a > b ? a : b));
    const ms = max - min;
    if (!isNaN(ms) && ms > 0) {
      const minutes = Math.round(ms / 60000);
      const h = Math.floor(minutes / 60);
      const m = minutes % 60;
      durationText = `${h}h ${m}m`;
    }
  }

  logSummaryPanel.innerHTML = `
    <strong>Summary:</strong>
    Engine hours this passage: ${ehText} |
    Fuel used: ${fuelUsed} |
    Fuel start: ${fuelStartPct}% |
    Fuel end: ${fuelEndPct}% |
    Final GLog: ${gLog} |
    Passage duration: ${durationText}
  `;
}

// --- CSV Export ----------------------------------------------------

function exportCurrentPassageToCsv() {
  const p = getCurrentPassage();
  if (!p) {
    alert("No passage selected.");
    return;
  }

  const date = p.plan.date || p.createdAt.slice(0, 10);
  const from = p.plan.from || "UnknownFrom";
  const to = p.plan.to || "UnknownTo";

  const filename = `${date} ${from} - ${to}.csv`.replace(/[/\\?%*:|"<>]/g, "-");

  let lines = [];

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
    lines.push([
      ts.name || "",
      ts.hw1 || "",
      ts.hw2 || "",
      ts.lw1 || "",
      ts.lw2 || ""
    ].map(quote).join(","));
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
    lines.push([
      ds.date || "",
      ds.fee || "",
      ds.notes || ""
    ].map(quote).join(","));
  });
  lines.push("");

  lines.push(`Engine hours start,${quote(p.plan.engineHoursStart)}`);
  lines.push(`Fuel start %,${quote(p.plan.fuelStartPercent)}`);
  lines.push("");

  lines.push("Log Entries");
  lines.push([
    "Time",
    "Lat",
    "Lon",
    "COG/Heading",
    "Speed (kn)",
    "RPM",
    "Eng T/P",
    "WLog (NM)",
    "GLog (NM)",
    "Fuel used",
    "Notes"
  ].map(quote).join(","));

  p.entries
    .slice()
    .sort((a, b) => (a.time > b.time ? 1 : -1))
    .forEach(e => {
      const water = e.waterLog || e.logReading || "";
      lines.push([
        e.time ? e.time.replace("T", " ") : "",
        e.lat,
        e.lon,
        e.course,
        e.speed,
        e.rpm,
        e.engTP,
        water,
        e.groundLog,
        e.fuelUsed,
        e.notes
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

// --- Load passage into UI -----------------------------------------

function loadPassageIntoUI() {
  const p = getCurrentPassage();
  if (!p) {
    if (planForm) planForm.reset();
    logEntriesContainer.innerHTML = "";
    if (logEmptyMessage) logEmptyMessage.style.display = "block";
    if (planSummaryPanel) planSummaryPanel.innerHTML = "<p>No passage selected.</p>";
    if (logSummaryPanel) logSummaryPanel.textContent = "";
    updatePassageHeader();
    return;
  }
  updatePassageHeader();
  loadPlanIntoForm(p);
  updatePlanSummaryPanel();
  renderLogEntries();
}

// --- Event listeners -----------------------------------------------

homeNewPassageBtn.addEventListener("click", () => {
  if (passages.length > 0) {
    const ok = confirm("Start a new passage? (Current ones will be kept in history.)");
    if (!ok) return;
  }
  createPassage();
  switchToTab("planTab");
});

addEntryBtn.addEventListener("click", addLogEntry);

exportCsvBtn.addEventListener("click", exportCurrentPassageToCsv);

// --- Initial load & SW --------------------------------------------

loadPassages();
loadPorts();
renderPortsDatalist();

// apply saved theme (default day)
const savedTheme = localStorage.getItem(THEME_KEY) || "day";
applyTheme(savedTheme);

refreshHomePassageList();

// If no current passageId but we have passages, pick the latest
if (!currentPassageId && passages.length > 0) {
  currentPassageId = passages[0].id;
}

loadPassageIntoUI();
setLogLayoutMode("split");

// Service worker registration (will only succeed on HTTPS/localhost)
if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => {
    navigator.serviceWorker.register("service-worker.js").catch(err => {
      console.warn("Service worker registration failed", err);
    });
  });
}
