// --- Constants & state ---------------------------------------------

const STORAGE_KEY = "steeler_logbook_passages_v5";
const THEME_KEY   = "steeler_logbook_theme_v1";
const PORTS_KEY   = "steeler_logbook_ports_v1";

let passages = [];
let currentPassageId = null;
let knownPorts = [];
let recentPorts = [];
const PORTS_RECENT_LIMIT = 20;




function portName(p){
  return (typeof p === "string") ? p : (p && typeof p === "object" ? (p.name || "") : "");
}
function portHasCoords(p){
  return p && typeof p === "object" && !isNaN(p.lat) && !isNaN(p.lon);
}
function findPortItemByName(name){
  const n = (name || "").trim();
  if (!n) return null;
  return knownPorts.find(p => portName(p) === n) || null;
}
function upsertPortItem(name, lat=null, lon=null){
  const n = (name || "").trim();
  if (!n) return;
  const existingIdx = knownPorts.findIndex(p => portName(p) === n);
  if (existingIdx >= 0){
    const existing = knownPorts[existingIdx];
    if (lat != null && lon != null){
      knownPorts[existingIdx] = { name: n, lat: Number(lat), lon: Number(lon) };
    } else {
      knownPorts[existingIdx] = existing;
    }
  } else {
    knownPorts.push((lat != null && lon != null) ? { name: n, lat: Number(lat), lon: Number(lon) } : n);
  }
  knownPorts.sort((a,b) => portName(a).localeCompare(portName(b)));
}

// --- Port autocomplete + management --------------------------------

function getPortSuggestions(query) {
  const q = (query || "").trim().toLowerCase();
  let list;

  if (!q) {
    // show MRU first, then fall back to alphabetical if MRU empty
    list = (recentPorts && recentPorts.length ? recentPorts.slice() : knownPorts.slice());
  } else {
    list = knownPorts.filter(p => portName(p).toLowerCase().includes(q));
    // prefer starts-with matches
    list.sort((a, b) => {
      const aStart = a.toLowerCase().startsWith(q) ? 0 : 1;
      const bStart = b.toLowerCase().startsWith(q) ? 0 : 1;
      if (aStart !== bStart) return aStart - bStart;
      return a.localeCompare(b);
    });
  }

  // ensure unique
  const seen = new Set();
  const out = [];
  for (const p of list) {
    const name = portName(p);
    if (!seen.has(name)) {
      seen.add(name);
      out.push(name);
    }
    if (out.length >= 6) break;
  }
  return out;
}

function renderPortSuggestBox(inputEl, boxEl) {
  if (!inputEl || !boxEl) return;

  const suggestions = getPortSuggestions(inputEl.value);
  boxEl.innerHTML = "";

  if (!suggestions.length) {
    boxEl.classList.add("hidden");
    return;
  }

  suggestions.forEach(name => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "port-suggest-item";
    btn.textContent = name;
    btn.addEventListener("mousedown", (e) => {
      // mousedown so we beat blur
      e.preventDefault();
      inputEl.value = name;
      rememberPort(name);
      boxEl.classList.add("hidden");
      // trigger any bound input handler
      inputEl.dispatchEvent(new Event("input", { bubbles: true }));
    });
    boxEl.appendChild(btn);
  });

  boxEl.classList.remove("hidden");
}

function setupSinglePortAutocomplete(inputId, boxId) {
  const inputEl = document.getElementById(inputId);
  const boxEl = document.getElementById(boxId);
  if (!inputEl || !boxEl) return;

  const show = () => renderPortSuggestBox(inputEl, boxEl);
  inputEl.addEventListener("input", show);
  inputEl.addEventListener("focus", show);
  inputEl.addEventListener("blur", () => {
    // allow click selection
    setTimeout(() => boxEl.classList.add("hidden"), 150);
  });

  // Escape hides
  inputEl.addEventListener("keydown", (e) => {
    if (e.key === "Escape") boxEl.classList.add("hidden");
  });
}

function setupPortAutocomplete() {
  setupSinglePortAutocomplete("planFrom", "planFromSuggest");
  setupSinglePortAutocomplete("planTo", "planToSuggest");
}

function deletePort(name) {
  const trimmed = (name || "").trim();
  if (!trimmed) return;
  knownPorts = knownPorts.filter(p => portName(p) !== trimmed);
  recentPorts = recentPorts.filter(p => p !== trimmed);
  savePorts();
  refreshPortUI();
}


function renderPortsManagerList() {
  const list = document.getElementById("portsManagerList");
  if (!list) return;
  list.innerHTML = "";

  const items = knownPorts.slice().sort((a, b) => portName(a).localeCompare(portName(b)));
  if (!items.length) {
    const empty = document.createElement("div");
    empty.className = "ports-empty";
    empty.textContent = "No saved ports yet.";
    list.appendChild(empty);
    return;
  }

  items.forEach(item => {
    const name = portName(item);

    const row = document.createElement("div");
    row.className = "ports-row";

    const left = document.createElement("div");
    left.className = "ports-left";

    const label = document.createElement("div");
    label.className = "ports-name";
    label.textContent = name;

    const coords = document.createElement("div");
    coords.className = "ports-coords";

    const latInput = document.createElement("input");
    latInput.type = "number";
    latInput.inputMode = "decimal";
    latInput.step = "0.0001";
    latInput.placeholder = "Lat";
    latInput.className = "ports-coord-input";
    latInput.value = (item && typeof item === "object" && item.lat != null) ? item.lat : "";

    const lonInput = document.createElement("input");
    lonInput.type = "number";
    lonInput.inputMode = "decimal";
    lonInput.step = "0.0001";
    lonInput.placeholder = "Lon";
    lonInput.className = "ports-coord-input";
    lonInput.value = (item && typeof item === "object" && item.lon != null) ? item.lon : "";

    const saveBtn = document.createElement("button");
    saveBtn.type = "button";
    saveBtn.className = "ports-mini";
    saveBtn.textContent = "Save coords";
    saveBtn.addEventListener("click", () => {
      const parsed = parseLatLon(latInput.value, lonInput.value);
      if (!parsed) {
        alert("Please enter valid decimal lat and lon (e.g. 50.757, -1.545).");
        return;
      }
      upsertPortItem(name, parsed.lat, parsed.lon);
      savePorts();
      renderPortsManagerList();
      autoComputeSunriseSetForCurrent();
    });

    const lookupBtn = document.createElement("button");
    lookupBtn.type = "button";
    lookupBtn.className = "ports-mini";
    lookupBtn.textContent = "Lookup";
    lookupBtn.addEventListener("click", async () => {
      try {
        const q = encodeURIComponent(name + " harbour");
        const viewbox = "-6.5,52.5,2.5,49.0";
        const url = `https://nominatim.openstreetmap.org/search?format=jsonv2&limit=1&viewbox=${viewbox}&bounded=1&q=${q}`;
        const res = await fetch(url, { headers: { "Accept": "application/json" } });
        if (!res.ok) throw new Error("Lookup failed");
        const data = await res.json();
        if (!data || !data[0]) {
          alert("No match found. Try manual lat/lon.");
          return;
        }
        const lat = parseFloat(data[0].lat);
        const lon = parseFloat(data[0].lon);
        if (isNaN(lat) || isNaN(lon)) {
          alert("Lookup returned invalid coordinates.");
          return;
        }
        if(!saneForSteeler(lat, lon)){
          alert("Lookup result looks too far away for a UK/Channel port. Please try manual entry or adjust the port name.");
          return;
        }
        latInput.value = lat.toFixed(6);
        lonInput.value = lon.toFixed(6);
      } catch (e) {
        console.error(e);
        alert("Could not look up that port (offline or blocked). You can enter lat/lon manually.");
      }
    });

    coords.appendChild(latInput);
    coords.appendChild(lonInput);

    const dmm = document.createElement("div");
    dmm.className = "ports-dmm";
    const latV = (item && typeof item === "object" && item.lat != null) ? item.lat : NaN;
    const lonV = (item && typeof item === "object" && item.lon != null) ? item.lon : NaN;
    dmm.textContent = (isNaN(latV)||isNaN(lonV)) ? "" : formatDMM(latV, lonV);
    coords.appendChild(dmm);

    coords.appendChild(saveBtn);
    coords.appendChild(lookupBtn);

    left.appendChild(label);
    left.appendChild(coords);

    const right = document.createElement("div");
    right.className = "ports-right";

    const del = document.createElement("button");
    del.type = "button";
    del.className = "ports-delete";
    del.textContent = "Remove";
    del.addEventListener("click", () => deletePort(name));

    right.appendChild(del);

    row.appendChild(left);
    row.appendChild(right);
    list.appendChild(row);
  });
}


function setupPortsManagerModal() {
  const openBtn = document.getElementById("managePortsBtn");
  const modal = document.getElementById("portsModal");
  const closeBtn = document.getElementById("portsModalClose");
  const overlay = document.getElementById("portsModalOverlay");

  if (!openBtn || !modal) return;

  const open = () => {
    renderPortsManagerList();
    modal.classList.remove("hidden");
  };
  const close = () => modal.classList.add("hidden");

  openBtn.addEventListener("click", open);
  if (closeBtn) closeBtn.addEventListener("click", close);
  if (overlay) overlay.addEventListener("click", close);
}

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
    if (!raw) {
      knownPorts = [];
      recentPorts = [];
      return;
    }
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed)) {
      knownPorts = parsed;
      recentPorts = [];
    } else if (parsed && typeof parsed === "object") {
      knownPorts = Array.isArray(parsed.all) ? parsed.all : [];
      recentPorts = Array.isArray(parsed.recent) ? parsed.recent : [];
    } else {
      knownPorts = [];
      recentPorts = [];
    }
  } catch {
    knownPorts = [];
    recentPorts = [];
  }
}

function savePorts() {
  try {
    const payload = { all: knownPorts, recent: recentPorts };
    localStorage.setItem(PORTS_KEY, JSON.stringify(payload));
  } catch (e) {
    console.warn("Failed to save ports", e);
  }
}

function rememberPort(name) {
  const trimmed = (name || "").trim();
  if (!trimmed) return;

  // Add to master list
  if (!knownPorts.includes(trimmed)) {
    knownPorts.push(trimmed);
    knownPorts.sort((a,b)=>String(a?.name ?? a ?? "").localeCompare(String(b?.name ?? b ?? ""), undefined, {sensitivity:"base"}));
  }

  // Update MRU list (most recent first)
  recentPorts = recentPorts.filter(p => p !== trimmed);
  recentPorts.unshift(trimmed);
  if (recentPorts.length > PORTS_RECENT_LIMIT) recentPorts.length = PORTS_RECENT_LIMIT;

  savePorts();
  refreshPortUI();
}

// --- Small helpers -------------------------------------------------

// --- Coordinate formatting/parsing + sanity checks (CL-073) --------
function formatDMM(lat, lon){
  function one(val, isLat){
    const hemi = isLat ? (val>=0 ? "N" : "S") : (val>=0 ? "E" : "W");
    const a = Math.abs(val);
    const deg = Math.floor(a);
    const min = (a - deg) * 60;
    // 3 decimals on minutes
    return `${deg}°${min.toFixed(3)}'${hemi}`;
  }
  if (isNaN(lat) || isNaN(lon)) return "";
  return one(lat,true) + "  " + one(lon,false);
}

function parseCoordPart(s, isLat){
  if (!s) return NaN;
  const t = String(s).trim().toUpperCase();
  // decimal
  if (/^-?\d+(\.\d+)?$/.test(t)) return parseFloat(t);

  // DMM forms like 50°45.123'N or 50 45.123 N
  const m = t.match(/^(\d{1,3})\s*(?:°|\s)\s*(\d{1,2}(?:\.\d+)?)\s*(?:'|\s)?\s*([NSEW])$/);
  if (!m) return NaN;
  const deg = parseInt(m[1],10);
  const mins = parseFloat(m[2]);
  const hemi = m[3];
  let val = deg + (mins/60);
  if (hemi === "S" || hemi === "W") val *= -1;
  // basic range sanity
  if (isLat && (val < -90 || val > 90)) return NaN;
  if (!isLat && (val < -180 || val > 180)) return NaN;
  return val;
}

function parseLatLon(latStr, lonStr){
  const lat = parseCoordPart(latStr,true);
  const lon = parseCoordPart(lonStr,false);
  if (!isNaN(lat) && !isNaN(lon)) return {lat, lon};
  return null;
}

// Haversine distance in km
function distanceKm(lat1, lon1, lat2, lon2){
  const R = 6371;
  const rad = Math.PI/180;
  const dLat = (lat2-lat1)*rad;
  const dLon = (lon2-lon1)*rad;
  const a = Math.sin(dLat/2)**2 + Math.cos(lat1*rad)*Math.cos(lat2*rad)*Math.sin(dLon/2)**2;
  return 2*R*Math.asin(Math.sqrt(a));
}

// UK-centric sanity check for STEELER usage: reject lookups > 1500km from Solent-ish default.
function saneForSteeler(lat, lon){
  const refLat = 50.76;   // Lymington-ish
  const refLon = -1.54;
  const km = distanceKm(refLat, refLon, lat, lon);
  return km <= 1500; // generous: covers UK + near continent
}

// --- Port coordinate helpers (offline-first) -----------------------------

function normalisePortQuery(name){
  return (name || "")
    .toString()
    .trim()
    .replace(/\s+/g, " ")
    .replace(/[,]/g, "")
    .replace(/\b(harbour|harbor|marina|port)\b/ig, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function getPortCoords(name){
  const q = normalisePortQuery(name);
  if (!q) return null;

  // 1) exact match against stored knownPorts (objects only)
  for (const p of (knownPorts || [])){
    if (p && typeof p === "object" && p.lat != null && p.lon != null){
      const pn = normalisePortQuery(p.name || "");
      if (pn && pn === q){
        return { name: p.name || name, lat: Number(p.lat), lon: Number(p.lon) };
      }
    }
  }

  // 2) offline baked-in UK/Channel micro-database (marine-sane only)
  const OFFLINE_PORTS = {
    "lymington": {lat:50.758, lon:-1.540},
    "cowes": {lat:50.763, lon:-1.297},
    "yarmouth": {lat:50.705, lon:-1.498},
    "portsmouth": {lat:50.802, lon:-1.109},
    "gosport": {lat:50.795, lon:-1.125},
    "port solent": {lat:50.845, lon:-1.138},
    "poole": {lat:50.714, lon:-1.985},
    "weymouth": {lat:50.613, lon:-2.455},
    "dartmouth": {lat:50.351, lon:-3.579},
    "salcombe": {lat:50.237, lon:-3.769},
    "plymouth": {lat:50.366, lon:-4.143},
    "falmouth": {lat:50.155, lon:-5.073},
    "fowey": {lat:50.336, lon:-4.638},
    "padstow": {lat:50.544, lon:-4.936},
    "st vaast": {lat:49.590, lon:-1.267},
    "cherbourg": {lat:49.642, lon:-1.622},
    "st helier": {lat:49.183, lon:-2.105},
    "st malo": {lat:48.649, lon:-2.025},
    "dunkerque": {lat:51.049, lon:2.377},
    "calais": {lat:50.958, lon:1.851},
    "dieppe": {lat:49.922, lon:1.077},
    "le havre": {lat:49.491, lon:0.107},
    "honfleur": {lat:49.419, lon:0.233},
    "deauville": {lat:49.363, lon:0.078},
    "brighton": {lat:50.820, lon:-0.142},
    "newhaven": {lat:50.793, lon:0.055},
    "eastbourne": {lat:50.770, lon:0.293},
    "chichester": {lat:50.814, lon:-0.876},
    "langstone": {lat:50.824, lon:-1.012}
  };

  if (OFFLINE_PORTS[q]) return { name, lat: OFFLINE_PORTS[q].lat, lon: OFFLINE_PORTS[q].lon };

  // 3) fuzzy: allow prefix match for e.g. "Chichester Harbour"
  const keys = Object.keys(OFFLINE_PORTS);
  const hit = keys.find(k => q === k || q.startsWith(k + " ") || k.startsWith(q + " "));
  if (hit) return { name, lat: OFFLINE_PORTS[hit].lat, lon: OFFLINE_PORTS[hit].lon };

  return null;
}

// --- Sunrise / sunset calculation (NOAA approximation, offline) ----------

function parseISODate(iso){
  // expects YYYY-MM-DD from <input type="date">
  const m = /^\s*(\d{4})-(\d{2})-(\d{2})\s*$/.exec(iso || "");
  if (!m) return null;
  const y = Number(m[1]), mo = Number(m[2]), d = Number(m[3]);
  if (!y || !mo || !d) return null;
  return { y, mo, d };
}

function dayOfYear(y, mo, d){
  const dt = new Date(Date.UTC(y, mo-1, d));
  const start = new Date(Date.UTC(y, 0, 1));
  return Math.floor((dt - start) / 86400000) + 1;
}

function degToRad(x){ return x * Math.PI / 180; }
function radToDeg(x){ return x * 180 / Math.PI; }

function calcSunTimeUtcMinutes(isRise, y, mo, d, lat, lon){
  // Based on NOAA solar calculations (approx). Returns minutes after 00:00 UTC.
  const N = dayOfYear(y, mo, d);
  const lngHour = lon / 15;

  const t = N + ((isRise ? 6 : 18) - lngHour) / 24;

  const M = (0.9856 * t) - 3.289;

  let L = M + (1.916 * Math.sin(degToRad(M))) + (0.020 * Math.sin(degToRad(2*M))) + 282.634;
  L = (L % 360 + 360) % 360;

  let RA = radToDeg(Math.atan(0.91764 * Math.tan(degToRad(L))));
  RA = (RA % 360 + 360) % 360;

  // Quadrant adjustment
  const Lquadrant  = Math.floor(L / 90) * 90;
  const RAquadrant = Math.floor(RA / 90) * 90;
  RA = RA + (Lquadrant - RAquadrant);
  RA = RA / 15;

  const sinDec = 0.39782 * Math.sin(degToRad(L));
  const cosDec = Math.cos(Math.asin(sinDec));

  // Official zenith for sunrise/sunset
  const zenith = 90.833;

  const cosH = (Math.cos(degToRad(zenith)) - (sinDec * Math.sin(degToRad(lat)))) / (cosDec * Math.cos(degToRad(lat)));
  if (cosH > 1 || cosH < -1) return null; // polar day/night edge cases

  let H = isRise ? (360 - radToDeg(Math.acos(cosH))) : radToDeg(Math.acos(cosH));
  H = H / 15;

  const T = H + RA - (0.06571 * t) - 6.622;
  let UT = T - lngHour;
  UT = (UT % 24 + 24) % 24;

  return Math.round(UT * 60);
}

function formatTimeEuropeLondon(dateUtc){
  try{
    return new Intl.DateTimeFormat("en-GB", {
      hour: "2-digit", minute: "2-digit",
      hour12: false,
      timeZone: "Europe/London"
    }).format(dateUtc);
  }catch{
    // fallback: local
    return dateUtc.toLocaleTimeString("en-GB", {hour:"2-digit", minute:"2-digit", hour12:false});
  }
}

function calcSunTimes(isoDate, lat, lon){
  const p = parseISODate(isoDate);
  if (!p) return null;
  const riseMin = calcSunTimeUtcMinutes(true, p.y, p.mo, p.d, lat, lon);
  const setMin  = calcSunTimeUtcMinutes(false, p.y, p.mo, p.d, lat, lon);
  if (riseMin == null || setMin == null) return null;

  const riseUtc = new Date(Date.UTC(p.y, p.mo-1, p.d, 0, 0, 0) + riseMin*60000);
  const setUtc  = new Date(Date.UTC(p.y, p.mo-1, p.d, 0, 0, 0) + setMin*60000);

  return {
    sunrise: formatTimeEuropeLondon(riseUtc),
    sunset:  formatTimeEuropeLondon(setUtc)
  };
}




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
  closePortsManagerModal();

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



async function ensurePortCoords(name){
  const n = (name || "").trim();
  if(!n) return null;

  // already stored?
  const existing = getPortCoords(n);
  if (existing) return existing;

  // try online lookup (if available)
  try{
    if (!navigator.onLine) return null;

    // Add a lightweight bias towards UK/Channel area using viewbox
    const q = encodeURIComponent(n + " harbour");
    const viewbox = "-6.5,52.5,2.5,49.0"; // approx UK south & Channel
    const url = `https://nominatim.openstreetmap.org/search?format=jsonv2&limit=1&viewbox=${viewbox}&bounded=1&q=${q}`;
    const res = await fetch(url, { headers: { "Accept":"application/json" } });
    if(!res.ok) return null;
    const data = await res.json();
    if(!data || !data[0]) return null;

    const lat = parseFloat(data[0].lat);
    const lon = parseFloat(data[0].lon);
    if(isNaN(lat) || isNaN(lon)) return null;

    // sanity check (prevents Chichester, TX etc)
    if(!saneForSteeler(lat, lon)){
      console.warn("Lookup failed sanity check for", n, lat, lon, data[0]);
      return null;
    }

    upsertPortItem(n, lat, lon);
    savePorts();
    return {name:n, lat, lon};
  }catch(e){
    console.warn("Port lookup failed:", e);
    return null;
  }
}



async function autoComputeSunriseSetForCurrent(){
  const p = getCurrentPassage();
  if (!p) return;

  const date = (p.plan.date || planDate?.value || "").trim();
  const from = (p.plan.from || planFrom?.value || "").trim();
  const to   = (p.plan.to   || planTo?.value || "").trim();

  if (!date || !from) return;

  const origin = await ensurePortCoords(from);
  const dest = isLocalDestination(to) ? origin : (to ? await ensurePortCoords(to) : null);

  if (!origin) return;

  const sunOrigin = calcSunTimes(date, origin.lat, origin.lon);
  if (!sunOrigin) return;

  let sunset = sunOrigin.sunset;
  if (dest && dest !== origin){
    const sunDest = calcSunTimes(date, dest.lat, dest.lon);
    if (sunDest && sunDest.sunset) sunset = sunDest.sunset;
  }

  const val = `${sunOrigin.sunrise} / ${sunset}`;
  p.plan.sunriseSet = val;

  if (planSunriseSet) planSunriseSet.value = val;

  savePassages();
  updatePassageHeader();
  updatePlanSummaryPanel();
}


// --- Ports datalist -----------------------------------------------

function refreshPortUI() {
  // Hook for any UI elements that depend on the port list.
  // (Autocomplete + Manage Ports modal)
  renderPortsManagerList();
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
      knownPorts: { all: knownPorts, recent: recentPorts },
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
      if (!Array.isArray(obj.data.passages) || !obj.data.knownPorts) {
        alert("Backup file is missing expected data.");
        return;
      }
      const ok = confirm("Restore backup? This will REPLACE the current logbook data on this device.");
      if (!ok) return;

      passages = obj.data.passages;

      // Support both legacy (array) and current (object with {all,recent}) port backup formats (CL-071)
      const portsPayload = obj.data.knownPorts;
      if (Array.isArray(portsPayload)) {
        knownPorts = portsPayload;
        recentPorts = portsPayload.slice(0, 6);
      } else {
        knownPorts = Array.isArray(portsPayload.all) ? portsPayload.all : [];
        recentPorts = Array.isArray(portsPayload.recent) ? portsPayload.recent : [];
      }

      localStorage.setItem(STORAGE_KEY, JSON.stringify(passages));
      localStorage.setItem(PORTS_KEY, JSON.stringify({ all: knownPorts, recent: recentPorts }));

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


let sunSyncTimer = null;
function scheduleAutoSunSync(){
  clearTimeout(sunSyncTimer);
  sunSyncTimer = setTimeout(() => {
    const p = getCurrentPassage();
    if (!p) return;
    p.plan.date = planDate.value;
    p.plan.from = planFrom.value.trim();
    p.plan.to   = planTo.value.trim();
    autoComputeSunriseSetForCurrent();
  }, 180);
}
planDate.addEventListener("input", scheduleAutoSunSync);
planFrom.addEventListener("input", scheduleAutoSunSync);
planTo.addEventListener("input", scheduleAutoSunSync);

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

function addSpecialEntry(noteText, notesOverride = null) {
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
    notes: (notesOverride !== null ? notesOverride : (noteText || ""))
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

  const deleted = p.entries[idx];

  const ok = confirm("Delete this log entry?");
  if (!ok) return;

  p.entries.splice(idx, 1);

  // If the Shutdown entry was deleted, clear the shutdown flag so a new one can be added (CL-070)
  if (
    deleted &&
    typeof deleted.notes === "string" &&
    deleted.notes.toLowerCase().startsWith("shutdown")
  ) {
    if (!p.finish) p.finish = {};
    p.finish.shutdownLogged = false;
    // Clear finish fields that are only meaningful after shutdown
    p.finish.finishedAt = null;
    p.finish.engineHoursEnd = null;
    p.finish.fuelEndPercent = null;
  }

  // Keep shutdown flag consistent even if something odd happens
  if (p.finish) {
    const hasShutdown = (p.entries || []).some(e => typeof e.notes === "string" && e.notes.toLowerCase().startsWith("shutdown"));
    p.finish.shutdownLogged = !!hasShutdown;
  }
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

      const startBits = [];
      if (eh) startBits.push(`EH ${eh}`);
      if (fu) startBits.push(`Fuel ${fu}%`);
      const startNotes = startBits.length ? `Engine start — ${startBits.join(" | ")}` : "Engine start";
      addSpecialEntry("Engine start", startNotes);
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

      // Include key figures in the notes for quick scanability (CL-066)
      const ehEnd = p.finish.engineHoursEnd;
      const fuelEnd = p.finish.fuelEndPercent;
      const shutBits = [];
      if (ehEnd) shutBits.push(`EH ${ehEnd}`);
      if (fuelEnd) shutBits.push(`Fuel ${fuelEnd}%`);
      const shutPrefix = shutBits.length ? `Shutdown / alongside — ${shutBits.join(" | ")}` : "Shutdown / alongside";
      const note = p.finish.notes ? `${shutPrefix} — ${p.finish.notes}` : shutPrefix;

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
setupPortAutocomplete();
setupPortsManagerModal();
refreshPortUI();
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

function closePortsManagerModal(){
  const modal = document.getElementById("portsModal");
  if (modal) modal.classList.add("hidden");
}
