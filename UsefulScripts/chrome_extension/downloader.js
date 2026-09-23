// ============================================================
// Scopus Research PDF Downloader — everything runs in-extension.
// No Consensus (needs a paid API key — intentionally excluded).
// No Sci-Hub (distributes copyrighted material without publisher
// authorization — intentionally excluded).
// No Google Scholar automation (scraping its result pages is
// against Google's Terms of Service — intentionally excluded).
//
// Only legal, key-free, CORS-friendly open-access lookups are used:
//   - Unpaywall API   (https://unpaywall.org)
//   - OpenAlex API    (https://openalex.org)
//   - a direct, verified check of the DOI resolver link itself
// Anything not legally open access is never force-downloaded; it is
// added to the Manual Download list instead, same pattern as the
// IEEE click-through list in the original sample extension.
// ============================================================

const STORAGE_KEY = "scopus_pdf_downloader_state_v1";
const SETTINGS_KEY = "scopus_pdf_downloader_settings_v1";
const BASE_FOLDER = "ScopusPDFs";
const DELAY_MS = 500; // politeness delay between rows / API calls

let jobs = [];       // [{sno, doi, title}]
let results = [];    // per-row outcome, index-aligned with jobs
let manualList = []; // [{sno, title, doi, url, filename, done}]
let cursor = 0;       // next job index to process
let isPaused = false;
let isStopped = false;
let userEmail = "";

// ---------------------------------------------------------------
// DOM refs
// ---------------------------------------------------------------
const emailInput = document.getElementById("emailInput");
const fileInput = document.getElementById("fileInput");
const mappingRow = document.getElementById("mappingRow");
const colSno = document.getElementById("colSno");
const colDoi = document.getElementById("colDoi");
const colTitle = document.getElementById("colTitle");
const startBtn = document.getElementById("startBtn");
const pauseBtn = document.getElementById("pauseBtn");
const stopBtn = document.getElementById("stopBtn");
const exportBtn = document.getElementById("exportBtn");
const exportManualBtn = document.getElementById("exportManualBtn");
const progressBar = document.getElementById("progressBar");
const summary = document.getElementById("summary");
const resultsBody = document.getElementById("resultsBody");
const manualBox = document.getElementById("manualBox");
const manualBody = document.getElementById("manualBody");
const manualCount = document.getElementById("manualCount");
const resumeBox = document.getElementById("resumeBox");
const resumeInfo = document.getElementById("resumeInfo");
const resumeBtn = document.getElementById("resumeBtn");
const discardBtn = document.getElementById("discardBtn");

let rawRows = [];   // parsed sheet rows (array of arrays)
let headerRow = [];

// ---------------------------------------------------------------
// Startup: load saved settings + offer to resume a saved session
// ---------------------------------------------------------------
(async function init() {
  const settings = await storageGet(SETTINGS_KEY);
  if (settings && settings.email) emailInput.value = settings.email;

  const saved = await storageGet(STORAGE_KEY);
  if (saved && Array.isArray(saved.jobs) && saved.jobs.length && saved.cursor < saved.jobs.length) {
    resumeBox.style.display = "block";
    resumeInfo.textContent =
      `${saved.cursor}/${saved.jobs.length} papers already processed. ` +
      `${saved.manualList ? saved.manualList.length : 0} in the manual list so far.`;
  }
})();

resumeBtn.addEventListener("click", async () => {
  const saved = await storageGet(STORAGE_KEY);
  if (!saved) return;
  jobs = saved.jobs;
  results = saved.results;
  manualList = saved.manualList || [];
  cursor = saved.cursor || 0;

  resultsBody.innerHTML = "";
  jobs.forEach((job, i) => addRow(i, job));
  results.forEach((r, i) => applyResultToRow(i, r));
  manualList.forEach((m) => addManualRow(m));
  refreshManualUi();

  resumeBox.style.display = "none";
  startBtn.disabled = false;
  startBtn.textContent = "Resume Downloading";
  summary.textContent = `Resumed: ${cursor}/${jobs.length} already processed.`;
});

discardBtn.addEventListener("click", async () => {
  await storageRemove(STORAGE_KEY);
  resumeBox.style.display = "none";
});

// ---------------------------------------------------------------
// File upload -> parse -> column mapping -> build jobs
// ---------------------------------------------------------------
fileInput.addEventListener("change", async () => {
  const file = fileInput.files[0];
  if (!file) return;

  if (file.name.toLowerCase().endsWith(".csv")) {
    const text = await file.text();
    const lines = text.replace(/\r\n/g, "\n").split("\n").filter((l) => l.trim() !== "");
    rawRows = lines.map(splitCsvLine);
  } else {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const sheetName = wb.SheetNames[0];
    rawRows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: "" });
  }

  if (!rawRows.length) {
    summary.textContent = "Could not read any rows from that file.";
    return;
  }

  headerRow = rawRows[0].map((h) => String(h || "").trim());
  populateColumnSelects();
  mappingRow.style.display = "flex";
  buildJobsFromMapping();
});

[colSno, colDoi, colTitle].forEach((el) => el.addEventListener("change", buildJobsFromMapping));

function populateColumnSelects() {
  const options = headerRow.map((h, i) => `<option value="${i}">${escapeHtml(h || `(column ${i + 1})`)}</option>`).join("");
  colSno.innerHTML = options;
  colDoi.innerHTML = options;
  colTitle.innerHTML = options;

  autoSelect(colSno, ["s.no", "sno", "s no", "serial", "serial no", "serial number", "paperid", "paper id"]);
  autoSelect(colDoi, ["doi"]);
  autoSelect(colTitle, ["title", "article title", "paper title"]);
}

function autoSelect(selectEl, candidates) {
  const idx = headerRow.findIndex((h) => candidates.includes(String(h).trim().toLowerCase()));
  if (idx !== -1) selectEl.value = String(idx);
}

function buildJobsFromMapping() {
  const iSno = Number(colSno.value);
  const iDoi = Number(colDoi.value);
  const iTitle = Number(colTitle.value);
  if ([iSno, iDoi, iTitle].some((v) => Number.isNaN(v))) return;

  jobs = [];
  for (let r = 1; r < rawRows.length; r++) {
    const row = rawRows[r];
    const sno = String(row[iSno] ?? "").trim();
    const doi = normalizeDoi(String(row[iDoi] ?? "").trim());
    const title = String(row[iTitle] ?? "").trim();
    if (!sno && !doi && !title) continue; // skip fully blank rows
    jobs.push({ sno: sno || String(r), doi, title: title || doi || `paper_${r}` });
  }

  results = new Array(jobs.length).fill(null);
  manualList = [];
  cursor = 0;

  resultsBody.innerHTML = "";
  jobs.forEach((job, i) => addRow(i, job));
  manualBody.innerHTML = "";
  refreshManualUi();

  summary.textContent = `Loaded ${jobs.length} row(s). ${jobs.filter((j) => !j.doi).length} have no DOI and will be skipped.`;
  startBtn.disabled = jobs.length === 0;
  startBtn.textContent = "Start Downloading";
}

function normalizeDoi(raw) {
  if (!raw) return "";
  let doi = raw.trim();
  doi = doi.replace(/^https?:\/\/(dx\.)?doi\.org\//i, "");
  doi = doi.replace(/^doi:\s*/i, "");
  return doi.trim();
}

// ---------------------------------------------------------------
// Run controls
// ---------------------------------------------------------------
startBtn.addEventListener("click", async () => {
  userEmail = emailInput.value.trim();
  if (!userEmail || !userEmail.includes("@")) {
    summary.textContent = "Please enter a valid email address (required by Unpaywall's API terms).";
    emailInput.focus();
    return;
  }
  await storageSet(SETTINGS_KEY, { email: userEmail });

  startBtn.disabled = true;
  fileInput.disabled = true;
  pauseBtn.disabled = false;
  stopBtn.disabled = false;
  exportBtn.disabled = false;
  isPaused = false;
  isStopped = false;
  pauseBtn.textContent = "Pause";
  progressBar.style.display = "block";

  await runAll();

  pauseBtn.disabled = true;
  stopBtn.disabled = true;
});

pauseBtn.addEventListener("click", () => {
  isPaused = !isPaused;
  pauseBtn.textContent = isPaused ? "Resume" : "Pause";
  summary.textContent = isPaused ? "Paused." : "Resumed.";
});

stopBtn.addEventListener("click", () => {
  isStopped = true;
  isPaused = false;
  pauseBtn.textContent = "Pause";
  summary.textContent = "Stopping after the current item...";
});

exportBtn.addEventListener("click", () => {
  const header = "sno,title,doi,status,source,local_file,error\n";
  const rows = jobs.map((job, i) => {
    const r = results[i] || {};
    return [
      csvEscape(job.sno), csvEscape(job.title), csvEscape(job.doi),
      csvEscape(r.status || "not processed"), csvEscape(r.source || ""),
      csvEscape(r.file || ""), csvEscape(r.error || ""),
    ].join(",");
  });
  downloadTextFile("Full_Report.csv", header + rows.join("\n"));
});

exportManualBtn.addEventListener("click", () => {
  const header = "sno,title,doi,article_url,suggested_filename,done\n";
  const rows = manualList.map((m) =>
    [csvEscape(m.sno), csvEscape(m.title), csvEscape(m.doi), csvEscape(m.url), csvEscape(m.filename), m.done ? "yes" : "no"].join(",")
  );
  downloadTextFile("Manual_Download_List.csv", header + rows.join("\n"));
});

function downloadTextFile(filename, text) {
  const blob = new Blob([text], { type: "text/csv" });
  const url = URL.createObjectURL(blob);
  chrome.downloads.download({ url, filename, saveAs: false }, () => URL.revokeObjectURL(url));
}

// ---------------------------------------------------------------
// Main processing loop
// ---------------------------------------------------------------
async function runAll() {
  let downloaded = 0, restricted = 0, skipped = 0, failed = 0;

  for (let i = 0; i < cursor; i++) {
    const r = results[i];
    if (!r) continue;
    if (r.status.startsWith("Downloaded") || r.status === "Already downloaded") downloaded++;
    else if (r.status === "Restricted") restricted++;
    else if (r.status === "No DOI") skipped++;
    else failed++;
  }

  for (let i = cursor; i < jobs.length; i++) {
    if (isStopped) break;
    while (isPaused && !isStopped) await sleep(300);
    if (isStopped) break;

    const job = jobs[i];
    setRowStatus(i, "checking...", "status-pending");
    progressBar.value = Math.round((i / jobs.length) * 100);
    summary.textContent = `Processing ${i + 1}/${jobs.length}: ${job.sno}`;

    const result = await processJob(i, job);
    results[i] = result;
    applyResultToRow(i, result);
    cursor = i + 1;

    if (result.status.startsWith("Downloaded") || result.status === "Already downloaded") downloaded++;
    else if (result.status === "Restricted") { restricted++; }
    else if (result.status === "No DOI") skipped++;
    else failed++;

    await persistState();
    await sleep(DELAY_MS);
  }

  progressBar.value = 100;
  const stoppedNote = isStopped ? ` (stopped early at ${cursor}/${jobs.length})` : "";
  summary.textContent =
    `Done${stoppedNote}.\n` +
    `Downloaded: ${downloaded}  |  Restricted (manual list): ${restricted}  |  ` +
    `No DOI / skipped: ${skipped}  |  Other failures: ${failed}`;

  exportManualBtn.disabled = manualList.length === 0;
}

async function processJob(i, job) {
  if (!job.doi) {
    return { status: "No DOI", source: "", file: "", error: "" };
  }

  const filenameBase = `${safeFilename(job.sno)}_${safeFilename(job.title)}.pdf`;

  // --- 0. Already downloaded in a previous run? ---------------
  const existing = await findExistingDownload(filenameBase);
  if (existing) {
    return { status: "Already downloaded", source: existing.source, file: existing.file, error: "" };
  }

  // --- 1. Unpaywall -------------------------------------------
  try {
    const uw = await fetchJson(
      `https://api.unpaywall.org/v2/${encodeURIComponent(job.doi)}?email=${encodeURIComponent(userEmail)}`
    );
    const pdfUrl = extractUnpaywallPdfUrl(uw);
    if (pdfUrl) {
      const path = `${BASE_FOLDER}/Unpaywall/${filenameBase}`;
      const outcome = await attemptDownload(pdfUrl, path);
      if (outcome.ok) return { status: "Downloaded - Unpaywall", source: "Unpaywall", file: path, error: "" };
    }
  } catch (e) { /* fall through to next stage */ }

  // --- 2. OpenAlex ----------------------------------------------
  try {
    const oa = await fetchJson(`https://api.openalex.org/works/https://doi.org/${encodeURIComponent(job.doi)}`);
    const pdfUrl = extractOpenAlexPdfUrl(oa);
    if (pdfUrl) {
      const path = `${BASE_FOLDER}/OpenAlex/${filenameBase}`;
      const outcome = await attemptDownload(pdfUrl, path);
      if (outcome.ok) return { status: "Downloaded - OpenAlex", source: "OpenAlex", file: path, error: "" };
    }
  } catch (e) { /* fall through */ }

  // --- 3. Last-resort direct check of the DOI link itself -------
  // Only kept if it turns out to genuinely be a PDF (verified by MIME
  // type below) — e.g. small/fully-OA journals not yet indexed by
  // Unpaywall or OpenAlex. Anything else (the normal case: an HTML
  // landing/login page) is discarded immediately, never saved.
  try {
    const path = `${BASE_FOLDER}/DirectOpenAccess/${filenameBase}`;
    const outcome = await attemptDownload(`https://doi.org/${encodeURIComponent(job.doi)}`, path);
    if (outcome.ok) return { status: "Downloaded - Direct", source: "DirectOpenAccess", file: path, error: "" };
  } catch (e) { /* fall through */ }

  // --- 4. Nothing legally open — hand to the human --------------
  const manualEntry = {
    sno: job.sno,
    title: job.title,
    doi: job.doi,
    url: `https://doi.org/${job.doi}`,
    filename: filenameBase,
    done: false,
  };
  manualList.push(manualEntry);
  addManualRow(manualEntry);
  refreshManualUi();

  return { status: "Restricted", source: "Manual", file: "", error: "No open-access copy found" };
}

// ---------------------------------------------------------------
// Unpaywall / OpenAlex JSON parsing helpers
// ---------------------------------------------------------------
function extractUnpaywallPdfUrl(data) {
  if (!data) return "";
  const best = data.best_oa_location;
  if (best && best.url_for_pdf) return best.url_for_pdf;
  if (Array.isArray(data.oa_locations)) {
    for (const loc of data.oa_locations) {
      if (loc && loc.url_for_pdf) return loc.url_for_pdf;
    }
  }
  if (best && best.url) return best.url; // may itself be a direct PDF
  return "";
}

function extractOpenAlexPdfUrl(data) {
  if (!data) return "";
  if (data.best_oa_location && data.best_oa_location.pdf_url) return data.best_oa_location.pdf_url;
  if (data.open_access && data.open_access.oa_url) return data.open_access.oa_url;
  if (Array.isArray(data.locations)) {
    for (const loc of data.locations) {
      if (loc && loc.is_oa && loc.pdf_url) return loc.pdf_url;
    }
  }
  return "";
}

async function fetchJson(url) {
  const resp = await fetch(url, { headers: { Accept: "application/json" } });
  if (!resp.ok) return null;
  try { return await resp.json(); } catch { return null; }
}

// ---------------------------------------------------------------
// Download + verification (same "was it actually a PDF?" check
// used by the original sample extension's IEEE/MDPI flow)
// ---------------------------------------------------------------
async function attemptDownload(url, filename) {
  const downloadId = await new Promise((resolve) => {
    chrome.downloads.download({ url, filename, saveAs: false, conflictAction: "overwrite" }, (id) => {
      resolve(chrome.runtime.lastError || !id ? null : id);
    });
  });
  if (downloadId === null) return { ok: false, msg: chrome.runtime.lastError?.message || "failed to start" };
  return await waitForDownloadResult(downloadId);
}

function waitForDownloadResult(downloadId) {
  return new Promise((resolve) => {
    let settled = false;
    const finish = async (result) => {
      if (settled) return;
      settled = true;
      chrome.downloads.onChanged.removeListener(listener);
      resolve(result);
    };
    const listener = (delta) => {
      if (delta.id !== downloadId) return;
      if (delta.state && delta.state.current && delta.state.current !== "in_progress") {
        finalizeDownload(downloadId).then(finish);
      }
    };
    chrome.downloads.onChanged.addListener(listener);
    setTimeout(async () => {
      if (settled) return;
      finish(await finalizeDownload(downloadId));
    }, 20000);
  });
}

async function finalizeDownload(downloadId) {
  const items = await new Promise((res) => chrome.downloads.search({ id: downloadId }, res));
  const item = items && items[0];
  if (!item) return { ok: false, msg: "download not found after completion" };

  if (item.state === "interrupted") {
    return { ok: false, msg: `interrupted: ${item.error || "unknown error"}` };
  }

  if (item.state === "complete") {
    const mime = (item.mime || "").toLowerCase();
    const looksLikeHtml = mime.includes("html") || mime.includes("text/plain");
    if (looksLikeHtml || (!mime.includes("pdf") && mime !== "")) {
      // Not actually a PDF (login wall, landing page, etc.) — discard.
      await new Promise((res) => chrome.downloads.removeFile(downloadId, () => res()));
      await new Promise((res) => chrome.downloads.erase({ id: downloadId }, res));
      return { ok: false, msg: "not a PDF (likely a paywall/login page)" };
    }
    return { ok: true, msg: "ok" };
  }
  return { ok: false, msg: `unexpected state: ${item.state}` };
}

async function findExistingDownload(filenameBase) {
  const sources = ["Unpaywall", "OpenAlex", "DirectOpenAccess"];
  for (const source of sources) {
    const path = `${BASE_FOLDER}/${source}/${filenameBase}`;
    const items = await new Promise((res) =>
      chrome.downloads.search({ filenameRegex: escapeRegex(path) + "$", state: "complete", limit: 1 }, res)
    );
    if (items && items.length) return { source, file: path };
  }
  return null;
}

// ---------------------------------------------------------------
// UI helpers
// ---------------------------------------------------------------
function addRow(i, job) {
  const tr = document.createElement("tr");
  tr.id = `row-${i}`;
  tr.innerHTML = `
    <td>${i + 1}</td>
    <td>${escapeHtml(job.sno)}</td>
    <td style="max-width:260px;">${escapeHtml(job.title).slice(0, 90)}</td>
    <td style="max-width:160px; overflow-wrap:break-word;">${escapeHtml(job.doi)}</td>
    <td class="src">-</td>
    <td class="status-pending status-cell">pending</td>
  `;
  resultsBody.appendChild(tr);
}

function setRowStatus(i, text, cls) {
  const tr = document.getElementById(`row-${i}`);
  if (!tr) return;
  const cell = tr.querySelector(".status-cell");
  cell.textContent = text;
  cell.className = `status-cell ${cls}`;
}

function applyResultToRow(i, result) {
  const tr = document.getElementById(`row-${i}`);
  if (!tr || !result) return;
  tr.querySelector(".src").textContent = result.source || "-";
  let cls = "status-pending";
  if (result.status.startsWith("Downloaded") || result.status === "Already downloaded") cls = "status-success";
  else if (result.status === "Restricted") cls = "status-manual";
  else if (result.status === "No DOI") cls = "status-skip";
  else cls = "status-failed";
  setRowStatus(i, result.status, cls);
}

function addManualRow(entry) {
  const tr = document.createElement("tr");
  tr.innerHTML = `
    <td>${manualBody.children.length + 1}</td>
    <td>${escapeHtml(entry.sno)}</td>
    <td style="max-width:220px;">${escapeHtml(entry.title).slice(0, 90)}</td>
    <td style="max-width:140px; overflow-wrap:break-word;">${escapeHtml(entry.doi)}</td>
    <td><a class="open-link" href="${escapeHtml(entry.url)}" target="_blank" rel="noopener">Open article &rarr;</a></td>
    <td><code>${escapeHtml(entry.filename)}</code>
        <button class="copy-btn" type="button">copy</button></td>
    <td><input type="checkbox" ${entry.done ? "checked" : ""}></td>
  `;
  tr.querySelector(".copy-btn").addEventListener("click", () => {
    navigator.clipboard.writeText(entry.filename).catch(() => {});
  });
  tr.querySelector('input[type="checkbox"]').addEventListener("change", (e) => {
    entry.done = e.target.checked;
    persistState();
  });
  manualBody.appendChild(tr);
}

function refreshManualUi() {
  manualBox.style.display = manualList.length ? "block" : "none";
  manualCount.textContent = String(manualList.length);
  exportManualBtn.disabled = manualList.length === 0;
}

function escapeHtml(s) {
  return String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}
function escapeRegex(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); }
function safeFilename(s) { return String(s || "").replace(/[\\/:*?"<>|]/g, "_").trim().slice(0, 120); }
function csvEscape(s) {
  s = String(s ?? "");
  if (s.includes(",") || s.includes('"') || s.includes("\n")) return '"' + s.replace(/"/g, '""') + '"';
  return s;
}
function sleep(ms) { return new Promise((r) => setTimeout(r, ms)); }

function splitCsvLine(line) {
  const result = [];
  let cur = "", inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') inQuotes = !inQuotes;
    else if (ch === "," && !inQuotes) { result.push(cur); cur = ""; }
    else cur += ch;
  }
  result.push(cur);
  return result;
}

// ---------------------------------------------------------------
// chrome.storage.local helpers (session persistence / resume)
// ---------------------------------------------------------------
function storageGet(key) {
  return new Promise((resolve) => chrome.storage.local.get([key], (res) => resolve(res[key])));
}
function storageSet(key, value) {
  return new Promise((resolve) => chrome.storage.local.set({ [key]: value }, resolve));
}
function storageRemove(key) {
  return new Promise((resolve) => chrome.storage.local.remove([key], resolve));
}
async function persistState() {
  await storageSet(STORAGE_KEY, { jobs, results, manualList, cursor });
}
