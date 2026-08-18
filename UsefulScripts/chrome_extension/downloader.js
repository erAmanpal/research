let jobs = [];
let results = [];
let isPaused = false;
let isStopped = false;

const fileInput = document.getElementById("csvFile");
const startBtn = document.getElementById("startBtn");
const pauseBtn = document.getElementById("pauseBtn");
const stopBtn = document.getElementById("stopBtn");
const exportBtn = document.getElementById("exportBtn");
const resultsBody = document.getElementById("resultsBody");
const summary = document.getElementById("summary");
const ieeeBox = document.getElementById("ieeeBox");
const ieeeLinksDiv = document.getElementById("ieeeLinks");

fileInput.addEventListener("change", async () => {
  const file = fileInput.files[0];
  if (!file) return;
  const text = await file.text();
  jobs = parseCsvToJobs(text);
  resultsBody.innerHTML = "";
  results = [];
  jobs.forEach((job, i) => addRow(i, job));
  summary.textContent = `Loaded ${jobs.length} link(s) from CSV.`;
  startBtn.disabled = jobs.length === 0;
});

startBtn.addEventListener("click", async () => {
  startBtn.disabled = true;
  fileInput.disabled = true;
  pauseBtn.disabled = false;
  stopBtn.disabled = false;
  isPaused = false;
  isStopped = false;
  pauseBtn.textContent = "Pause";
  await runAll();
  pauseBtn.disabled = true;
  stopBtn.disabled = true;
  exportBtn.disabled = false;
});

pauseBtn.addEventListener("click", () => {
  isPaused = !isPaused;
  pauseBtn.textContent = isPaused ? "Resume" : "Pause";
  summary.textContent = isPaused ? "Paused." : "Resumed.";
});

stopBtn.addEventListener("click", () => {
  isStopped = true;
  isPaused = false; // in case it was paused, let the loop notice isStopped and exit
  pauseBtn.textContent = "Pause";
  summary.textContent = "Stopping after the current item...";
});

exportBtn.addEventListener("click", () => {
  const header = "paper_id,title,url,resolved_url,status,file\n";
  const rows = results.map(r => [
    csvEscape(r.paper_id), csvEscape(r.title), csvEscape(r.url),
    csvEscape(r.resolved_url), csvEscape(r.status), csvEscape(r.file)
  ].join(","));
  const blob = new Blob([header + rows.join("\n")], { type: "text/csv" });
  const url = URL.createObjectURL(blob);
  chrome.downloads.download({ url, filename: "download_report.csv", saveAs: false }, () => {
    URL.revokeObjectURL(url);
  });
});

// ---------------------------------------------------------------------
// CSV parsing - supports both:
//   extracted_pdf_links.csv : paper_id, title, source_column, url
//   scholar_found_urls.csv  : paper_id, title, ..., direct_pdf_links, mdpi_links
// ---------------------------------------------------------------------
function parseCsvToJobs(text) {
  const lines = text.replace(/\r\n/g, "\n").split("\n").filter(l => l.trim() !== "");
  if (lines.length < 2) return [];

  const header = splitCsvLine(lines[0]).map(h => h.trim());
  const idx = name => header.findIndex(h => h.toLowerCase() === name.toLowerCase());

  const idxPaperId = idx("paper_id");
  const idxTitle = idx("title");
  const idxUrl = idx("url");
  const idxDirect = idx("direct_pdf_links");
  const idxMdpi = idx("mdpi_links");

  const out = [];
  for (let i = 1; i < lines.length; i++) {
    const cols = splitCsvLine(lines[i]);
    const paperId = (cols[idxPaperId] || "").trim();
    const title = (cols[idxTitle] || "").trim();

    if (idxUrl !== -1) {
      const u = (cols[idxUrl] || "").trim();
      if (u.startsWith("http")) out.push({ paper_id: paperId, title, url: u });
    } else {
      [idxDirect, idxMdpi].forEach(colIdx => {
        if (colIdx === -1) return;
        const val = (cols[colIdx] || "").trim();
        if (!val) return;
        val.split(";").forEach(u => {
          u = u.trim();
          if (u.startsWith("http")) out.push({ paper_id: paperId, title, url: u });
        });
      });
    }
  }
  return out;
}

function splitCsvLine(line) {
  const result = [];
  let cur = "", inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      inQuotes = !inQuotes;
    } else if (ch === "," && !inQuotes) {
      result.push(cur);
      cur = "";
    } else {
      cur += ch;
    }
  }
  result.push(cur);
  return result;
}

function csvEscape(s) {
  s = s || "";
  if (s.includes(",") || s.includes('"')) return '"' + s.replace(/"/g, '""') + '"';
  return s;
}

// ---------------------------------------------------------------------
// URL resolution
// ---------------------------------------------------------------------
async function resolvePdfUrl(url) {
  try {
    const u = new URL(url);

    // --- IEEE: the iel8/iel7/.../<arnumber>.pdf URLs in the CSV are
    // internal viewer-iframe URLs and don't work visited standalone, even
    // with valid IP-based institutional access. IEEE's real public
    // download endpoint is stamp.jsp?arnumber=<id>, confirmed working.
    if (u.hostname.includes("ieeexplore.ieee.org")) {
      const match = u.pathname.match(/\/(\d+)\.pdf$/i);
      if (match) {
        const arnumber = match[1];
        return `${u.protocol}//${u.hostname}/stamp/stamp.jsp?arnumber=${arnumber}`;
      }
      return url;
    }

    // --- MDPI: landing-page URL -> real /pdf?version=... link, read off
    // the actual article page (the plain version number is unique per
    // article and only appears there).
    if (u.hostname.includes("mdpi.com")) {
      if (u.searchParams.has("version")) return url; // already a real versioned link

      let path = u.pathname.replace(/\/pdf\/?$/i, "").replace(/\/$/, "");
      const landingUrl = `${u.protocol}//${u.hostname}${path}`;

      const resp = await fetch(landingUrl, { credentials: "include" });
      if (resp.ok) {
        const html = await resp.text();
        let match = html.match(/class="[^"]*UD_ArticlePDF[^"]*"[^>]*href="([^"]+)"/i);
        if (!match) match = html.match(/href="([^"]+)"[^>]*class="[^"]*UD_ArticlePDF[^"]*"/i);
        if (!match) match = html.match(/href="([^"]*\/pdf\?version=[^"]+)"/i);
        if (match) return new URL(match[1], landingUrl).href;
      }
      return landingUrl + "/pdf";
    }
  } catch (e) {
    return url;
  }
  return url;
}

function sanitizeFilename(s) {
  return (s || "").replace(/[\\/:*?"<>|]/g, "_").slice(0, 120);
}

// ---------------------------------------------------------------------
// Row UI helpers
// ---------------------------------------------------------------------
function addRow(i, job) {
  const tr = document.createElement("tr");
  tr.id = `row-${i}`;
  tr.innerHTML = `
    <td>${i + 1}</td>
    <td>${escapeHtml(job.paper_id)}</td>
    <td>${escapeHtml(job.title).slice(0, 60)}</td>
    <td style="max-width:280px; overflow-wrap:break-word;">${escapeHtml(job.url)}</td>
    <td class="status-pending">pending</td>
  `;
  resultsBody.appendChild(tr);
}

function setRowStatus(i, statusText, cls) {
  const tr = document.getElementById(`row-${i}`);
  if (!tr) return;
  const cell = tr.children[4];
  cell.textContent = statusText;
  cell.className = cls;
}

function escapeHtml(s) {
  return (s || "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

// Rename downloads triggered by the IEEE click-list to paper_id-based names.
//
// IMPORTANT: stamp.jsp internally redirects to a different, unpredictable
// IEEE URL that actually streams the PDF - so matching on the stamp.jsp
// URL itself (what you clicked) fails once Chrome asks us to name the
// FINAL download, which has a different URL. The IEEE article ID does
// survive into that final URL/filename though (that's why Chrome's
// default name was e.g. "10606464.pdf"), so we match on that instead.
const pendingRenames = new Map(); // key: article id (or any unique substring) -> filename

chrome.downloads.onDeterminingFilename.addListener((downloadItem, suggest) => {
  const haystack = `${downloadItem.url} ${downloadItem.finalUrl || ""} ${downloadItem.filename || ""}`;
  for (const [idKey, filename] of pendingRenames.entries()) {
    if (haystack.includes(idKey)) {
      suggest({ filename, conflictAction: "overwrite" });
      pendingRenames.delete(idKey);
      return;
    }
  }
  suggest();
});

function addIeeeLink(job, url, filename) {
  ieeeBox.style.display = "block";

  // Extract the IEEE article id from the stamp.jsp URL (?arnumber=NNNNN)
  // to use as the robust match key, instead of the full URL.
  let idKey;
  try {
    idKey = new URL(url).searchParams.get("arnumber") || url;
  } catch {
    idKey = url;
  }
  pendingRenames.set(idKey, filename);

  const displayName = filename.split("/").pop();
  const a = document.createElement("a");
  a.href = url;
  a.target = "_blank";
  a.rel = "noopener";
  a.innerHTML = `${escapeHtml(job.paper_id)} - ${escapeHtml(job.title).slice(0, 90)}<small>${escapeHtml(url)} &rarr; will save as ${escapeHtml(displayName)}</small>`;
  a.addEventListener("click", () => a.classList.add("clicked"), { once: true });
  ieeeLinksDiv.appendChild(a);
}

// ---------------------------------------------------------------------
// Main download loop
// ---------------------------------------------------------------------
async function runAll() {
  let success = 0;

  for (let i = 0; i < jobs.length; i++) {
    if (isStopped) {
      summary.textContent = `Stopped by user at item ${i}/${jobs.length}. ${success} succeeded so far.`;
      break;
    }
    while (isPaused && !isStopped) {
      await sleep(300);
    }
    if (isStopped) {
      summary.textContent = `Stopped by user at item ${i}/${jobs.length}. ${success} succeeded so far.`;
      break;
    }

    const job = jobs[i];
    setRowStatus(i, "resolving link...", "status-pending");
    const resolvedUrl = await resolvePdfUrl(job.url);

    const isIeee = (() => { try { return new URL(resolvedUrl).hostname.includes("ieeexplore.ieee.org"); } catch { return false; } })();
    const filename = `downloaded_pdfs/${sanitizeFilename(job.paper_id + "_" + job.title)}.pdf`;

    if (isIeee) {
      // Confirmed by testing: IEEE requires a genuine human click - no
      // programmatic download call or automated navigation can produce
      // that signal. Hand it to the click-through list instead of
      // pretending automation can do it.
      addIeeeLink(job, resolvedUrl, filename);
      setRowStatus(i, "needs manual click (see IEEE list above)", "status-manual");
      results.push({
        paper_id: job.paper_id, title: job.title, url: job.url,
        resolved_url: resolvedUrl, status: "manual_click_required (IEEE)", file: filename,
      });
      await sleep(50);
      continue;
    }

    setRowStatus(i, "downloading...", "status-pending");
    summary.textContent = `Processing ${i + 1}/${jobs.length}...`;

    const outcome = await downloadOne(resolvedUrl, filename);

    if (outcome.ok) {
      success++;
      setRowStatus(i, "success", "status-success");
    } else {
      setRowStatus(i, "failed: " + outcome.msg, "status-failed");
    }

    results.push({
      paper_id: job.paper_id, title: job.title, url: job.url,
      resolved_url: resolvedUrl, status: outcome.ok ? "success" : "failed: " + outcome.msg,
      file: outcome.ok ? filename : "",
    });

    await sleep(700); // be polite between requests
  }

  if (!isStopped) {
    summary.textContent = `Done. ${success}/${jobs.length} downloaded successfully.`;
  }
}

// ---------------------------------------------------------------------
// Download engine: chrome.downloads.download() directly.
//
// This is a privileged extension API - it's specifically designed to let
// extensions trigger downloads WITHOUT needing a user click, so it
// doesn't hit Chrome's "automated download" blocking the way a
// script-driven tab navigation does. Earlier failures on IEEE weren't
// actually about navigation-vs-fetch - they were because the CSV's raw
// iel8/.../<id>.pdf URLs are internal viewer-iframe links that don't
// resolve to anything outside that embedded context. Now that
// resolvePdfUrl() rewrites those to the real stamp.jsp endpoint, a plain
// download call should carry your session cookies correctly.
// ---------------------------------------------------------------------
async function downloadOne(url, filename) {
  const downloadId = await new Promise((resolve) => {
    chrome.downloads.download({ url, filename, saveAs: false, conflictAction: "overwrite" }, (id) => {
      resolve(chrome.runtime.lastError || !id ? null : id);
    });
  });

  if (downloadId === null) {
    return { ok: false, msg: chrome.runtime.lastError?.message || "failed to start download" };
  }

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
    return { ok: false, msg: `download interrupted: ${item.error || "unknown error"}` };
  }

  if (item.state === "complete") {
    const mime = (item.mime || "").toLowerCase();
    if (mime.includes("html")) {
      await new Promise((res) => chrome.downloads.removeFile(downloadId, () => res()));
      await new Promise((res) => chrome.downloads.erase({ id: downloadId }, res));
      return { ok: false, msg: "downloaded a webpage instead of a PDF (login/paywall/access page)" };
    }
    if (mime && !mime.includes("pdf")) {
      return { ok: false, msg: `unexpected file type (mime=${mime})` };
    }
    return { ok: true, msg: "ok" };
  }

  return { ok: false, msg: `unexpected download state: ${item.state}` };
}

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }
