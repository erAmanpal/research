# Scopus Research PDF Downloader — Chrome Extension

Everything is now embedded in the extension itself — no Python, no
separate scripts, nothing to run outside Chrome.

## What it does

1. You upload your Scopus export (`.xlsx`, `.xls`, or `.csv`) with a
   serial-number column (e.g. `S.No`) and a `DOI` column.
2. For every row, it checks — automatically, using only free,
   key-free, legal APIs:
   - **Unpaywall** (open-access index)
   - **OpenAlex** (open-access index)
   - a direct, verified check of the DOI link itself
3. If a legally open-access PDF is found, it downloads **automatically**
   into a folder named after the source that found it:
   ```
   Downloads/ScopusPDFs/Unpaywall/<S.No>_<Title>.pdf
   Downloads/ScopusPDFs/OpenAlex/<S.No>_<Title>.pdf
   Downloads/ScopusPDFs/DirectOpenAccess/<S.No>_<Title>.pdf
   ```
4. If nothing open access is found (i.e. the article is restricted /
   paywalled), the row is added to the **Manual Download List** — a
   click-through table, exactly like the IEEE list in the original
   sample extension — with a link to the article and the exact
   filename you should save it as.

## What's intentionally NOT in here (and why)

- **Consensus** — needs a paid API key. Left out per your instruction.
- **Sci-Hub** — distributes copyrighted, paywalled articles without
  publisher authorization. Including it would mean this extension
  helps commit copyright infringement, so it's not built in, in any
  form.
- **Google Scholar automation** — scraping Scholar's result pages is
  against Google's Terms of Service. Left out entirely, unlike the
  earlier Python version which offered it as an opt-in stage.
- **Bypassing any login/paywall page** — every download is verified
  by its actual file type (MIME). If a "download" turns out to be an
  HTML login or paywall page instead of a real PDF, it's deleted
  immediately and the row goes to the Manual Download list instead of
  silently saving a broken/wrong file.

## Install (unpacked, for development/personal use)

1. Open `chrome://extensions`.
2. Turn on **Developer mode** (top-right).
3. Click **Load unpacked** and select this `chrome_extension` folder.
4. Click the extension icon → **Open Downloader**.

## Using it

1. Enter your email (Unpaywall's API terms require one — it's only
   sent to Unpaywall's API, nowhere else).
2. Upload your Scopus file. The extension tries to auto-detect the
   S.No / DOI / Title columns; correct them from the dropdowns if it
   guesses wrong.
3. Click **Start Downloading**. Watch the table update per row.
4. If you close the tab partway through, reopening the downloader page
   offers **Resume** — it picks up where it left off and skips rows
   already downloaded (checked against Chrome's own download history).
5. When done, use **Export Full Report CSV** for a complete audit
   trail (status/source/file per row), and **Export Manual List CSV**
   for the restricted items.
6. Work through the Manual Download list: click "Open article" for
   each row, download the PDF through your own institutional login,
   and save it as the suggested filename shown (copy button provided)
   into `Downloads/ScopusPDFs/Manual/`.

## Notes / limitations

- All paths are relative to your browser's **default Downloads
  folder** — a Chrome extension cannot write to an arbitrary absolute
  path on disk. If you want everything under one project folder, set
  that as your default Downloads location in Chrome settings first.
- "Resume" and duplicate-skipping rely on Chrome's own download
  history (`chrome.downloads.search`), not on reading the filesystem
  directly — extensions aren't allowed arbitrary disk read access.
  If you manually delete a downloaded file from disk but not from
  Chrome's download history, it will still be treated as "already
  downloaded" and skipped.
- Unpaywall/OpenAlex coverage is good but not universal. A "Restricted"
  result means *these two indexes don't know of a legal open copy*,
  not necessarily that no legal open copy exists anywhere — the Manual
  list exists for exactly that gap.
- Large datasets (hundreds+ rows) will take a while — there's a
  built-in ~0.5s delay between rows to stay polite to the free APIs.
  Progress is saved after every row, so it's safe to pause/close and
  resume later.

## Files

```
manifest.json      - MV3 manifest (downloads, storage, unlimitedStorage permissions)
popup.html/js       - small launcher popup
downloader.html      - main UI (upload, column mapping, results table, manual list)
downloader.js        - all logic: parsing, Unpaywall/OpenAlex lookups, download +
                        verification, manual list, resume/persistence
libs/xlsx.full.min.js - SheetJS, bundled locally (Manifest V3 disallows loading
                         scripts from a CDN at runtime)
```
