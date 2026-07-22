/*
 * GTI Quote Cleaner — browser glue.
 *
 * Loads Pyodide, installs the pure-Python packages the pipeline needs
 * (xlrd to read legacy .xls, openpyxl to write .xlsx), loads cleaner.py,
 * and wires the drop zone / Clean button / download link / run summary.
 *
 * The user's file is read into memory and handed to Python in-tab. It is
 * never uploaded or transmitted anywhere.
 */

(function () {
  "use strict";

  const drop = document.getElementById("drop");
  const fileInput = document.getElementById("fileInput");
  const fileNameEl = document.getElementById("fileName");
  const cleanBtn = document.getElementById("cleanBtn");
  const downloadLink = document.getElementById("downloadLink");
  const statusEl = document.getElementById("status");
  const errorEl = document.getElementById("error");
  const summaryEl = document.getElementById("summary");
  const flagsEl = document.getElementById("flags");

  const namesText = document.getElementById("namesText");
  const namesSave = document.getElementById("namesSave");
  const namesExport = document.getElementById("namesExport");
  const namesImport = document.getElementById("namesImport");
  const namesFile = document.getElementById("namesFile");
  const namesStatus = document.getElementById("namesStatus");

  // The "Created By" mapping lives ONLY in this browser (localStorage). It is
  // never uploaded and never stored in the code.
  const NAMES_KEY = "gti_created_by_map_v1";

  let pyodide = null;
  let ready = false;
  let selectedFile = null;
  let lastObjectUrl = null;

  function setStatus(text, busy) {
    statusEl.textContent = text || "";
    statusEl.classList.toggle("busy", !!busy);
  }
  function showError(text) {
    errorEl.textContent = text;
    errorEl.classList.add("show");
  }
  function clearError() {
    errorEl.textContent = "";
    errorEl.classList.remove("show");
  }

  // ---- "Created By" name mapping (browser-local) ------------------------

  function loadNamesMap() {
    try {
      const raw = localStorage.getItem(NAMES_KEY);
      const map = raw ? JSON.parse(raw) : {};
      return map && typeof map === "object" ? map : {};
    } catch (e) {
      return {};
    }
  }

  function mapToText(map) {
    return Object.keys(map).map((k) => k + " = " + map[k]).join("\n");
  }

  // Parse lines "RAW = Clean" (also accepts =>, ->, :). '#' starts a comment.
  function textToMap(text) {
    const map = {};
    let count = 0;
    const lines = text.split(/\r?\n/);
    for (let raw of lines) {
      const line = raw.replace(/#.*$/, "").trim();
      if (!line) continue;
      const m = line.match(/^(.*?)\s*(?:=>|->|[:=])\s*(.*)$/);
      if (!m) throw new Error('Could not read this line: "' + raw.trim() + '". Use RAW = Clean Name.');
      const key = m[1].trim();
      const val = m[2].trim();
      if (!key || !val) throw new Error('Both sides are required on: "' + raw.trim() + '".');
      map[key] = val;
      count++;
    }
    return { map: map, count: count };
  }

  function setNamesStatus(text, ok) {
    namesStatus.textContent = text || "";
    namesStatus.classList.toggle("ok", !!ok);
  }

  namesSave.addEventListener("click", () => {
    try {
      const { map, count } = textToMap(namesText.value);
      localStorage.setItem(NAMES_KEY, JSON.stringify(map));
      namesText.value = mapToText(map);
      setNamesStatus("Saved " + count + " name" + (count === 1 ? "" : "s") + " to this browser.", true);
    } catch (err) {
      setNamesStatus(err.message, false);
    }
  });

  namesExport.addEventListener("click", () => {
    let map;
    try {
      map = textToMap(namesText.value).map;
    } catch (err) {
      setNamesStatus(err.message, false);
      return;
    }
    const blob = new Blob([JSON.stringify(map, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "gti-created-by-names.json";
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    setNamesStatus("Exported to gti-created-by-names.json — keep it somewhere private.", true);
  });

  namesImport.addEventListener("click", () => namesFile.click());
  namesFile.addEventListener("change", async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    try {
      const map = JSON.parse(await file.text());
      if (!map || typeof map !== "object" || Array.isArray(map)) {
        throw new Error("That file is not a names mapping.");
      }
      namesText.value = mapToText(map);
      localStorage.setItem(NAMES_KEY, JSON.stringify(map));
      setNamesStatus("Imported " + Object.keys(map).length + " names and saved to this browser.", true);
    } catch (err) {
      setNamesStatus("Could not import: " + err.message, false);
    } finally {
      namesFile.value = "";
    }
  });

  // ---- Pyodide initialization ------------------------------------------

  async function init() {
    try {
      setStatus("Getting the cleaner ready… (first visit downloads a one-time engine)", true);
      pyodide = await loadPyodide({
        indexURL: "https://cdn.jsdelivr.net/pyodide/v0.26.4/full/",
      });

      setStatus("Loading spreadsheet tools…", true);
      await pyodide.loadPackage("micropip");
      const micropip = pyodide.pyimport("micropip");
      // Both are pure-Python; installed from PyPI at load time (code only,
      // never any user data).
      await micropip.install(["xlrd", "openpyxl"]);

      setStatus("Loading the cleaning rules…", true);
      // no-store so an updated cleaner.py always takes effect (it is small and
      // is the file most likely to be edited).
      const py = await fetch("cleaner.py", { cache: "no-store" });
      if (!py.ok) throw new Error("Could not load cleaner.py (" + py.status + ").");
      pyodide.FS.writeFile("cleaner.py", await py.text());
      pyodide.runPython("import cleaner\nimport json as _json");

      ready = true;
      drop.classList.remove("disabled");
      setStatus("Ready. Drop your .xls export to begin.", false);
      refreshButton();
    } catch (err) {
      setStatus("", false);
      showError(
        "The cleaner could not start: " + (err && err.message ? err.message : err) +
        " Please refresh the page and try again."
      );
    }
  }

  // ---- File selection ---------------------------------------------------

  function selectFile(file) {
    clearError();
    resetOutput();
    if (!file) return;
    if (!/\.xls$/i.test(file.name)) {
      showError("Please choose the raw .xls export (not .xlsx or another format).");
      selectedFile = null;
      fileNameEl.textContent = "";
      refreshButton();
      return;
    }
    selectedFile = file;
    fileNameEl.textContent = file.name;
    refreshButton();
  }

  function refreshButton() {
    cleanBtn.disabled = !(ready && selectedFile);
  }

  function resetOutput() {
    summaryEl.classList.remove("show");
    downloadLink.style.display = "none";
    flagsEl.innerHTML = "";
    if (lastObjectUrl) {
      URL.revokeObjectURL(lastObjectUrl);
      lastObjectUrl = null;
    }
  }

  drop.addEventListener("click", () => { if (ready) fileInput.click(); });
  fileInput.addEventListener("change", (e) => selectFile(e.target.files[0]));

  ["dragenter", "dragover"].forEach((ev) =>
    drop.addEventListener(ev, (e) => {
      e.preventDefault();
      if (ready) drop.classList.add("hot");
    })
  );
  ["dragleave", "drop"].forEach((ev) =>
    drop.addEventListener(ev, (e) => {
      e.preventDefault();
      drop.classList.remove("hot");
    })
  );
  drop.addEventListener("drop", (e) => {
    if (!ready) return;
    const file = e.dataTransfer.files && e.dataTransfer.files[0];
    selectFile(file);
  });

  // ---- Cleaning ---------------------------------------------------------

  cleanBtn.addEventListener("click", runClean);

  async function runClean() {
    if (!ready || !selectedFile) return;
    clearError();
    resetOutput();
    cleanBtn.disabled = true;
    setStatus("Cleaning “" + selectedFile.name + "”…", true);

    try {
      const buf = await selectedFile.arrayBuffer();
      const bytes = new Uint8Array(buf);
      pyodide.globals.set("_raw_bytes", bytes);
      pyodide.globals.set("_cb_json", JSON.stringify(loadNamesMap()));

      const resultProxy = await pyodide.runPythonAsync(
        "cleaner.clean_workbook(bytes(_raw_bytes.to_py()), _json.loads(_cb_json))"
      );

      const ok = resultProxy.get("ok");
      if (!ok) {
        const message = resultProxy.get("error");
        showError(message);
        setStatus("No file was produced.", false);
        resultProxy.destroy();
        cleanBtn.disabled = false;
        return;
      }

      const xlsxProxy = resultProxy.get("xlsx");
      const xlsxBytes = xlsxProxy.toJs();
      xlsxProxy.destroy();

      const summaryProxy = resultProxy.get("summary");
      const summary = summaryProxy.toJs({ dict_converter: Object.fromEntries });
      summaryProxy.destroy();
      resultProxy.destroy();

      offerDownload(xlsxBytes, selectedFile.name);
      renderSummary(summary);
      setStatus("Done. Your cleaned file is ready to download.", false);
    } catch (err) {
      showError("Something went wrong while cleaning: " + (err && err.message ? err.message : err));
      setStatus("", false);
    } finally {
      pyodide.globals.set("_raw_bytes", undefined);
      cleanBtn.disabled = false;
    }
  }

  function offerDownload(xlsxBytes, sourceName) {
    const blob = new Blob([xlsxBytes], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    lastObjectUrl = URL.createObjectURL(blob);
    const base = sourceName.replace(/\.xls$/i, "");
    downloadLink.href = lastObjectUrl;
    downloadLink.download = base + "-cleaned.xlsx";
    downloadLink.style.display = "inline-block";
  }

  // ---- Summary rendering ------------------------------------------------

  function renderSummary(s) {
    document.getElementById("stRowsIn").textContent = fmt(s.rows_in);
    document.getElementById("stQuotes").textContent = fmt(s.quotes_out);
    document.getElementById("stOrdered").textContent = fmt(s.quotes_ordered);
    document.getElementById("stAmount").textContent =
      "$" + Number(s.total_amount).toLocaleString(undefined, {
        minimumFractionDigits: 2, maximumFractionDigits: 2,
      });

    const flagBlocks = [];
    if (nonEmpty(s.flag_unparseable_dates)) {
      flagBlocks.push(flagBlock(
        "Dates that could not be read (left as text for you to check):",
        s.flag_unparseable_dates
      ));
    }
    if (nonEmpty(s.flag_unknown_created_by)) {
      flagBlocks.push(flagBlock(
        "Unrecognized “Created By” names (left exactly as written — add them to the mapping if needed):",
        s.flag_unknown_created_by
      ));
    }
    if (nonEmpty(s.flag_nonadjacent_duplicate_quotes)) {
      flagBlocks.push(flagBlock(
        "Quote numbers that appear more than once, separated by other quotes (not merged — please review):",
        s.flag_nonadjacent_duplicate_quotes
      ));
    }

    if (flagBlocks.length === 0) {
      flagsEl.innerHTML = '<div class="allclear">✓ No anomalies flagged. Dates, names, and quote numbers all checked out.</div>';
    } else {
      flagsEl.innerHTML = flagBlocks.join("");
    }
    summaryEl.classList.add("show");
  }

  function flagBlock(title, items) {
    const list = items.map((x) => "<li>" + escapeHtml(String(x)) + "</li>").join("");
    return '<div class="flag">' + escapeHtml(title) + "<ul>" + list + "</ul></div>";
  }

  function nonEmpty(arr) { return Array.isArray(arr) && arr.length > 0; }
  function fmt(n) { return Number(n).toLocaleString(); }
  function escapeHtml(s) {
    return s.replace(/[&<>"']/g, (c) => (
      { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]
    ));
  }

  namesText.value = mapToText(loadNamesMap());
  init();
})();
