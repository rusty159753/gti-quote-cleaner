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
      const py = await fetch("cleaner.py");
      if (!py.ok) throw new Error("Could not load cleaner.py (" + py.status + ").");
      pyodide.FS.writeFile("cleaner.py", await py.text());
      pyodide.runPython("import cleaner");

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

      const resultProxy = await pyodide.runPythonAsync(
        "cleaner.clean_workbook(bytes(_raw_bytes.to_py()))"
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

  init();
})();
