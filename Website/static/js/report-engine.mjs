import { executeWorkerExport } from "/assets/officeimo/officeimo-execution.mjs";
import { operationalRows, operationalRowsBySite } from "/js/report-engine-job.mjs";
import { analyticalRows } from "/js/analytical-report-job.mjs";
import { wideAnalyticalRows } from "/js/wide-analytical-report-job.mjs";

const byId = id => document.getElementById(id);
const form = byId("report-form"), rowSelect = byId("report-rows"), formatSelect = byId("report-format");
const layoutSelect = byId("report-layout"), languageSelect = byId("report-language");
const start = byId("report-generate"), cancel = byId("report-cancel"), status = byId("report-status");
const progress = byId("report-progress"), download = byId("report-download"), measurement = byId("report-measurement");
let active, downloadUrl;
const latencyFormat = new Intl.NumberFormat("en", { minimumFractionDigits: 1, maximumFractionDigits: 1 });
function sourceRows(count, layout = layoutSelect.value) {
  return layout === "wide" ? wideAnalyticalRows(count) : layout === "hierarchical" ? analyticalRows(count) : layout === "grouped" ? operationalRowsBySite(count) : operationalRows(count);
}
function translatedLayout() { return layoutSelect.value !== "analysis"; }
function preview() {
  const count = Number(rowSelect.value), body = byId("report-preview");
  body.replaceChildren();
  let shown = 0;
  for (const row of sourceRows(count)) {
    if (shown === 8) break;
    const tr = document.createElement("tr");
    for (const [index, value] of row.slice(0, 4).entries()) {
      const td = document.createElement("td"); td.textContent = value == null ? "—" : index === 3 ? latencyFormat.format(value) + " ms" : value;
      if (index === 2) { td.className = "report-health"; td.dataset.state = value; }
      tr.append(td);
    }
    body.append(tr); shown++;
  }
  if (!shown) {
    const tr = document.createElement("tr"), td = document.createElement("td");
    td.colSpan = 4; td.textContent = "Empty view — the report keeps its headings."; tr.append(td); body.append(tr);
  }
  byId("report-preview-count").textContent = "First " + shown + " of " + count.toLocaleString("en") + " services";
  const scope = byId("report-preview-scope");
  scope.hidden = layoutSelect.value !== "wide";
  scope.textContent = scope.hidden ? "" : "Previewing 4 of 16 report fields. PDF repeats service IDs across readable column bands; Excel keeps every field together.";
}
function busy(value) {
  start.disabled = rowSelect.disabled = formatSelect.disabled = value;
  layoutSelect.disabled = value || formatSelect.value === "csv";
  languageSelect.disabled = value || !translatedLayout() || formatSelect.value === "csv";
  cancel.disabled = !value;
  cancel.hidden = !value;
  form.setAttribute("aria-busy", String(value));
}
function clearResult() {
  if (downloadUrl) { URL.revokeObjectURL(downloadUrl); downloadUrl = undefined; }
  download.hidden = true; download.removeAttribute("href");
  start.textContent = "Generate report"; start.classList.replace("imo-btn-ghost", "imo-btn-primary");
  progress.value = 0; measurement.textContent = ""; measurement.hidden = true;
}
function settingsChanged() {
  clearResult(); preview(); busy(false);
  status.textContent = "Ready. Choose a workload and generate a report.";
}
async function font(name, signal) {
  const response = await fetch("/assets/report-fonts/Carlito-" + name + ".ttf", { signal });
  if (!response.ok) throw new Error("The report font could not be loaded. Please try again.");
  return new Uint8Array(await response.arrayBuffer());
}
form.addEventListener("submit", async event => {
  event.preventDefault(); if (active) return;
  clearResult();
  const controller = new AbortController(), format = formatSelect.value, count = Number(rowSelect.value);
  active = { controller }; busy(true); progress.max = Math.max(1, count); progress.value = 0;
  status.textContent = "Preparing " + format.toUpperCase() + " report…"; measurement.textContent = "";
  const started = performance.now(), parts = [];
  let worker, accepted = 0;
  try {
    const layout = format === "csv" ? "analysis" : layoutSelect.value;
    const language = languageSelect.value;
    const input = { format, layout, ...(layout !== "analysis" ? { language } : {}), rowCount: count,
      ...(layout === "hierarchical" || layout === "wide" ? { scope: language === "pl" ? `Wybrany widok usług · ${count.toLocaleString("pl")} rekordów.` :
        `Selected service view · ${count.toLocaleString("en")} services.` } : {}),
      ...(format === "pdf" ? { fonts: { regular: await font("Regular", controller.signal), bold: await font("Bold", controller.signal) } } : {}) };
    controller.signal.throwIfAborted();
    worker = new Worker("/js/report-engine-worker.mjs", { type: "module", name: "OfficeIMO report" });
    worker.addEventListener("error", event => { event.preventDefault(); controller.abort(new Error("The report worker could not run. Refresh the page and try again.")); });
    worker.addEventListener("messageerror", () => controller.abort(new Error("The report worker could not read a message.")));
    const result = await executeWorkerExport(worker, sourceRows(count, layout), { write(bytes) {
      accepted += bytes.length;
      if (accepted > 96 * 1024 * 1024) throw new Error("This download exceeds the workbench's 96 MiB limit. Choose fewer services.");
      parts.push(bytes.slice());
    } }, { input, signal: controller.signal, onProgress(update) {
      progress.value = Math.min(count, update.rows);
      status.textContent = "Writing " + update.rows.toLocaleString("en") + " rows…";
    } });
    if (result.rows !== count) throw new Error("The report row count differed from the selected view.");
    const blob = new Blob(parts, { type: format === "xlsx" ? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" : format === "pdf" ? "application/pdf" : "text/csv;charset=utf-8" });
    if (downloadUrl) URL.revokeObjectURL(downloadUrl);
    downloadUrl = URL.createObjectURL(blob); download.href = downloadUrl; download.download = "service-health." + format;
    download.textContent = "Download " + format.toUpperCase() + " report"; download.hidden = false;
    start.textContent = "Generate again"; start.classList.replace("imo-btn-primary", "imo-btn-ghost");
    progress.value = Math.max(1, count);
    status.textContent = "Ready. " + count.toLocaleString("en") + " services exported.";
    measurement.textContent = (performance.now() - started).toLocaleString("en", { maximumFractionDigits: 0 }) + " ms elapsed · " +
      (blob.size / 1024).toLocaleString("en", { maximumFractionDigits: 1 }) + " KiB. This device and this run; includes preparation and transfer.";
    measurement.hidden = false;
  } catch (error) {
    progress.value = 0;
    status.textContent = controller.signal.aborted && controller.signal.reason?.name === "AbortError"
      ? "Cancelled. You can change the settings and try again." : (error instanceof Error ? error.message : "The report could not be generated.");
  } finally {
    // Hiding a focused Cancel must not leave keyboard users on the document body.
    // Preserve an intentional move to another control while cancellation settles.
    const restoreFocus = document.activeElement === cancel || active?.cancelOwnedFocus && document.activeElement === document.body;
    parts.length = 0; worker?.terminate(); active = undefined; busy(false);
    if (restoreFocus && document.hasFocus()) start.focus();
  }
});
cancel.addEventListener("click", () => {
  if (active) { active.cancelOwnedFocus = document.activeElement === cancel; active.controller.abort(new DOMException("Report cancelled", "AbortError")); }
  cancel.disabled = true; status.textContent = "Cancelling…";
});
rowSelect.addEventListener("change", settingsChanged);
layoutSelect.addEventListener("change", settingsChanged);
languageSelect.addEventListener("change", settingsChanged);
formatSelect.addEventListener("change", () => { if (formatSelect.value === "csv") layoutSelect.value = "analysis"; settingsChanged(); });
addEventListener("pagehide", event => {
  active?.controller.abort(new DOMException("Page closed", "AbortError"));
  // A back/forward-cache entry keeps this document and its visible download.
  // Release completed bytes only when the document is actually being left.
  if (!event.persisted && downloadUrl) { URL.revokeObjectURL(downloadUrl); downloadUrl = undefined; }
});
preview(); busy(false); status.textContent = "Ready. Choose a workload and generate a report.";
