// Capture an actual HtmlForgeX-produced DOM; the host owns selection, order and visibility.
document.addEventListener("DOMContentLoaded", () => {
  const table = document.getElementById("operational-inventory"), status = document.getElementById("inventory-status");
  const header = table.tHead.rows[0], body = table.tBodies[0], schema = JSON.parse(table.dataset.columnSchema);
  const fields = table.dataset.columnFields.split(" "), controls = [...document.querySelectorAll("button,input")];
  [...header.cells].slice(1).forEach((cell, index) => { cell.dataset.field = fields[index]; });
  const dataRows = () => [...body.rows];
  for (const row of dataRows()) row.dataset.rowId = row.querySelector("[data-row-id]").dataset.rowId;
  document.getElementById("inventory-filter").addEventListener("input", event => {
    const query = event.target.value.toLocaleLowerCase();
    for (const row of dataRows()) row.hidden = !row.textContent.toLocaleLowerCase().includes(query);
  });
  let descending = false;
  document.getElementById("inventory-sort").addEventListener("click", () => {
    descending = !descending;
    dataRows().sort((left, right) => left.dataset.rowId.localeCompare(right.dataset.rowId) * (descending ? -1 : 1))
      .forEach(row => body.append(row));
  });
  document.getElementById("inventory-hide-site").addEventListener("click", () => {
    const index = [...header.cells].findIndex(cell => cell.dataset.field === "site"), hide = !header.cells[index].hidden;
    for (const row of [header, ...dataRows()]) row.cells[index].hidden = hide;
  });
  document.getElementById("inventory-move-latency").addEventListener("click", () => {
    const index = [...header.cells].findIndex(cell => cell.dataset.field === "latency");
    for (const row of [header, ...dataRows()]) {
      const cell = row.cells[index];
      if (index === 1) row.append(cell); else row.insertBefore(cell, row.cells[1]);
    }
  });
  document.getElementById("inventory-select").addEventListener("click", () => {
    for (const row of dataRows()) if (!row.hidden) row.querySelector("input[type=checkbox]").checked = true;
  });
  function scalar(span) {
    const raw = span.dataset.rawValue;
    if (span.dataset.valueType === "null") return null;
    if (span.dataset.valueType === "number") return Number(raw);
    if (span.dataset.valueType === "date") return new Date(raw);
    return raw;
  }
  function capture() {
    const columns = [...header.cells].filter(cell => cell.dataset.field && !cell.hidden).map(cell => {
      const source = schema.find(column => column.key === cell.dataset.field);
      const { format, ...plain } = source;
      return { ...plain, ...(format === null ? {} : { format }), header: cell.textContent.trim() };
    });
    const inventory = dataRows(), visible = inventory.filter(row => !row.hidden);
    const selected = visible.filter(row => row.querySelector("input[type=checkbox]").checked);
    const filter = document.getElementById("inventory-filter").value;
    const scope = "Checked services in the visible view. Selection: " + selected.length + " of " + visible.length + " visible services (" + inventory.length + " in the inventory). "
      + "Row text filter: " + (filter ? JSON.stringify(filter) : "none") + ". Order: service " + (descending ? "descending" : "ascending")
      + ". Fields: visible columns in their current order.";
    function* rows() {
      for (const row of selected) {
        const identity = row.querySelector("[data-row-id]");
        yield [...columns.map(column => scalar(row.querySelector('[data-field="' + column.key + '"]'))),
          identity.dataset.portfolio, identity.dataset.environment, identity.dataset.owner];
      }
    }
    return { columns, rows, ids: selected.map(row => row.dataset.rowId), scope };
  }
  let lastExport;
  async function exportView(format) {
    controls.forEach(control => { control.disabled = true; });
    status.textContent = "Exporting…";
    let owned, url;
    try {
      const selected = capture(), assets = globalThis.operationalInventoryAssets, chunks = [];
      const bootstrap = assets.classicScript + "\n" + assets.executionScript + "\n" + assets.workerScript;
      url = URL.createObjectURL(new Blob([bootstrap], { type: "text/javascript" }));
      owned = new Worker(url); // Ordinary classic Blob worker, including file:// hosts.
      const observed = { maxRows: 0, maxChunkBytes: 0 }, listeners = new Map();
      const endpoint = {
        postMessage(message) { if (message.kind === "rows") observed.maxRows = Math.max(observed.maxRows, message.rows.length); owned.postMessage(message); },
        addEventListener(type, listener) {
          const wrapped = event => {
            if (type === "message" && event.data.kind === "output") observed.maxChunkBytes = Math.max(observed.maxChunkBytes, event.data.bytes.length);
            listener(event);
          };
          listeners.set(listener, { type, wrapped }); owned.addEventListener(type, wrapped);
        },
        removeEventListener(type, listener) {
          const entry = listeners.get(listener); if (entry) { owned.removeEventListener(type, entry.wrapped); listeners.delete(listener); }
        }
      };
      const completed = await OfficeIMO.executeWorkerExport(endpoint, selected.rows(), { write(bytes) { chunks.push(bytes.slice()); } }, {
        input: { format, columns: selected.columns, fonts: assets.fonts, scope: selected.scope, formatLocale: "en-US" },
        batchRows: 3, maxBatchCells: 57, maxChunkBytes: 4096,
        onProgress(progress) { if (progress.phase === "rows") status.textContent = "Exporting " + progress.rows + " services…"; }
      });
      lastExport = { format, ids: selected.ids, keys: selected.columns.map(column => column.key), completed, observed, remainingListeners: listeners.size };
      OfficeIMO.saveBlob(new Blob(chunks, { type: format === "pdf" ? "application/pdf" : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), "operational-inventory." + format);
      status.textContent = "Exported " + selected.ids.length + " selected services.";
    } catch (error) { lastExport = { error: error.message }; status.textContent = error.message; }
    finally { owned?.terminate(); if (url) URL.revokeObjectURL(url); controls.forEach(control => { control.disabled = false; }); }
  }
  // Readable host capture and completion are also available to the opt-in consumer qualification.
  globalThis.operationalInventory = { capture, get lastExport() { return lastExport; } };
  for (const format of ["xlsx", "pdf"]) document.getElementById("inventory-export-" + format).addEventListener("click", () => exportView(format));
});
