// This example projects the displayed DOM. Grid hosts provide their own current-view iterator.
document.addEventListener("DOMContentLoaded", () => {
  const table = document.querySelector("table"), status = document.getElementById("status");
  const header = table.tHead.rows[0], body = table.tBodies[0];
  const dataRows = () => [...body.rows];
  let descending = false;
  document.getElementById("filter").addEventListener("input", event => {
    const search = event.target.value.toLocaleLowerCase();
    for (const row of dataRows()) row.hidden = !row.textContent.toLocaleLowerCase().includes(search);
  });
  function columnIndex(name) { return [...header.cells].findIndex(cell => cell.textContent.trim() === name); }
  document.getElementById("sort").addEventListener("click", () => {
    descending = !descending;
    const index = columnIndex("Name");
    dataRows().sort((a, b) => a.cells[index].textContent.localeCompare(b.cells[index].textContent) * (descending ? -1 : 1))
      .forEach(row => body.append(row));
  });
  document.getElementById("hide-site").addEventListener("click", () => {
    const index = columnIndex("Site"), hide = !header.cells[index].hidden;
    for (const row of [header, ...dataRows()]) row.cells[index].hidden = hide;
  });
  document.getElementById("move-site").addEventListener("click", () => {
    const index = columnIndex("Site");
    for (const row of [header, ...dataRows()]) {
      const site = row.cells[index];
      if (index === 0) row.append(site); else row.prepend(site);
    }
  });
  async function exportView(format) {
    const buttons = [...document.querySelectorAll("button")];
    buttons.forEach(button => { button.disabled = true; });
    status.textContent = "Exporting…";
    try {
      const indexes = [...header.cells].map((cell, index) => cell.hidden ? -1 : index).filter(index => index >= 0);
      if (!indexes.length) throw new Error("Show at least one column.");
      const columns = indexes.map(index => {
        const name = header.cells[index].textContent.trim();
        return { header: name, type: name === "Latency (ms)" ? "number" : "string", width: 20,
          ...(name === "Latency (ms)" ? { format: "0.00" } : {}) };
      });
      function* rows() {
        for (const row of dataRows()) if (!row.hidden) yield indexes.map((index, column) => {
          const text = row.cells[index].textContent;
          return columns[column].type === "number" ? Number(text) : text;
        });
      }
      let blob;
      if (format === "xlsx") {
        const book = new OfficeIMO.Workbook({ creator: "HtmlForgeX example" });
        const headerStyle = book.styles.add({ font: { bold: true, color: "FFFFFF" }, fill: { color: "203864" }, verticalAlignment: "center" });
        await book.addWorksheet("Current view", { columns, freezeHeader: true, freezeColumns: 1,
          table: { name: "CurrentView", style: "TableStyleMedium9" }, headerStyle, headerHeight: 28, rowHeight: 22,
          cellStyle: ({ column, value }) => column.header === "Latency (ms)" && value > 20
            ? { fill: { color: "FCE4D6" }, font: { color: "9C0006", bold: true } } : undefined }).addRows(rows());
        blob = await book.toBlob();
      } else blob = await OfficeIMO.writeCsv(rows(), { columns, bom: true });
      OfficeIMO.saveBlob(blob, "current-view." + format);
      status.textContent = "Exported " + dataRows().filter(row => !row.hidden).length + " rows.";
    } catch (error) { status.textContent = error.message; }
    finally { buttons.forEach(button => { button.disabled = false; }); }
  }
  document.getElementById("export-xlsx").addEventListener("click", () => exportView("xlsx"));
  document.getElementById("export-csv").addEventListener("click", () => exportView("csv"));
});
