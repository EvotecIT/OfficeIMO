// A classic Blob worker hosts the ordinary public grouped writers. It owns no query or codec.
OfficeIMO.handleWorkerExports(self, async ({ input, rows, sink, signal, onProgress }) => {
  const api = OfficeIMO, text = {
    formatLocale: "en-US", summary: "Summary", operations: { sum: "sum", count: "count", average: "average", min: "minimum", max: "maximum" },
    title: "Operational inventory review", subtotal: "Group summary",
    workbookNotes: "Counts in group labels describe the exported selection. Summary formulas recalculate after edits; those count labels remain export-time snapshots.",
    sourceValues: "Source names and status values are shown as supplied.",
    missingLatency: "A dash marks an unavailable detail latency.",
    advice: "Review services with unavailable readings or adverse status, assign follow-up owners and retain this selection for the next review."
  };
  const formatLocale = input.formatLocale ?? text.formatLocale;
  const number = digits => new Intl.NumberFormat(formatLocale, { minimumFractionDigits: digits, maximumFractionDigits: digits });
  const percent = digits => new Intl.NumberFormat(formatLocale, { style: "percent", minimumFractionDigits: digits, maximumFractionDigits: digits });
  const formats = { latency: number(1), availability: percent(2), requests: number(0), errorRate: percent(2), cpu: percent(1), memory: number(0), storage: number(1) };
  const decimal = new Intl.NumberFormat(formatLocale, { maximumFractionDigits: 3 });
  const capturedColumns = input.columns.map(column => ({ ...column }));
  const totals = Object.freeze(Object.fromEntries([["latency", "average"], ["availability", "average"], ["requests", "sum"]]
    .filter(([key]) => capturedColumns.some(column => column.key === key))));
  const columns = capturedColumns.map((column, index) => ({ ...column,
    header: input.format === "pdf" && Object.hasOwn(totals, column.key) ? `${column.header}\n${text.summary}: ${text.operations[totals[column.key]]}` : column.header,
    ...(column.type === "number" && column.alignment === undefined ? { alignment: "right" } : {}),
    value: row => {
      const value = row[index];
      return column.key === "latency" && value == null ? new api.ExportCell(value, { text: "—" }) : value;
    } }));
  const keys = ["portfolio", "environment", "owner"];
  const maxRows = 100;
  let consumed = 0;
  async function* records() { for await (const row of rows) { consumed++; yield row; } }
  const report = api.createGroupedReport(records(), {
    columns, groupBy: keys.map((key, index) => ({ key, value: row => row[columns.length + index] })),
    totals,
    signal, maxGroups: 100, limits: { maxRows, maxCells: columns.length * (maxRows + 1) }
  });
  const summaryLegend = "Summary statistics: " + report.columns.filter(column => Object.hasOwn(report.totals, column.key))
    .map(column => capturedColumns.find(source => source.key === column.key).header + " — " + text.operations[report.totals[column.key]]).join("; ") + ".";
  const labels = {
    labelColumn: { header: "Group", width: 27 },
    subtotalLabel: group => text.subtotal + " · " + group.label + " · " + group.rows,
    grandTotalLabel: summary => "All selected services · " + summary.rows,
    continuationLabel: group => group.label + " (continued)"
  };
  const options = {
    ...labels, signal, onProgress,
    limits: { maxRows: 200, maxCells: 100000, maxPages: 2000, maxOutputBytes: 16 * 1024 * 1024 }
  };
  if (input.format === "xlsx") {
    const book = new api.Workbook({ sink, signal, dateMode: "utc", limits: options.limits,
      onProgress: event => onProgress?.({ ...event, rows: consumed }) });
    try {
      const context = book.addWorksheet("Report context", { columns: [{ header: "Report context", width: 110, wrapText: true }] });
      await context.addRows([[text.title], [input.scope], [summaryLegend], [text.workbookNotes], [text.sourceValues]]);
      const grouped = await api.addGroupedXlsx(book, report, {
        ...labels, sheet: { name: "Selected inventory", title: { text: "Selected operational services" },
          freezeHeader: true, freezeColumns: 1, print: { orientation: "landscape", repeatHeaders: true } },
        groupStyle: { font: { bold: true, color: "163C75" }, fill: { color: "DBEAFE" } },
        subtotalStyle: { font: { bold: true }, fill: { color: "EEF2F6" } }
      });
      const completed = await book.finish();
      return { rows: grouped.rows, columns: grouped.columns, bytes: completed.bytes };
    } catch (error) { await book.discard(error); throw error; }
  }
  if (input.format !== "pdf") throw new TypeError("Choose Excel or PDF.");
  const font = name => Uint8Array.from(atob(input.fonts[name]), character => character.charCodeAt(0));
  const fonts = { regular: await api.PdfFont.create(font("regular"), { signal }), bold: await api.PdfFont.create(font("bold"), { signal }) };
  return api.writeGroupedPdfTo(report, sink, {
    ...options, fonts, title: text.title, messageTop: input.scope + "\n" + text.sourceValues + (capturedColumns.some(column => column.key === "latency") ? " " + text.missingLatency : ""), messageBottom: text.advice,
    tagged: { language: "en" }, pageSize: "A4", orientation: "landscape", margins: 30, fontSize: 9,
    wideTable: { mode: "bands", minimumColumnWidth: 72, repeatColumns: [0, "service"] },
    headerPresentation: { background: "163C75", color: "FFFFFF", bold: true }, alternateRowColor: "F3F6FB",
    formatValue(value, context) {
      return value instanceof Date ? value.toISOString().slice(0, context.column.key === "certificateExpiry" ? 10 : 16).replace("T", " ")
        : typeof value === "number" ? (formats[context.column.key] ?? decimal).format(value) : String(value ?? "");
    }
  });
});
