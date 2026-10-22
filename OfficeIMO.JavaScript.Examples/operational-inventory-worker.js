// A classic Blob worker hosts the ordinary public grouped writers. It owns no query or codec.
OfficeIMO.handleWorkerExports(self, async ({ input, rows, sink, signal, onProgress }) => {
  const api = OfficeIMO, columns = input.columns.map((column, index) => ({ ...column, value: row => row[index] }));
  const keys = ["portfolio", "environment", "owner"];
  const maxRows = 100;
  const report = api.createGroupedReport(rows, {
    columns, groupBy: keys.map((key, index) => ({ key, value: row => row[columns.length + index] })),
    totals: Object.fromEntries([["latency", "average"], ["availability", "average"], ["requests", "sum"]]
      .filter(([key]) => columns.some(column => column.key === key))),
    signal, maxGroups: 100, limits: { maxRows, maxCells: columns.length * (maxRows + 1) }
  });
  const options = {
    signal, onProgress, labelColumn: { header: "Group", width: 27 },
    subtotalLabel: group => "Group total · " + group.label + " · " + group.rows,
    grandTotalLabel: summary => "All selected services · " + summary.rows,
    continuationLabel: group => group.label + " (continued)",
    limits: { maxRows: 200, maxCells: 100000, maxPages: 2000, maxOutputBytes: 16 * 1024 * 1024 }
  };
  if (input.format === "xlsx") return api.writeGroupedXlsxTo(report, sink, {
    ...options, dateMode: "utc", sheet: { name: "Selected inventory", title: { text: input.scope },
      freezeHeader: true, freezeColumns: 1, print: { orientation: "landscape", repeatHeaders: true } },
    groupStyle: { font: { bold: true, color: "163C75" }, fill: { color: "DBEAFE" } },
    subtotalStyle: { font: { bold: true }, fill: { color: "EEF2F6" } }
  });
  if (input.format !== "pdf") throw new TypeError("Choose Excel or PDF.");
  const font = name => Uint8Array.from(atob(input.fonts[name]), character => character.charCodeAt(0));
  const fonts = { regular: await api.PdfFont.create(font("regular"), { signal }), bold: await api.PdfFont.create(font("bold"), { signal }) };
  const percentages = new Set(["availability", "errorRate", "cpu"]);
  const decimal = new Intl.NumberFormat("en", { maximumFractionDigits: 3 });
  const percent = new Intl.NumberFormat("en", { style: "percent", maximumFractionDigits: 2 });
  return api.writeGroupedPdfTo(report, sink, {
    ...options, fonts, title: input.scope, tagged: { language: "en" }, pageSize: "A4", orientation: "landscape", margins: 30, fontSize: 9,
    wideTable: { mode: "bands", minimumColumnWidth: 72, repeatColumns: [0, "service"] },
    headerPresentation: { background: "163C75", color: "FFFFFF", bold: true }, alternateRowColor: "F3F6FB",
    formatValue(value, context) {
      return value instanceof Date ? value.toISOString().slice(0, 16).replace("T", " ")
        : typeof value === "number" ? (percentages.has(context.column.key) ? percent : decimal).format(value) : String(value ?? "");
    }
  });
});
