// A complete report over one selected, preordered capture. Hosts own selection and query execution.
// Portable rows contain selected visible cells followed by hidden grouping scalars; workers reconstruct getters here.
export const analyticalReportText = Object.freeze({
  en: Object.freeze({ headers: { service: "Service", site: "Site", status: "Status", latency: "Latency (ms)", availability: "Availability", checked: "Checked (UTC)" },
    title: "Selected inventory analysis", context: "Report context", detail: "Grouped inventory", overview: "Group overview",
    group: "Group", records: "Records", subtotal: "Group total", total: "All selected records", continued: " (continued)",
    next: "Next actions", advice: "Assign owners to the findings and retain the exported scope for the next review.",
    scope: "Illustrative selected inventory. Grouping fields are metadata; detail values remain typed.",
    complete: count => `${count.toLocaleString("en")} selected records were consumed once.`,
    page: { prefix: "Page ", separator: " of ", suffix: "" } }),
  pl: Object.freeze({ headers: { service: "Usługa", site: "Lokalizacja", status: "Stan", latency: "Opóźnienie (ms)", availability: "Dostępność", checked: "Sprawdzone (UTC)" },
    title: "Analiza wybranych zasobów", context: "Zakres raportu", detail: "Zasoby według grup", overview: "Podsumowanie grup",
    group: "Grupa", records: "Rekordy", subtotal: "Suma grupy", total: "Wszystkie wybrane rekordy", continued: " (ciąg dalszy)",
    next: "Dalsze działania", advice: "Przypisz właścicieli do ustaleń i zachowaj zakres eksportu do następnego przeglądu.",
    scope: "Przykładowy wybrany zbiór zasobów. Pola grupowania są metadanymi; wartości szczegółowe zachowują typy.",
    complete: count => `Wybrane rekordy przetworzone jednokrotnie: ${count.toLocaleString("pl")}.`,
    page: { prefix: "Strona ", separator: " z ", suffix: "" } })
});

/** This sample uses three hidden dimensions; its ordered generator never constructs a complete data array. */
export function* analyticalRows(count) {
  if (!Number.isSafeInteger(count) || count < 0 || count > 100000) throw new RangeError("Choose 0 through 100,000 records.");
  const leafRows = Math.max(1, Math.ceil(count / 24));
  for (let index = 0; index < count; index++) {
    const leaf = Math.floor(index / leafRows);
    yield ["WEB-" + String(index + 1).padStart(6, "0"), ["Warsaw", "Łódź", "Praha", "Berlin"][leaf % 4],
        index % 37 === 0 ? "Offline" : index % 11 === 0 ? "Attention" : "Healthy",
        index % 17 === 0 ? null : index % 97 / 10 + 4.75,
        index % 23 === 0 ? null : .995 - index % 17 / 10000,
        new Date(Date.UTC(2026, 9, 21, 8, index % 60)),
      "Portfolio " + Math.floor(leaf / 12), ["Production", "Staging", "Development"][Math.floor(leaf / 4) % 3], "Owner " + leaf % 4];
  }
}

/** Compose existing workbook/PDF owners; there is no query, sorting, format codec or transport implementation here. */
export async function writeAnalyticalReport(api, { input = {}, rows, columns, sink, signal, onProgress, pdfTableOptions = {}, pdfLimits = {} }) {
  const language = input.language ?? "en";
  if (!Object.hasOwn(analyticalReportText, language)) throw new TypeError("This example has explicit en and pl resources.");
  const text = analyticalReportText[language], keys = input.groupKeys ?? ["portfolio", "environment", "owner"];
  if (!Array.isArray(keys) || keys.length < 1 || keys.length > 4 || keys.some(key => typeof key !== "string" || !key))
    throw new TypeError("Declare one through four stable grouping keys.");
  const selectedColumns = columns.map((column, index) => ({ ...column,
    ...(input.language !== undefined && text.headers[column.key] ? { header: text.headers[column.key] } : {}),
    value: record => record[index] }));
  const totals = input.totals ?? Object.fromEntries(["latency", "availability"].filter(key => columns.some(column => column.key === key)).map(key => [key, "average"]));
  let consumed = 0;
  async function* records() {
    for await (const record of rows) {
      if (!Array.isArray(record) || record.length !== columns.length + keys.length)
        throw new TypeError("An analytical record needs every selected value and grouping field.");
      consumed++; yield record;
    }
  }
  const report = api.createGroupedReport(records(), { columns: selectedColumns, totals, signal,
    groupBy: keys.map((key, level) => columns.some(column => column.key === key) ? key : { key, value: record => record[columns.length + level] }),
    maxGroups: 4000, limits: { maxRows: 100000, maxCells: 100001 * columns.length, maxTextCharacters: 16000000 } });
  const labels = { labelColumn: { header: text.group, width: 26 },
    subtotalLabel: group => `${text.subtotal} · ${group.label} · ${group.rows.toLocaleString(language)}`,
    grandTotalLabel: summary => `${text.total} · ${summary.rows.toLocaleString(language)}`,
    continuationLabel: group => group.label + text.continued };
  const scope = input.scope ?? text.scope;
  if (input.format === "xlsx") {
    const book = new api.Workbook({ sink, signal, dateMode: "utc", compression: input.compression ?? "auto",
      limits: { maxRows: 112010, maxCells: 112010 * (columns.length + 1), maxOutputBytes: 128 * 1024 * 1024 },
      onProgress: event => onProgress?.({ ...event, rows: consumed }) });
    try {
      const context = book.addWorksheet(text.context, { columns: [{ header: text.context, width: 80, wrapText: true }] });
      await context.addRows([[text.title], [scope], ...(input.overview ? [[input.overview]] : [])]);
      const grouped = await api.addGroupedXlsx(book, report, { ...labels,
        sheet: { name: text.detail, title: { text: text.title }, freezeHeader: true, freezeColumns: 1, print: { repeatHeaders: true } },
        groupStyle: { font: { bold: true, color: "163C75" }, fill: { color: "DBEAFE" } },
        subtotalStyle: { font: { bold: true }, fill: { color: "EEF2F6" } },
        ...(columns.some(column => column.key === "status") ? {
          conditionalFormats: [{ type: "cellIs", range: { column: "status" }, operator: "equal", value: "Offline", style: { font: { bold: true, color: "991B1B" } } }],
          dataValidation: [{ type: "list", column: "status", items: ["Healthy", "Attention", "Offline"], allowBlank: false }]
        } : {}) });
      const overview = book.addWorksheet(text.overview, { columns: [{ header: text.group, key: "group", width: 30 }, { header: text.records, key: "records", type: "number", width: 16 }] });
      const parents = grouped.groups.filter(group => group.level === 0);
      if (parents.length) overview.addChart({ type: "columnClustered", title: text.overview,
        data: { kind: "range", sheet: overview, rowCount: parents.length, categoryColumn: "group", series: [{ column: "records" }] },
        row: 2, column: 4, width: 640, height: 340 });
      await overview.addRows(parents.map(group => [group.label, group.rows]));
      const result = sink === undefined ? await book.toBlob() : await book.finish();
      return sink === undefined ? result : { rows: grouped.rows, columns: grouped.columns, bytes: result.bytes };
    } catch (error) { await book.discard(error); throw error; }
  }
  if (input.format !== "pdf") throw new TypeError("Analytical reports support Excel and PDF.");
  if (!input.fonts?.regular || !input.fonts?.bold) throw new TypeError("The analytical PDF example needs regular and bold font bytes.");
  const fonts = { regular: await api.PdfFont.create(input.fonts.regular, { signal }), bold: await api.PdfFont.create(input.fonts.bold, { signal }) };
  const decimal = new Intl.NumberFormat(language, { maximumFractionDigits: 3 });
  async function* blocks() {
    yield { kind: "section", title: text.title, headingStyle: "title" };
    yield { kind: "paragraph", text: scope, style: "body" };
    if (input.overview) yield { kind: "paragraph", text: input.overview, style: "body" };
    if (input.figure) yield { kind: "image", image: input.figure.image, width: input.figure.width ?? 420,
      alternativeText: input.figure.alternativeText, caption: { text: input.figure.caption, style: "caption" } };
    yield { kind: "section", title: text.detail, level: 2, headingStyle: "heading", orientation: "landscape" };
    yield api.pdfGroupedTable(report, { ...labels, groupHeadingLevel: 3, headerPresentation: { background: "163C75", color: "FFFFFF", bold: true },
      alternateRowColor: "F3F6FB", formatValue(value) { return value instanceof Date ? value.toISOString().slice(0, 16).replace("T", " ")
        : typeof value === "number" ? decimal.format(value) : String(value ?? ""); }, ...pdfTableOptions });
    yield { kind: "section", title: text.next, level: 2, headingStyle: "heading" };
    yield { kind: "paragraph", text: text.advice, style: "body" };
    yield { kind: "paragraph", text: text.complete(consumed), style: "body" };
  }
  const options = { signal, onProgress, fonts, tagged: { language }, pageSize: "A4", margins: 30, fontSize: 9, pageNumberLabels: text.page,
    limits: { maxRows: 112010, maxCells: 112010 * (columns.length + 1), maxPages: 20000, maxOutputBytes: 128 * 1024 * 1024,
      maxStructureElements: 2500000, maxMarkedContentReferences: 8000000, ...pdfLimits },
    paragraphStyles: { title: { fontSize: 22, spaceAfter: 12, presentation: { bold: true, color: "163C75" } },
      heading: { fontSize: 15, spaceAfter: 8, presentation: { bold: true, color: "163C75" } },
      body: { fontSize: 10, spaceAfter: 10 }, caption: { fontSize: 9, spaceAfter: 6 } },
    pageHeader: context => context.sectionTitle ?? text.title };
  return sink === undefined ? api.writePdfDocument(blocks(), options) : api.writePdfDocumentTo(blocks(), sink, options);
}
