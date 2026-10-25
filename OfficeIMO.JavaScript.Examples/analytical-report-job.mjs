// A complete report over one selected, preordered capture. Hosts own selection and query execution.
// Portable rows contain selected visible cells followed by hidden grouping scalars; workers reconstruct getters here.
export const analyticalReportText = Object.freeze({
  en: Object.freeze({ headers: { service: "Service", site: "Site", status: "Status", latency: "Latency (ms)", availability: "Availability", checked: "Checked (UTC)" },
    formatLocale: "en-US", summary: "Summary", operations: { sum: "sum", count: "count", average: "average", min: "minimum", max: "maximum" },
    title: "Selected inventory analysis", context: "Report context", detail: "Grouped inventory", overview: "Group overview",
    group: "Group", records: "Records", subtotal: "Group summary", total: "All selected records", continued: " (continued)",
    next: "Next actions", advice: "Assign owners to the findings and retain the exported scope for the next review.",
    scope: "Selected service inventory for operational review, grouped by the supplied portfolio, environment and owner.",
    summaryIntro: "Summary statistics", sourceValues: "Source names and status values are shown as supplied.",
    missingLatency: "A dash marks an unavailable detail latency.",
    workbookNotes: "Counts in group labels describe the exported selection. Summary formulas recalculate after edits; those count labels remain export-time snapshots.",
    complete: (count, locale) => `This review covers ${count.toLocaleString(locale)} records from the stated selection.`,
    page: { prefix: "Page ", separator: " of ", suffix: "" } }),
  pl: Object.freeze({ headers: { service: "Usługa", site: "Lokalizacja", status: "Stan", latency: "Opóźnienie (ms)", availability: "Dostępność", checked: "Sprawdzone (UTC)" },
    formatLocale: "pl-PL", summary: "Podsumowanie", operations: { sum: "suma", count: "liczba", average: "średnia", min: "minimum", max: "maksimum" },
    title: "Analiza wybranych zasobów", context: "Zakres raportu", detail: "Zasoby według grup", overview: "Podsumowanie grup",
    group: "Grupa", records: "Rekordy", subtotal: "Podsumowanie grupy", total: "Wszystkie wybrane rekordy", continued: " (ciąg dalszy)",
    next: "Dalsze działania", advice: "Przypisz właścicieli do ustaleń i zachowaj zakres eksportu do następnego przeglądu.",
    scope: "Wybrany zbiór usług do przeglądu operacyjnego, pogrupowany według podanego portfela, środowiska i właściciela.",
    summaryIntro: "Statystyki podsumowania", sourceValues: "Nazwy i wartości stanów zachowują treść źródłową.",
    missingLatency: "Kreska oznacza niedostępne opóźnienie w szczegółach.",
    workbookNotes: "Liczby rekordów w etykietach grup opisują wyeksportowany wybór. Formuły podsumowań przeliczają się po edycji; liczby w etykietach pozostają stanem z chwili eksportu.",
    complete: (count, locale) => `Przegląd obejmuje ${count.toLocaleString(locale)} rekordów z podanego zakresu.`,
    page: { prefix: "Strona ", separator: " z ", suffix: "" } })
});

/** Explicit report presentation; spreadsheet formats and tagged language never select a locale. */
function valueFormatter(formatLocale) {
  const number = digits => new Intl.NumberFormat(formatLocale, { minimumFractionDigits: digits, maximumFractionDigits: digits });
  const percent = digits => new Intl.NumberFormat(formatLocale, { style: "percent", minimumFractionDigits: digits, maximumFractionDigits: digits });
  const formats = { latency: number(1), availability: percent(2), requests: number(0), errorRate: percent(2), cpu: percent(1), memory: number(0), storage: number(1) };
  const fallback = new Intl.NumberFormat(formatLocale, { maximumFractionDigits: 3 });
  return (value, context) => value instanceof Date ? value.toISOString().slice(0, context.column.key === "certificateExpiry" ? 10 : 16).replace("T", " ")
    : typeof value === "number" ? (formats[context.column.key] ?? fallback).format(value) : String(value ?? "");
}

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
  const text = analyticalReportText[language], formatLocale = input.formatLocale ?? text.formatLocale;
  const countFormat = new Intl.NumberFormat(formatLocale, { maximumFractionDigits: 0 });
  const keys = input.groupKeys ?? ["portfolio", "environment", "owner"];
  if (!Array.isArray(keys) || keys.length < 1 || keys.length > 4 || keys.some(key => typeof key !== "string" || !key))
    throw new TypeError("Declare one through four stable grouping keys.");
  const capturedColumns = columns.map(column => ({ ...column }));
  const columnCount = capturedColumns.length;
  const totals = Object.freeze({ ...(input.totals ?? Object.fromEntries(["latency", "availability"].filter(key => capturedColumns.some(column => column.key === key)).map(key => [key, "average"]))) });
  const normalHeaders = new Map();
  const selectedColumns = capturedColumns.map((column, index) => {
    const header = input.language !== undefined && text.headers[column.key] ? text.headers[column.key] : column.header;
    normalHeaders.set(column.key, header);
    const operation = Object.hasOwn(totals, column.key) ? totals[column.key] : undefined;
    return { ...column, header: input.format === "pdf" && operation !== undefined ? `${header}\n${text.summary}: ${text.operations[operation]}` : header,
      ...(column.type === "number" && column.alignment === undefined ? { alignment: "right" } : {}),
      // This is detail-only presentation; generated null averages remain blank.
      value: record => {
        const value = record[index];
        return column.key === "latency" && value == null ? new api.ExportCell(value, { text: "—" }) : value;
      } };
  });
  let consumed = 0;
  async function* records() {
    for await (const record of rows) {
      if (!Array.isArray(record) || record.length !== columnCount + keys.length)
        throw new TypeError("An analytical record needs every selected value and grouping field.");
      consumed++; yield record;
    }
  }
  const report = api.createGroupedReport(records(), { columns: selectedColumns, totals, signal,
    groupBy: keys.map((key, level) => capturedColumns.some(column => column.key === key) ? key : { key, value: record => record[columnCount + level] }),
    maxGroups: 4000, limits: { maxRows: 100000, maxCells: 100001 * columnCount, maxTextCharacters: 16000000 } });
  const labels = { labelColumn: { header: text.group, width: 26 },
    subtotalLabel: group => `${text.subtotal} · ${group.label} · ${countFormat.format(group.rows)}`,
    grandTotalLabel: summary => `${text.total} · ${countFormat.format(summary.rows)}`,
    continuationLabel: group => group.label + text.continued };
  const summaryStatistics = report.columns.filter(column => Object.hasOwn(report.totals, column.key))
    .map(column => normalHeaders.get(column.key) + " — " + text.operations[report.totals[column.key]]);
  const summaryLegend = summaryStatistics.length ? text.summaryIntro + ": " + summaryStatistics.join("; ") + "." : "";
  const scope = input.scope ?? text.scope;
  if (input.format === "xlsx") {
    const book = new api.Workbook({ sink, signal, dateMode: "utc", compression: input.compression ?? "auto",
      limits: { maxRows: 112010, maxCells: 112010 * (columnCount + 1), maxOutputBytes: 128 * 1024 * 1024 },
      onProgress: event => onProgress?.({ ...event, rows: consumed }) });
    try {
      const context = book.addWorksheet(text.context, { columns: [{ header: text.context, width: 80, wrapText: true }] });
      await context.addRows([[text.title], [scope], ...(input.overview ? [[input.overview]] : []), ...(summaryLegend ? [[summaryLegend]] : []), [text.workbookNotes], [text.sourceValues]]);
      const grouped = await api.addGroupedXlsx(book, report, { ...labels,
        sheet: { name: text.detail, title: { text: text.title }, freezeHeader: true, freezeColumns: 1, print: { repeatHeaders: true } },
        groupStyle: { font: { bold: true, color: "163C75" }, fill: { color: "DBEAFE" } },
        subtotalStyle: { font: { bold: true }, fill: { color: "EEF2F6" } },
        ...(capturedColumns.some(column => column.key === "status") ? {
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
  const formatValue = valueFormatter(formatLocale);
  async function* blocks() {
    yield { kind: "section", title: text.title, headingStyle: "title" };
    yield { kind: "paragraph", text: scope, style: "body" };
    if (input.overview) yield { kind: "paragraph", text: input.overview, style: "body" };
    yield { kind: "paragraph", text: text.sourceValues + (normalHeaders.has("latency") ? " " + text.missingLatency : ""), style: "body" };
    if (input.figure) yield { kind: "image", image: input.figure.image, width: input.figure.width ?? 420,
      alternativeText: input.figure.alternativeText, caption: { text: input.figure.caption, style: "caption" } };
    yield { kind: "paragraph", text: text.detail, headingLevel: 2, style: "heading", keepTogether: true, keepWithNext: true,
      bookmark: { title: text.detail, level: 2 } };
    yield api.pdfGroupedTable(report, { ...labels, groupHeadingLevel: 3, headerPresentation: { background: "163C75", color: "FFFFFF", bold: true },
      alternateRowColor: "F3F6FB", formatValue, ...pdfTableOptions });
    yield { kind: "group", keepTogether: true, blocks: [
      { kind: "paragraph", text: text.next, headingLevel: 2, style: "heading", bookmark: { title: text.next, level: 2 } },
      { kind: "paragraph", text: text.advice, style: "body" },
      { kind: "paragraph", text: text.complete(consumed, formatLocale), style: "body" }
    ] };
  }
  const options = { signal, onProgress, fonts, tagged: { language }, pageSize: "A4", orientation: "landscape", margins: 30, fontSize: 9, pageNumberLabels: text.page,
    limits: { maxRows: 112010, maxCells: 112010 * (columnCount + 1), maxPages: 20000, maxOutputBytes: 128 * 1024 * 1024,
      maxStructureElements: 2500000, maxMarkedContentReferences: 8000000, ...pdfLimits },
    paragraphStyles: { title: { fontSize: 22, spaceAfter: 12, presentation: { bold: true, color: "163C75" } },
      heading: { fontSize: 15, spaceAfter: 8, presentation: { bold: true, color: "163C75" } },
      body: { fontSize: 10, spaceAfter: 10 }, caption: { fontSize: 9, spaceAfter: 6 } },
    pageHeader: context => context.sectionTitle ?? text.title };
  return sink === undefined ? api.writePdfDocument(blocks(), options) : api.writePdfDocumentTo(blocks(), sink, options);
}
