import { analyticalRows, writeAnalyticalReport } from "./analytical-report-job.mjs";
import { reportColumns } from "./report-engine-job.mjs";

// Explicit example resources; report behavior uses stable field keys, not translated headings.
export const wideAnalyticalText = Object.freeze({
  en: Object.freeze({ headers: { protocol: "Protocol", requests: "Requests / min", errorRate: "Error rate", cpu: "CPU usage", memory: "Memory (MiB)",
    storage: "Storage (GiB)", certificateExpiry: "Certificate expiry (UTC)", version: "Version", lastIncident: "Last incident (UTC)", support: "Support tier" },
    scope: "Illustrative selected inventory with 16 typed operational fields, grouped by portfolio, environment and owner.",
    overview: "The PDF repeats service identifiers and group context across readable column bands. Excel retains all 16 editable columns in one sheet." }),
  pl: Object.freeze({ headers: { protocol: "Protokół", requests: "Żądania / min", errorRate: "Odsetek błędów", cpu: "Użycie CPU", memory: "Pamięć (MiB)",
    storage: "Dysk (GiB)", certificateExpiry: "Ważność certyfikatu (UTC)", version: "Wersja", lastIncident: "Ostatni incydent (UTC)", support: "Poziom wsparcia" },
    scope: "Przykładowy wybrany zbiór z 16 polami operacyjnymi, pogrupowany według portfela, środowiska i właściciela.",
    overview: "PDF powtarza identyfikatory usług i kontekst grup w czytelnych pasmach kolumn. Excel zachowuje wszystkie 16 edytowalnych kolumn w jednym arkuszu." })
});

/** The same visible-column capture serves a wide typed workbook and the banded PDF. */
export const wideAnalyticalColumns = Object.freeze([...reportColumns,
  { header: "Protocol", key: "protocol", width: 14 },
  { header: "Requests / min", key: "requests", type: "number", width: 19, format: "#,##0" },
  { header: "Error rate", key: "errorRate", type: "number", width: 16, format: "0.00%" },
  { header: "CPU usage", key: "cpu", type: "number", width: 16, format: "0.0%" },
  { header: "Memory (MiB)", key: "memory", type: "number", width: 18, format: "#,##0" },
  { header: "Storage (GiB)", key: "storage", type: "number", width: 18, format: "#,##0.0" },
  { header: "Certificate expiry (UTC)", key: "certificateExpiry", type: "date", width: 25, format: "yyyy-mm-dd" },
  { header: "Version", key: "version", width: 16 },
  { header: "Last incident (UTC)", key: "lastIncident", type: "date", width: 24, format: "yyyy-mm-dd hh:mm" },
  { header: "Support tier", key: "support", width: 18 }
].map(column => Object.freeze(column)));

/** Lazily extend the ordinary three-level source; grouping metadata stays after visible cells. */
export function* wideAnalyticalRows(count) {
  let index = 0;
  for (const row of analyticalRows(count)) {
    yield [...row.slice(0, reportColumns.length),
      index % 3 === 0 ? "HTTP/3" : "HTTP/2", 1200 + index % 400 * 17,
      index % 37 === 0 ? .015 : index % 13 / 10000, .18 + index % 70 / 100,
      512 + index % 12 * 256, 40 + index % 31 * 2.5,
      new Date(Date.UTC(2027, 0, 1 + index % 28)), "3." + index % 5 + "." + index % 11,
      index % 37 === 0 ? new Date(Date.UTC(2026, 9, 20, index % 24)) : null,
      index % 4 === 0 ? "24 × 7" : "Business hours", ...row.slice(reportColumns.length)];
    index++;
  }
}

/** Choose the public engine's band policy; the example owns no pagination or source replay. */
export async function writeWideAnalyticalReport(api, context) {
  const input = context.input ?? {}, language = input.language ?? "en";
  if (!Object.hasOwn(wideAnalyticalText, language)) throw new TypeError("The wide example supports explicit en and pl resources.");
  const text = wideAnalyticalText[language];
  const columns = wideAnalyticalColumns.map(column => ({ ...column, header: text.headers[column.key] ?? column.header }));
  const decimal = new Intl.NumberFormat(language, { maximumFractionDigits: 3 });
  const percent = new Intl.NumberFormat(language, { style: "percent", maximumFractionDigits: 2 });
  const percentages = new Set(["availability", "errorRate", "cpu"]);
  return writeAnalyticalReport(api, { ...context, columns,
    input: { ...input, scope: input.scope ?? text.scope, overview: input.overview ?? text.overview,
      totals: { latency: "average", availability: "average", requests: "sum", errorRate: "average", cpu: "average", memory: "average", storage: "average" } },
    pdfTableOptions: { wideTable: { mode: "bands", repeatColumns: [0, 1], minimumColumnWidth: 72 },
      formatValue(value, cell) { return value instanceof Date ? value.toISOString().slice(0, 16).replace("T", " ")
        : typeof value === "number" ? (percentages.has(cell.column.key) ? percent : decimal).format(value) : String(value ?? ""); } },
    // These finite physical-work budgets include repeated identifiers and headings, not only logical source cells.
    pdfLimits: { maxTableBands: 8, maxTableWindowRows: 64, maxCells: 16000000, maxTextCharacters: 64000000 } });
}
