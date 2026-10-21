// Fixture declarations and public operation wiring; PowerForge owns all measurements.
export const groupedFixtures = [
  { name: "grouped-analysis-xlsx-100000-sink", format: "xlsx", rows: 100000, groupRows: 2500, delivery: "sink", kind: "grouped-analysis", currentOnly: true },
  { name: "grouped-analysis-pdf-5000-sink", format: "pdf", rows: 5000, groupRows: 250, delivery: "sink", kind: "grouped-analysis", currentOnly: true }
];

const columns = [
  { key: "service", header: "Service", width: 18 }, { key: "site", header: "Site", width: 14 },
  { key: "latency", header: "Latency (ms)", type: "number", format: "0.000000", width: 18 },
  { key: "availability", header: "Availability", type: "number", format: "0.000000", width: 18 },
  { key: "requests", header: "Requests", type: "number", width: 16 },
  { key: "checked", header: "Checked (UTC)", type: "date", format: "yyyy-mm-dd", width: 20 }
];

/** The first entire group has null metrics; remaining groups contain deterministic numeric/null mixtures. */
export function groupedRow(index, fixture) {
  const group = Math.floor(index / fixture.groupRows), empty = group === 0;
  return { service: "SRV-" + String(index + 1).padStart(6, "0"), site: "Site" + String(group).padStart(2, "0"),
    latency: empty || index % 17 === 0 ? null : index % 97 / 10 + 4.75,
    availability: empty || index % 23 === 0 ? null : (9950 - index % 17) / 10000,
    requests: empty || index % 29 === 0 ? null : index % 31,
    checked: new Date(Date.UTC(2026, 9, 20, 0, index % 60)) };
}

export function* groupedRows(fixture) { for (let index = 0; index < fixture.rows; index++) yield groupedRow(index, fixture); }

/** Six typed report values, finite editable group/grand totals and ordinary untagged PDF composition. */
export async function writeGroupedWorkload(api, fixture, { rows = groupedRows(fixture), sink, signal, input = {}, onProgress } = {}) {
  const source = api.createGroupedReport(rows, { columns, groupBy: "site", totals: { latency: "average", availability: "min", requests: "sum" },
    maxGroups: 1000, ...(signal ? { signal } : {}) });
  const labels = { labelColumn: { header: "Group", width: 24 }, subtotalLabel: group => "Subtotal-" + group.label,
    grandTotalLabel: () => "Grand-total", ...(signal ? { signal } : {}), ...(onProgress ? { onProgress } : {}) };
  if (fixture.format === "xlsx") return api.writeGroupedXlsxTo(source, sink, { ...labels,
    dateMode: "utc", compression: input.compression ?? "auto", sheet: { name: "Analysis", freezeHeader: true } });
  return api.writeGroupedPdfTo(source, sink, { ...labels,
    pageSize: "A3", orientation: "landscape", fontSize: 8, padding: 2, pageNumbers: false,
    compression: input.compression !== "store", columnWidths: [130, 110, 90, 120, 120, 95, 130],
    continuationLabel: group => group.label + " continued",
    formatValue: value => value instanceof Date ? value.toISOString().slice(0, 10) : value == null ? "-" : typeof value === "number" ? String(Number(value.toFixed(6))) : String(value) });
}
