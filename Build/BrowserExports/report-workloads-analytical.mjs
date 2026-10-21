// Case/operation declarations only. The shared PowerForge runner owns measurement policy.
import { readFile } from "node:fs/promises";
import { analyticalRows, writeAnalyticalReport } from "../../OfficeIMO.JavaScript.Examples/analytical-report-job.mjs";
import { reportColumns } from "../../OfficeIMO.JavaScript.Examples/report-engine-job.mjs";

export const analyticalFixtures = [
  { name: "analytical-xlsx-60000-sink", kind: "analytical", format: "xlsx", rows: 60000, delivery: "sink", currentOnly: true },
  { name: "analytical-xlsx-100000-sink", kind: "analytical", format: "xlsx", rows: 100000, delivery: "sink", currentOnly: true },
  { name: "analytical-pdf-10000-sink", kind: "analytical", format: "pdf", rows: 10000, delivery: "sink", currentOnly: true }
];

/** Font intake and module loading are setup, outside the measured export operation. */
export async function analyticalSetup(fixture) {
  if (fixture.format !== "pdf") return {};
  const root = new URL("../../Website/Apps/OfficeIMO.Web.Converter/Assets/Fonts/", import.meta.url);
  return { fonts: { regular: new Uint8Array(await readFile(new URL("Carlito-Regular.ttf", root))),
    bold: new Uint8Array(await readFile(new URL("Carlito-Bold.ttf", root))) } };
}

export function analyticalInput(fixture) { return analyticalRows(fixture.rows); }

/** The real composition job uses one captured stream, native editing rules and a post-detail overview/chart. */
export function writeAnalyticalWorkload(api, fixture, { rows = analyticalInput(fixture), sink, signal, onProgress, input = {} } = {}) {
  return writeAnalyticalReport(api, { rows, columns: reportColumns, sink, signal, onProgress,
    input: { format: fixture.format, language: "en", ...input } });
}
