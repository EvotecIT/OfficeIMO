// report.worker.ts; compile this module with the application's worker configuration.
import { writeXlsx, writeCsv } from "@evotecit/officeimo";
import type { Column } from "@evotecit/officeimo";

interface Sale { customer: { name: string }; amount: number; seen: Date; }
const columns: readonly Column<Sale>[] = [
  { header: "Customer", value: row => row.customer.name },
  { header: "Amount", key: "amount", type: "number", format: "0.00" },
  { header: "Seen", key: "seen", type: "date", format: "yyyy-mm-dd" }
];
async function* pages(url: string, signal: AbortSignal): AsyncGenerator<Sale> {
  let next: string | null = url;
  while (next) {
    const response: Response = await fetch(next, { signal });
    if (!response.ok) throw new Error("Export source failed: " + response.status);
    const page: { rows: Sale[]; next: string | null } = await response.json();
    for (const row of page.rows) yield { ...row, seen: new Date(row.seen) };
    next = page.next;
  }
}
let active: AbortController | undefined;
self.onmessage = async ({ data }) => {
  if (data.cancel) { active?.abort(); return; }
  if (active) return; // One export at a time in this worker.
  const controller = active = new AbortController();
  try {
    const write = data.format === "csv" ? writeCsv : writeXlsx;
    const blob = await write(pages(data.url, controller.signal), {
      columns, signal: controller.signal,
      limits: { maxRows: 1_000_000, maxCells: 8_000_000, maxOutputBytes: 512_000_000 },
      onProgress: progress => self.postMessage({ progress })
    });
    self.postMessage({ blob });
  } catch (error) { self.postMessage({ error: String(error) }); }
  finally { active = undefined; }
};
