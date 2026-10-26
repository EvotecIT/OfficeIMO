import { Workbook, addGroupedXlsx } from "../../xlsx/index.js";
import type { XlsxOptions } from "../../xlsx/index.js";
import { portableWorkbook } from "../../xlsx/portable.js";
import type { GroupedReportSource } from "../../core/grouped-report.js";
import type { ExportResult, OutputDestination } from "../../core/index.js";
import { withDestination } from "../../core/sinks.js";
import type { CanopyExport, CanopyWriteOptions } from "./types.js";

/** Compose only the adapter's data/grouped sheet and optional metadata sheet in the existing workbook owner. */
export async function writeCanopyWorkbook(source: CanopyExport, grouped: GroupedReportSource | undefined,
  destination: OutputDestination, options: XlsxOptions, labels: CanopyWriteOptions["grouped"],
  about: readonly (readonly [string, string | number])[] | undefined): Promise<ExportResult> {
  const { sheet: sheetSettings = {}, columns: _columns, ...settings } = options;
  const { name = "Data", ...sheet } = sheetSettings;
  if (about && name.toLowerCase() === "about") throw new TypeError("The optional About worksheet needs a distinct data-sheet name.");
  if (grouped && settings.cellValueWriters !== undefined) throw new TypeError("Grouped Canopy exports use resolved source values rather than custom cellValueWriters.");
  return withDestination(destination, async sink => {
    const workbookOptions = { ...portableWorkbook(settings), sink };
    // Grouped worksheets generate native FormulaCells; flat tables keep the portable-value boundary.
    const book = grouped ? new Workbook(workbookOptions) : Workbook.forTable(workbookOptions);
    try {
      let rows: number, columns: number;
      if (grouped) {
        const { conditionalFormats, ...groupedSheet } = sheet;
        const written = await addGroupedXlsx(book, grouped, {
          labelColumn: { header: labels?.labelColumn?.header ?? "Group", width: 18 },
          subtotalLabel: labels?.subtotalLabel ?? (group => group.label + " total"),
          grandTotalLabel: labels?.grandTotalLabel ?? (() => "Grand total"),
          ...(labels?.includeGrandTotal === undefined ? {} : { includeGrandTotal: labels.includeGrandTotal }),
          sheet: { ...groupedSheet, name }, ...(conditionalFormats === undefined ? {} : { conditionalFormats })
        });
        rows = written.rows; columns = written.columns;
      } else {
        const data = book.addWorksheet(name, { boldHeader: true, autoFilter: sheet.includeHeader !== false,
          autoSize: { minWidth: 6, maxWidth: 54 }, ...sheet, columns: source.columns });
        await data.addRows(source.rows()); await data.close();
        rows = data.rowCount; columns = source.columns.length;
      }
      if (about) {
        const metadata = book.addWorksheet("About", { columns: [{ header: "Field", width: 18 }, { header: "Value", width: 72, wrapText: true }],
          boldHeader: true, freezeHeader: true });
        await metadata.addRows(about); await metadata.close();
      }
      const result = await book.finish();
      return { rows, columns, bytes: result.bytes };
    } catch (error) { await book.discard(error); throw error; }
  });
}
