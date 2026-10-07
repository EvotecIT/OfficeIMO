import { checkAbort, pause, taskYieldDue } from "../../core/iteration.js";
import { ExportBudget } from "../../core/limits.js";
import type { Column, ExportValue } from "../../core/index.js";
import { array, call, indexes, member } from "./api.js";
import { headings, text, value } from "./headings.js";
import type { DataTablesApi, DataTablesExport, DataTablesExportOptions, DataTablesHost, DataTablesOptions } from "./types.js";

function safeOptions(options: DataTablesExportOptions): DataTablesExportOptions {
  const format = options.format ? { ...options.format } : undefined, customize = options.customizeData;
  return { ...options, ...(format ? { format: {
    ...(format.header ? { header: (v: unknown, c: number, n: unknown) => value(format.header!(v, c, n)) } : {}),
    ...(format.footer ? { footer: (v: unknown, c: number, n: unknown) => value(format.footer!(v, c, n)) } : {}),
    ...(format.body ? { body: (v: unknown, r: number, c: number, n: unknown) => value(format.body!(v, r, c, n)) } : {})
  } } : {}), ...(customize ? { customizeData: (data: unknown) => {
    const result: unknown = customize(data);
    // Buttons ignores synchronous return values; only asynchronous callbacks violate this contract.
    if (typeof member(result, "then") === "function") value(result);
  } } : {}) };
}

/** Capture export scope and headings, then produce values in bounded batches using public DataTables APIs. */
export function createDataTablesExport(host: DataTablesHost, table: DataTablesApi, options: DataTablesOptions = {}): DataTablesExport {
  for (const column of Object.values(options.columnOptions ?? {}))
    if ((column as { style?: unknown }).style !== undefined) throw new TypeError("Workbook-local column styles require the advanced Workbook API; use portable ExportCell presentation.");
  const mode = options.mode ?? "batched", headingMode = options.headings ?? "grouped";
  if (!["batched", "compatibility"].includes(mode)) throw new TypeError("Unknown DataTables export mode.");
  if (!["grouped", "leaf", "structured"].includes(headingMode)) throw new TypeError("Unknown DataTables heading mode.");
  if (options.serverSide !== undefined && !["reject", "loaded"].includes(options.serverSide)) throw new TypeError("Unknown server-side export policy.");
  const batchRows = options.batchRows ?? 1024, maxBatchCells = options.maxBatchCells ?? 65536;
  if (!Number.isInteger(batchRows) || batchRows < 1 || batchRows > 4096) throw new RangeError("batchRows must be between 1 and 4,096.");
  if (!Number.isSafeInteger(maxBatchCells) || maxBatchCells < 1) throw new RangeError("maxBatchCells must be a positive safe integer.");
  const signal = options.signal, projectValue = options.project, budget = new ExportBudget(options.limits);
  checkAbort(signal);
  if (member(call(table.page, "info"), "serverSide") && options.serverSide !== "loaded")
    throw new TypeError("Server-side DataTables exports require a separate full-data source or explicit serverSide: 'loaded'.");
  const config = safeOptions(options.exportOptions ?? {});
  const stripOptions = { stripHtml: true, stripNewlines: true, decodeEntities: true, trim: true, ...config };
  const stripOwner = host.Buttons, strip = member(stripOwner, "stripData");
  const stripData = typeof strip === "function" ? (strip as (input: unknown, options: unknown) => unknown).bind(stripOwner) : undefined;
  if (mode === "batched" && !config.format?.body && typeof strip !== "function")
    throw new TypeError("DataTables API requires stripData().");
  if (mode === "batched" && config.customizeData) throw new TypeError("customizeData requires compatibility mode.");
  if (config.customizeData && options.project) throw new TypeError("project cannot identify source indexes after customizeData.");
  const modifier = { search: "applied", order: "applied", ...config.modifier };
  if (member(modifier, "selected") === undefined && typeof member(member(table, "select"), "info") === "function"
    && Number(call(call(table, "rows", config.rows ?? null, { ...modifier, selected: true }), "count")) > 0)
    Object.assign(modifier, { selected: true });
  const rowIndexes = indexes(call(table, "rows", config.rows ?? null, modifier));
  const columnIndexes = indexes(call(table, "columns", config.columns ?? ""));
  if (columnIndexes.length > maxBatchCells) throw new RangeError("A selected row exceeds maxBatchCells.");
  if (new Set(rowIndexes).size !== rowIndexes.length || new Set(columnIndexes).size !== columnIndexes.length)
    throw new TypeError("DataTables export indexes must be unique.");
  const data = call(table.buttons, "exportData", { ...config, modifier, ...(mode === "batched" ? { rows: [] } : {}) });
  const leaf = array(member(data, "header"));
  if (leaf.length !== columnIndexes.length) throw new TypeError("Export header must match the selected columns.");
  const heading = headings(member(data, "headerStructure"), leaf, headingMode);
  for (const index of columnIndexes) if (member(options.columnOptions?.[index], "value") !== undefined)
    throw new TypeError("DataTables columnOptions describes columns; resolve values with project.");
  const columns: readonly Column<readonly ExportValue[]>[] = Object.freeze(leaf.map((label, index) => Object.freeze({
    header: text(label), key: "dt:" + columnIndexes[index]!, ...options.columnOptions?.[columnIndexes[index]!],
    ...(heading.groups[index]!.length ? { groups: Object.freeze(heading.groups[index]!) } : {})
  })));
  // Explicit heading overrides also apply to CSV's leaf heading.
  if (!heading.structure) heading.rows[heading.rows.length - 1] = columns.map(column => column.header);
  const headerStructure = heading.structure?.map((row, level) => Object.freeze(row.map((cell, index) => {
    if (!cell || level + (cell.rowSpan ?? 1) !== heading.structure!.length) return cell;
    const override = options.columnOptions?.[columnIndexes[index]!] ?.header;
    if ((cell.columnSpan ?? 1) > 1 && columnIndexes.slice(index, index + (cell.columnSpan ?? 1)).some(c => options.columnOptions?.[c]?.header !== undefined))
      throw new TypeError("A spanning leaf heading cannot have per-column header overrides.");
    return override === undefined ? cell : Object.freeze({ ...cell, value: override });
  })));
  const footerStructure = member(data, "footerStructure");
  if (headingMode !== "structured" && options.includeFooter !== false && Array.isArray(footerStructure) && footerStructure.length > 1)
    throw new TypeError("Only a single footer row is supported; select includeFooter: false to omit it explicitly.");
  const rawFooter = options.includeFooter === false ? undefined : member(data, "footer");
  const footer = rawFooter == null ? undefined : array(rawFooter).map(value);
  if (footer && footer.length !== columns.length) throw new TypeError("Export footer must match the selected columns.");
  const footerHeading = footer && Array.isArray(footerStructure) && footerStructure.length ? headings(footerStructure, footer, headingMode === "structured" ? "structured" : "grouped") : undefined;
  const body = mode === "compatibility" ? array(member(data, "body")) : undefined;
  const count = columns.length ? body?.length ?? rowIndexes.length : 0;
  budget.check("maxRows", count);
  const batchSize = columns.length ? Math.min(batchRows, Math.floor(maxBatchCells / columns.length)) : batchRows;
  const columnPositions = new Map(columnIndexes.map((index, ordinal) => [index, ordinal]));
  // Keep the public cell API context without selecting or traversing the table's rows.
  const emptyCells = mode === "batched" && count ? call(table, "cells", [], []) : undefined;
  if (emptyCells) {
    let tables = 0;
    call(emptyCells, "iterator", "table", () => { tables++; });
    if (tables !== 1) throw new TypeError("A batched export requires exactly one DataTables table.");
  }
  let consumed = false;
  const rows: AsyncIterable<readonly ExportValue[]> = { [Symbol.asyncIterator]() {
    if (consumed) throw new TypeError("A DataTables export source can be consumed only once.");
    consumed = true; return iterate();
  } };
  async function* iterate(): AsyncGenerator<readonly ExportValue[]> {
    for (let first = 0; first < count; first += batchSize) {
      checkAbort(signal);
      let batch: ExportValue[][];
      if (body) {
        batch = body.slice(first, first + batchSize).map((row, ordinal) => {
          if (!Array.isArray(row) || row.length !== columns.length) throw new TypeError("Export body row must match the selected columns.");
          return row.map((cell, index) => project(cell, rowIndexes[first + ordinal] ?? first + ordinal, index, first + ordinal));
        });
      } else {
        const selectedRows = rowIndexes.slice(first, first + batchSize);
        // Public result-set operations replace the bounded cell indexes. Row selectors would
        // rescan the complete table on every batch, even with constant-time membership.
        const requested = new Array<{ row: number; column: number }>(selectedRows.length * columns.length);
        for (let row = 0, cell = 0; row < selectedRows.length; row++)
          for (const column of columnIndexes) requested[cell++] = { row: selectedRows[row]!, column };
        call(emptyCells, "pop");
        call(emptyCells, "push", requested);
        const cells = emptyCells;
        const rendered = array(call(cells, "render", config.orthogonal ?? "display"));
        const positions = array(call(cells, "indexes"));
        const nodes = config.format?.body ? array(call(cells, "nodes")) : undefined;
        if (rendered.length !== selectedRows.length * columns.length || positions.length !== rendered.length || nodes && nodes.length > rendered.length)
          throw new TypeError("The table changed or returned an incomplete export batch.");
        batch = selectedRows.map(() => new Array<ExportValue>(columns.length));
        // The ordinary public result set preserves the requested order. Verify it before
        // using ordinals; adapters returning another order retain coordinate-based mapping.
        const ordered = positions.every((position, i) => member(position, "row") === requested[i]!.row
          && member(position, "column") === requested[i]!.column);
        const rowPositions = ordered ? undefined : new Map(selectedRows.map((index, ordinal) => [index, ordinal]));
        const seen = ordered ? undefined : new Set<number>();
        for (let cell = 0; cell < rendered.length; cell++) {
          checkAbort(signal);
          const rowIndex = ordered ? requested[cell]!.row : member(positions[cell], "row"), columnIndex = ordered ? requested[cell]!.column : member(positions[cell], "column");
          const row = ordered ? Math.floor(cell / columns.length) : rowPositions!.get(rowIndex as number);
          const column = ordered ? cell % columns.length : columnPositions.get(columnIndex as number);
          if (row === undefined || column === undefined || seen?.has(row * columns.length + column)) throw new TypeError("Invalid DataTables cell indexes.");
          seen?.add(row * columns.length + column);
          let node = nodes?.[cell];
          if (nodes && nodes.length !== rendered.length) {
            // nodes() omits deferred cells without DOM nodes. Resolve each coordinate
            // through the same bounded result set so compacted nodes cannot shift rows.
            call(cells, "pop");
            call(cells, "push", [{ row: rowIndex, column: columnIndex }]);
            const exact = array(call(cells, "nodes"));
            if (exact.length > 1) throw new TypeError("Invalid DataTables cell nodes.");
            node = exact[0];
          }
          const formatted = config.format?.body ? config.format.body(rendered[cell], rowIndex as number, columnIndex as number, node)
            : stripData!(rendered[cell], stripOptions);
          batch[row]![column] = project(formatted, rowIndex as number, column, first + row);
          if ((cell & 127) === 127 && taskYieldDue()) { await pause(); checkAbort(signal); }
        }
      }
      for (const row of batch) { checkAbort(signal); yield row; }
      if (taskYieldDue()) { await pause(); checkAbort(signal); }
    }
  }
  function project(input: unknown, rowIndex: number, column: number, rowOrdinal: number): ExportValue {
    checkAbort(signal);
    const scalar = value(input);
    return projectValue ? value(projectValue(scalar, { sourceRowIndex: rowIndex, sourceColumnIndex: columnIndexes[column]!, rowIndex: rowOrdinal, columnIndex: column })) : scalar;
  }
  return Object.freeze({ columns, headers: Object.freeze(heading.rows.map(row => Object.freeze(row))),
    footer: footer ? Object.freeze(footer) : undefined,
    ...(headerStructure ? { headerStructure: Object.freeze(headerStructure) } : {}),
    ...(footerHeading?.structure ? { footerStructure: footerHeading.structure } : {}), rowCount: count, rows });
}
