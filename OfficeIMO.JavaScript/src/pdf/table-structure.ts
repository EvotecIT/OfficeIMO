import type { PdfLayoutCell, PdfLayoutRow } from "./table-layout.js";
import type { PdfStructure } from "./structure.js";

interface ColumnHeader {
  readonly row: number;
  readonly order: number;
  readonly identifier: string;
}

/** @internal Source-cell identities survive page fragments and horizontal band projections.
 * Only column headers and still-owned row objects remain reachable from this table owner. */
export class PdfTableStructure {
  private readonly cells = new WeakMap<PdfLayoutCell, number>();
  private readonly rows = new WeakMap<PdfLayoutRow, number>();
  private readonly columnHeaders: ColumnHeader[][] = [];
  /** Ordinary cells without row headers share immutable dictionary references, bounded by column count. */
  private readonly ordinaryAttributes: string[] = [];
  private headerCount = 0;
  private headersPrepared = false;
  constructor(private readonly structure: PdfStructure, readonly table: number,
    private readonly headers: readonly PdfLayoutRow[]) {}

  /** Allocate the complete logical header grid before its first visual band is painted. */
  async prepareHeaders(): Promise<void> {
    if (this.headersPrepared) return;
    for (let rowIndex = 0; rowIndex < this.headers.length; rowIndex++) {
      const row = this.headers[rowIndex]!, parent = this.structure.element("TR", this.table);
      this.rows.set(row, parent);
      for (const cell of row.cells) {
        const header = cell.lines.some(line => line.text.trim().length > 0);
        const associations = this.associations(cell, rowIndex);
        const id = this.structure.element(header ? "TH" : "TD", parent,
          this.attributes(cell, associations, header ? "Column" : undefined));
        this.cells.set(cell, id);
        if (header) {
          const heading = { row: rowIndex, order: this.headerCount++, identifier: await this.structure.identifier(id) };
          for (let column = cell.first; column < cell.first + cell.span; column++)
            (this.columnHeaders[column] ??= []).push(heading);
        }
      }
    }
    this.headersPrepared = true;
  }

  /** Call after every initial physical band; repeated visual headings then become artifacts. */
  async completeHeaders(): Promise<void> {
    for (const row of this.headers) await this.completeRow(row);
  }

  /** Register one original row before projecting or painting it. Summary label headings are explicit. */
  async prepareRow(row: PdfLayoutRow, _kind: "data" | "footer" = "data", parent = this.table,
    rowHeaderIndex?: 0): Promise<readonly number[]> {
    const existing = this.rows.get(row);
    if (existing !== undefined) return row.cells.map(cell => this.cells.get(cell)!);
    const rowId = this.structure.element("TR", parent);
    this.rows.set(row, rowId);
    // Allocate in source order, then resolve every row header before associating even earlier data cells.
    // Generated summary labels override detail roles so numeric subtotal cells never become row headers.
    const ids = row.cells.map((cell, index) => {
      const header = rowHeaderIndex === undefined ? cell.rowHeader === true : index === rowHeaderIndex;
      const id = this.structure.element(header ? "TH" : "TD", rowId);
      this.cells.set(cell, id); return id;
    });
    let rowHeaders: string[] | undefined;
    for (let index = 0; index < ids.length; index++)
      if (rowHeaderIndex === undefined ? row.cells[index]!.rowHeader === true : index === rowHeaderIndex)
        (rowHeaders ??= []).push(await this.structure.identifier(ids[index]!));
    for (let index = 0; index < row.cells.length; index++) {
      const cell = row.cells[index]!;
      if (!rowHeaders && cell.span === 1 && (cell.rowSpan ?? 1) === 1) {
        const attributes = this.ordinaryAttributes[cell.first] ??=
          await this.structure.attributeReference(this.attributeDictionary(cell, this.associations(cell)));
        this.structure.setAttributes(ids[index]!, attributes);
        continue;
      }
      const associations = this.associations(cell);
      const header = rowHeaderIndex === undefined ? cell.rowHeader === true : index === rowHeaderIndex;
      if (!header && rowHeaders) associations.push(...rowHeaders);
      this.structure.setAttributes(ids[index]!, this.attributes(cell, associations, header ? "Row" : undefined));
    }
    return ids;
  }

  owner(cell: PdfLayoutCell): number | undefined { return this.cells.get(cell); }

  async completeRow(row: PdfLayoutRow): Promise<void> {
    await this.structure.complete(this.rows.get(row));
  }

  async finish(): Promise<void> {
    await this.completeHeaders();
    await this.structure.complete(this.table);
  }

  private associations(cell: PdfLayoutCell, beforeRow = Number.POSITIVE_INFINITY): string[] {
    // Ordinary cells visit only their heading depth, independent of the table's width.
    if (cell.span === 1) return (this.columnHeaders[cell.first] ?? []).filter(header => header.row < beforeRow).map(header => header.identifier);
    const headers = new Set<ColumnHeader>();
    for (let column = cell.first; column < cell.first + cell.span; column++)
      for (const header of this.columnHeaders[column] ?? []) if (header.row < beforeRow) headers.add(header);
    // Spans can encounter a shared ancestor through several columns. Preserve logical allocation
    // order, rather than column visitation order, and name that ancestor exactly once.
    return [...headers].sort((left, right) => left.order - right.order).map(header => header.identifier);
  }

  private attributes(cell: PdfLayoutCell, headers: readonly string[], scope?: "Column" | "Row"): string {
    return " /A " + this.attributeDictionary(cell, headers, scope);
  }

  private attributeDictionary(cell: PdfLayoutCell, headers: readonly string[], scope?: "Column" | "Row"): string {
    return "<< /O /Table" + (scope === undefined ? "" : " /Scope /" + scope) +
      (cell.span > 1 ? " /ColSpan " + cell.span : "") + ((cell.rowSpan ?? 1) > 1 ? " /RowSpan " + cell.rowSpan : "") +
      (headers.length ? " /Headers [" + headers.map(id => "(" + id + ")").join(" ") + "]" : "") + " >>";
  }
}
