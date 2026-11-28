using System.Xml;
using OfficeIMO.Excel.Utilities;
using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Projection {
    internal static partial class XlsbWorksheetGridProjector {
        private static void WriteArrayGrid(XmlWriter writer, XlsbSheet source, Func<string, int?> sharedStringIndex,
            Action<string?, bool> recordMutation, bool includeLayout, XlsbStylesheet? styles, ArrayProjection arrays, CancellationToken cancellationToken) {
            writer.WriteStartElement("sheetData", Namespace);
            IReadOnlyList<ArrayCell> arrayCells = arrays.Cells;
            int cellIndex = 0, rowIndex = 0, arrayIndex = 0;
            IReadOnlyList<XlsbRowInfo> rows = includeLayout ? source.Rows : Array.Empty<XlsbRowInfo>();
            while (cellIndex < source.Cells.Count || rowIndex < rows.Count || arrayIndex < arrayCells.Count) {
                cancellationToken.ThrowIfCancellationRequested();
                XlsbRowInfo? metadata = rowIndex < rows.Count ? rows[rowIndex] : null;
                int cellRow = cellIndex < source.Cells.Count ? source.Cells[cellIndex].Row : int.MaxValue;
                int arrayRow = arrayIndex < arrayCells.Count ? arrayCells[arrayIndex].Row : int.MaxValue;
                int row = Math.Min(Math.Min(cellRow, arrayRow), metadata?.Row ?? int.MaxValue);
                writer.WriteStartElement("row", Namespace);
                writer.WriteAttributeString("r", InvariantNumberText.Get(row));
                if (metadata?.Row == row) {
                    XlsbWorksheetGeometryMarkup.WriteRowAttributes(writer, metadata);
                    rowIndex++;
                }
                if (row == arrayRow) {
                    WriteArrayRow(row);
                } else {
                    while (cellIndex < source.Cells.Count && source.Cells[cellIndex].Row == row) {
                        cancellationToken.ThrowIfCancellationRequested();
                        XlsbCell cell = source.Cells[cellIndex++];
                        WriteCell(writer, cell, sharedStringIndex, recordMutation, includeLayout, styles);
                    }
                }
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
            arrays.RecordOwners(recordMutation);

            void WriteArrayRow(int row) {
                while (true) {
                    int sourceColumn = cellIndex < source.Cells.Count && source.Cells[cellIndex].Row == row
                        ? source.Cells[cellIndex].Column : int.MaxValue;
                    int arrayColumn = arrayIndex < arrayCells.Count && arrayCells[arrayIndex].Row == row
                        ? arrayCells[arrayIndex].Column : int.MaxValue;
                    int column = Math.Min(sourceColumn, arrayColumn);
                    if (column == int.MaxValue) break;
                    cancellationToken.ThrowIfCancellationRequested();

                    XlsbCell? cell = column == sourceColumn ? source.Cells[cellIndex++] : null;
                    if (column == arrayColumn) {
                        ArrayFormula? owner = arrayCells[arrayIndex++].Owner;
                        WriteArrayCell(writer, cell, row, column, sharedStringIndex, recordMutation, includeLayout, styles, owner);
                    } else {
                        WriteCell(writer, cell!, sharedStringIndex, recordMutation, includeLayout, styles);
                    }
                }
            }
        }

        private static void WriteArrayCell(XmlWriter writer, XlsbCell? cell, int row, int column, Func<string, int?> sharedStringIndex,
            Action<string?, bool> recordMutation, bool richText, XlsbStylesheet? styles,
            ArrayFormula? owner) {
            string reference = ExcelA1.CellReference(row, column);
            (string? type, string? value) = cell == null ? (null, null) : CellValue(cell, sharedStringIndex);
            string? sourceFormula = string.IsNullOrWhiteSpace(cell?.FormulaText) ? null : ExcelSanitizer.SanitizeFormula(cell!.FormulaText!);
            string? formula = owner?.Text;
            writer.WriteStartElement("c", Namespace);
            writer.WriteAttributeString("r", reference);
            if (cell != null) writer.WriteAttributeString("s", InvariantNumberText.Get(cell.StyleIndex));
            // BIFF12 text runs project to an inline string. Keep the initial shared-string
            // lookup because it determines indices of later plain strings in the same workbook.
            bool inline = richText && cell is { TextRuns: not null, FormulaBytes: null, Value: string };
            if (inline) {
                type = "inlineStr";
                value = null;
            }
            if (type != null) writer.WriteAttributeString("t", type);
            if (formula != null) {
                writer.WriteStartElement("f", Namespace);
                if (owner != null) {
                    writer.WriteAttributeString("t", "array");
                    writer.WriteAttributeString("ref", owner.Reference);
                    owner.RetainsCache = value != null;
                }
                if (value != null) writer.WriteAttributeString("ca", "1");
                writer.WriteString(formula);
                writer.WriteEndElement();
            }
            if (inline) WriteRichText(writer, (string)cell!.Value!, cell.TextRuns!, styles);
            else if (value != null) writer.WriteElementString("v", Namespace, value);
            writer.WriteEndElement();

            if (cell != null) recordMutation(sourceFormula == null ? null : reference, value != null || inline);
        }

        private readonly record struct ArrayCell(int Row, int Column, ArrayFormula? Owner);

        private sealed class ArrayFormula {
            internal ArrayFormula(XlsbCellRange range, string formula) {
                Anchor = ExcelA1.CellReference(range.FirstRow, range.FirstColumn);
                Reference = range.ToA1Reference();
                Text = ExcelSanitizer.SanitizeFormula(formula);
            }

            internal string Anchor { get; }
            internal string Reference { get; }
            internal string Text { get; }
            internal bool RetainsCache { get; set; }
        }

        /// <summary>Supplies the declared cells, including any members absent from the native cell records.</summary>
        private sealed class ArrayProjection {
            private readonly List<ArrayFormula> _owners = new();
            internal List<ArrayCell> Cells { get; } = new();

            internal static ArrayProjection? Create(XlsbSheet sheet, CancellationToken cancellationToken) {
                ArrayProjection? projection = null;
                foreach (XlsbFormulaGroup group in sheet.FormulaGroups) {
                    if (group.Type != 426) continue;
                    cancellationToken.ThrowIfCancellationRequested();
                    string? formula = FindAnchorFormula(sheet.Cells, group.AnchorRow, group.AnchorColumn);
                    if (string.IsNullOrWhiteSpace(formula)) continue;

                    projection ??= new ArrayProjection();
                    XlsbCellRange range = group.Range;
                    var owner = new ArrayFormula(range, formula!);
                    projection._owners.Add(owner);
                    // The workbook preflight bounded the total projected cells before opening the streamed part.
                    for (int row = range.FirstRow; row <= range.LastRow; row++) {
                        for (int column = range.FirstColumn; column <= range.LastColumn; column++) {
                            cancellationToken.ThrowIfCancellationRequested();
                            projection.Cells.Add(new ArrayCell(row, column,
                                row == range.FirstRow && column == range.FirstColumn ? owner : null));
                        }
                    }
                }
                projection?.Cells.Sort(static (left, right) => Compare(left.Row, left.Column, right.Row, right.Column));
                return projection;
            }

            internal void RecordOwners(Action<string?, bool> recordMutation) {
                foreach (ArrayFormula owner in _owners) recordMutation(owner.Anchor, owner.RetainsCache);
            }

            private static string? FindAnchorFormula(IReadOnlyList<XlsbCell> cells, int row, int column) {
                int minimum = 0, maximum = cells.Count - 1;
                while (minimum <= maximum) {
                    int middle = minimum + (maximum - minimum) / 2;
                    XlsbCell cell = cells[middle];
                    int order = Compare(cell.Row, cell.Column, row, column);
                    if (order == 0) return cell.FormulaText;
                    if (order < 0) minimum = middle + 1;
                    else maximum = middle - 1;
                }
                return null;
            }

            private static int Compare(int leftRow, int leftColumn, int rightRow, int rightColumn) =>
                leftRow == rightRow ? leftColumn.CompareTo(rightColumn) : leftRow.CompareTo(rightRow);
        }
    }
}
