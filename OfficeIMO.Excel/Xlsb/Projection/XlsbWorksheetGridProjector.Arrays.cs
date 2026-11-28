using OfficeIMO.Excel.Utilities;
using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Projection {
    internal static partial class XlsbWorksheetGridProjector {
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
