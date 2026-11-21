using System.Globalization;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Excel.Pdf {
    public static partial class ExcelPdfConverterExtensions {
        /// <summary>Reports clipped worksheet cells once per contiguous range of a sheet, located by A1 reference.</summary>
        private sealed class WorksheetCellClippingSummary : PdfCore.PdfTableCellDiagnosticSummary {
            internal override string FormatSource(PdfCore.PdfTableCellDiagnosticRange range) {
                string first = FormatCell(range.FirstRow, range.FirstColumn);
                string cells = range.FirstRow == range.LastRow && range.FirstColumn == range.LastColumn
                    ? first : first + ":" + FormatCell(range.LastRow, range.LastColumn);
                return FormatSheetName(range.Scope) + "!" + cells;
            }

            internal override string FormatMessage(PdfCore.PdfTableCellDiagnosticRange range) {
                string samples = string.Join(", ", range.Samples.Select(cell => FormatCell(cell.Row, cell.Column)));
                string count = range.CellCount.ToString(CultureInfo.InvariantCulture);
                string cells = range.CellCount == 1
                    ? "1 cell (" + samples + ")"
                    : count + " cells (" + samples + (range.CellCount > range.Samples.Count ? ", ..." : string.Empty) + ")";
                return cells + ". First cell: " + range.FirstMessage;
            }

            private static string FormatCell(int row, int column) =>
                ExcelA1.ColumnIndexToLetters(column) + row.ToString(CultureInfo.InvariantCulture);

            private static string FormatSheetName(string name) {
                bool plain = name.Length > 0 && !char.IsDigit(name[0]) && name.All(character => char.IsLetterOrDigit(character) || character == '_' || character == '.')
                    && ExcelA1.ParseCellRef(name).Row == 0;
                return plain ? name : "'" + name.Replace("'", "''") + "'";
            }
        }

        /// <summary>
        /// Names each worksheet table cell by its sheet cell, for summarized clipping findings. A merged cell, including a
        /// fragment of one continued from an earlier page, is named by the merge's own top-left cell, which holds its text.
        /// </summary>
        private static Func<int, int, PdfCore.PdfTableCellDiagnosticAddress?>? CreateWorksheetCellAddress(string sheetName,
            string?[,]? references, MergeLayoutData? merges) {
            if (references == null) return null;
            int rows = references.GetLength(0), columns = references.GetLength(1);
            return (row, column) => {
                if (row < 0 || row >= rows || column < 0 || column >= columns) return null;
                string? reference = merges?.GetRegion(row, column)?.Span.ImageAnchorReference ?? references[row, column];
                if (string.IsNullOrEmpty(reference)) return null;
                (int sheetRow, int sheetColumn) = ExcelA1.ParseCellRef(reference!);
                return sheetRow > 0 ? new PdfCore.PdfTableCellDiagnosticAddress(sheetName, sheetRow, sheetColumn) : null;
            };
        }
    }
}
