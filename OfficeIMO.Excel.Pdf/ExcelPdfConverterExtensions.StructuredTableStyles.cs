using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Excel.Pdf {
    public static partial class ExcelPdfConverterExtensions {
        private static IReadOnlyList<StructuredTableVisualData> ReadStructuredTableVisuals(
            ExcelDocument document,
            string sheetName,
            ExcelToPdfOptions options) {
            if (!options.UseWorksheetCellStyles) {
                return Array.Empty<StructuredTableVisualData>();
            }

            var visuals = new List<StructuredTableVisualData>();
            foreach (ExcelTableInfo table in document.GetTables()) {
                if (!string.Equals(table.SheetName, sheetName, StringComparison.OrdinalIgnoreCase) ||
                    !A1.TryParseRange(table.Range, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                    continue;
                }

                if (!ExcelBuiltInTableStylePaletteResolver.TryCreate(
                        document,
                        table.StyleName,
                        out ExcelBuiltInTableStylePalette? palette)) {
                    if (!string.IsNullOrWhiteSpace(table.StyleName)) {
                        AddWarning(
                            options,
                            sheetName,
                            "WorksheetTableStyle",
                            $"Excel table '{table.DisplayName}' uses custom or unknown style '{table.StyleName}'. Its values and direct cell formatting were preserved, but the table style could not be projected.");
                    }
                    continue;
                }

                visuals.Add(new StructuredTableVisualData(
                    firstRow,
                    firstColumn,
                    lastRow,
                    lastColumn,
                    table.HasHeaderRow,
                    table.TotalsRowShown,
                    table.ShowFirstColumn,
                    table.ShowLastColumn,
                    table.ShowRowStripes,
                    table.ShowColumnStripes,
                    palette!));
            }

            return visuals;
        }

        private static StructuredTableCellVisual? GetStructuredTableCellVisual(
            IReadOnlyList<StructuredTableVisualData>? tables,
            string?[,]? cellReferences,
            int row,
            int column) {
            if (tables == null || tables.Count == 0 || cellReferences == null ||
                row < 0 || column < 0 ||
                row >= cellReferences.GetLength(0) || column >= cellReferences.GetLength(1)) {
                return null;
            }

            string? reference = cellReferences[row, column];
            if (string.IsNullOrWhiteSpace(reference)) {
                return null;
            }

            (int Row, int Col) cell = A1.ParseCellRef(reference!.Replace("$", string.Empty));
            if (cell.Row <= 0 || cell.Col <= 0) {
                return null;
            }

            for (int index = tables.Count - 1; index >= 0; index--) {
                if (tables[index].TryResolve(cell.Row, cell.Col, out StructuredTableCellVisual? visual)) {
                    return visual;
                }
            }

            return null;
        }

        private sealed class StructuredTableVisualData {
            private readonly int _firstRow;
            private readonly int _firstColumn;
            private readonly int _lastRow;
            private readonly int _lastColumn;
            private readonly bool _hasHeader;
            private readonly bool _hasTotals;
            private readonly bool _showFirstColumn;
            private readonly bool _showLastColumn;
            private readonly bool _showRowStripes;
            private readonly bool _showColumnStripes;
            private readonly ExcelBuiltInTableStylePalette _palette;

            internal StructuredTableVisualData(
                int firstRow,
                int firstColumn,
                int lastRow,
                int lastColumn,
                bool hasHeader,
                bool hasTotals,
                bool showFirstColumn,
                bool showLastColumn,
                bool showRowStripes,
                bool showColumnStripes,
                ExcelBuiltInTableStylePalette palette) {
                _firstRow = firstRow;
                _firstColumn = firstColumn;
                _lastRow = lastRow;
                _lastColumn = lastColumn;
                _hasHeader = hasHeader;
                _hasTotals = hasTotals;
                _showFirstColumn = showFirstColumn;
                _showLastColumn = showLastColumn;
                _showRowStripes = showRowStripes;
                _showColumnStripes = showColumnStripes;
                _palette = palette;
            }

            internal bool TryResolve(int row, int column, out StructuredTableCellVisual? visual) {
                visual = null;
                if (row < _firstRow || row > _lastRow || column < _firstColumn || column > _lastColumn) {
                    return false;
                }

                bool header = _hasHeader && row == _firstRow;
                bool totals = _hasTotals && row == _lastRow;
                int bodyRow = row - _firstRow - (_hasHeader ? 1 : 0);
                int tableColumn = column - _firstColumn;
                string? fill = header ? _palette.HeaderFill : _palette.BodyFill;
                if (!header && !totals) {
                    if (_showRowStripes && bodyRow >= 0 && bodyRow % 2 == 0) {
                        fill = _palette.StripeFill ?? fill;
                    }
                    if (_showColumnStripes && tableColumn % 2 == 0) {
                        fill = _palette.StripeFill ?? fill;
                    }
                }

                bool emphasizedColumn =
                    _showFirstColumn && column == _firstColumn ||
                    _showLastColumn && column == _lastColumn;
                visual = new StructuredTableCellVisual(
                    fill,
                    header ? _palette.HeaderText : _palette.BodyText,
                    header ? _palette.HeaderBold : emphasizedColumn,
                    _palette.Border);
                return true;
            }
        }

        private sealed class StructuredTableCellVisual {
            internal StructuredTableCellVisual(string? fill, string? text, bool bold, string? border) {
                Fill = fill;
                Text = text;
                Bold = bold;
                Border = border;
            }

            internal string? Fill { get; }
            internal string? Text { get; }
            internal bool Bold { get; }
            internal string? Border { get; }
        }
    }
}
