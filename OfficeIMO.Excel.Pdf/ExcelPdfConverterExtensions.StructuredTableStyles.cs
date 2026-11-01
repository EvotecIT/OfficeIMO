using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Excel.Pdf {
    public static partial class ExcelPdfConverterExtensions {
        private static IReadOnlyList<StructuredTableVisualData> ReadStructuredTableVisuals(
            ExcelDocument document,
            string sheetName,
            ExcelPdfExportOptions options) {
            var visuals = new List<StructuredTableVisualData>();
            foreach (ExcelTableInfo table in document.GetTables()) {
                if (!string.Equals(table.SheetName, sheetName, StringComparison.OrdinalIgnoreCase) ||
                    !ExcelA1.TryParseRange(table.Range, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                    continue;
                }

                ExcelBuiltInTableStylePalette? palette = null;
                if (options.UseWorksheetCellStyles && !ExcelBuiltInTableStylePaletteResolver.TryCreate(
                        document,
                        table.StyleName,
                        out palette)) {
                    if (!string.IsNullOrWhiteSpace(table.StyleName)) {
                        AddWarning(
                            options,
                            sheetName,
                            "worksheet-table-style",
                            $"Excel table '{table.DisplayName}' uses custom or unknown style '{table.StyleName}'. Its values and direct cell formatting were preserved, but the table style could not be projected.");
                    }
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
                    palette));
            }

            if (options.UseWorksheetCellStyles) AddPivotTableVisuals(document, sheetName, visuals);
            return visuals;
        }

        private static Dictionary<(int Row, int Column), PdfCore.PdfCanvasTableHeaderScope>? CreateStructuredTableHeaderScopes(
            string?[,]? references, IReadOnlyList<StructuredTableVisualData> tables,
            IReadOnlyList<int> rows, int columnOffset, int columns) {
            if (references == null || tables.Count == 0) return null;
            var scopes = new Dictionary<(int Row, int Column), PdfCore.PdfCanvasTableHeaderScope>();
            for (int row = 0; row < rows.Count; row++) {
                for (int column = 0; column < columns; column++) {
                    int sourceColumn = column + columnOffset;
                    if (rows[row] >= references.GetLength(0) || sourceColumn >= references.GetLength(1)) continue;
                    string? reference = references[rows[row], sourceColumn];
                    if (string.IsNullOrEmpty(reference)) continue;
                    var source = ExcelA1.ParseCellRef(reference!);
                    if (tables.Any(table => table.IsHeaderCell(source.Row, source.Col)))
                        scopes[(row, column)] = PdfCore.PdfCanvasTableHeaderScope.Column;
                }
            }
            return scopes.Count == 0 ? null : scopes;
        }

        /// <summary>
        /// A pivot table prints with its pivot style: a coloured header row and a bold grand total
        /// row with a rule above it. The built-in pivot styles share their palettes with the table
        /// style of the same family and number, so the pivot area is drawn with that palette.
        /// </summary>
        private static void AddPivotTableVisuals(ExcelDocument document, string sheetName, List<StructuredTableVisualData> visuals) {
            IReadOnlyList<ExcelPivotTableInfo> pivots;
            try {
                pivots = document.GetPivotTables();
            } catch (InvalidOperationException) {
                return;
            }

            foreach (ExcelPivotTableInfo pivot in pivots) {
                if (!string.Equals(pivot.SheetName, sheetName, StringComparison.OrdinalIgnoreCase) ||
                    string.IsNullOrWhiteSpace(pivot.PivotStyle) ||
                    !pivot.PivotStyle!.StartsWith("PivotStyle", StringComparison.OrdinalIgnoreCase) ||
                    !ExcelA1.TryParseRange(pivot.Location ?? string.Empty, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn) ||
                    lastRow <= firstRow) {
                    continue;
                }

                string tableStyle = "TableStyle" + pivot.PivotStyle.Substring("PivotStyle".Length);
                if (!ExcelBuiltInTableStylePaletteResolver.TryCreate(document, tableStyle, out ExcelBuiltInTableStylePalette? palette)) {
                    continue;
                }

                bool hasGrandTotal = pivot.RowGrandTotals != false && pivot.RowFields.Count > 0;
                visuals.Add(new StructuredTableVisualData(
                    firstRow,
                    firstColumn,
                    lastRow,
                    lastColumn,
                    hasHeader: true,
                    hasTotals: hasGrandTotal,
                    showFirstColumn: false,
                    showLastColumn: false,
                    showRowStripes: false,
                    showColumnStripes: false,
                    palette!));
            }
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

            (int Row, int Col) cell = ExcelA1.ParseCellRef(reference!.Replace("$", string.Empty));
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
            private readonly ExcelBuiltInTableStylePalette? _palette;

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
                ExcelBuiltInTableStylePalette? palette) {
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

            internal bool IsHeaderCell(int row, int column) =>
                _hasHeader && row == _firstRow && column >= _firstColumn && column <= _lastColumn;

            internal bool TryResolve(int row, int column, out StructuredTableCellVisual? visual) {
                visual = null;
                if (_palette == null || row < _firstRow || row > _lastRow || column < _firstColumn || column > _lastColumn) {
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
                bool top = true, bottom = true, left = true, right = true;
                switch (_palette.BorderLayout) {
                    case ExcelTableStyleBorderLayout.Horizontal:
                        left = right = false;
                        break;
                    case ExcelTableStyleBorderLayout.OutlineHorizontal:
                        left = right = false;
                        top = row == _firstRow || totals;
                        bottom = row == _lastRow || header;
                        break;
                    case ExcelTableStyleBorderLayout.Box:
                        top = row == _firstRow || totals;
                        bottom = row == _lastRow || header;
                        left = column == _firstColumn;
                        right = column == _lastColumn;
                        break;
                    case ExcelTableStyleBorderLayout.None:
                        top = bottom = left = right = false;
                        break;
                }

                visual = new StructuredTableCellVisual(
                    fill,
                    header ? _palette.HeaderText : _palette.BodyText,
                    header ? _palette.HeaderBold : emphasizedColumn || totals,
                    top || bottom || left || right ? _palette.Border : null,
                    top, right, bottom, left);
                return true;
            }
        }

        private sealed class StructuredTableCellVisual {
            internal StructuredTableCellVisual(string? fill, string? text, bool bold, string? border,
                bool borderTop = true, bool borderRight = true, bool borderBottom = true, bool borderLeft = true) {
                Fill = fill;
                Text = text;
                Bold = bold;
                Border = border;
                BorderTop = borderTop;
                BorderRight = borderRight;
                BorderBottom = borderBottom;
                BorderLeft = borderLeft;
            }

            internal bool BorderTop { get; }
            internal bool BorderRight { get; }
            internal bool BorderBottom { get; }
            internal bool BorderLeft { get; }

            internal string? Fill { get; }
            internal string? Text { get; }
            internal bool Bold { get; }
            internal string? Border { get; }
        }
    }
}
