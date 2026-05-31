using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.Runtime.CompilerServices;

namespace OfficeIMO.Word.Pdf {
    internal static class TableLayoutCache {
        private static readonly ConditionalWeakTable<WordTable, TableLayout> _cache = new();

        internal static TableLayout GetLayout(WordTable table) {
            if (_cache.TryGetValue(table, out TableLayout? existingLayout) && existingLayout != null) {
                return existingLayout;
            }

            List<IReadOnlyList<WordTableCell>> rows = TableBuilder.Map(table).ToList();
            int columnCount = ResolveColumnCount(table, rows);
            float[] widths = new float[columnCount];

            List<int> gridColumnWidths = table.GridColumnWidth;
            if (gridColumnWidths.Count > 0) {
                for (int i = 0; i < widths.Length && i < gridColumnWidths.Count; i++) {
                    widths[i] = gridColumnWidths[i] / 20f;
                }
            }

            foreach (IReadOnlyList<WordTableCell> row in rows) {
                int logicalColumn = 0;
                for (int i = 0; i < row.Count && logicalColumn < widths.Length; i++) {
                    WordTableCell cell = row[i];
                    if (cell.HorizontalMerge == MergedCellValues.Continue) {
                        continue;
                    }

                    int columnSpan = System.Math.Max(1, cell.ColumnSpan);
                    float width = 0f;
                    if (cell.Width.HasValue && cell.WidthType == TableWidthUnitValues.Dxa) {
                        width = cell.Width.Value / 20f;
                    }

                    if (cell.HasNestedTables) {
                        foreach (WordTable nested in cell.NestedTables) {
                            TableLayout nestedLayout = GetLayout(nested);
                            float nestedWidth = nestedLayout.ColumnWidths.Sum();
                            if (nestedWidth > width) {
                                width = nestedWidth;
                            }
                        }
                    }

                    if (width > 0f) {
                        float widthPerColumn = width / columnSpan;
                        for (int columnIndex = logicalColumn; columnIndex < logicalColumn + columnSpan && columnIndex < widths.Length; columnIndex++) {
                            if (widthPerColumn > widths[columnIndex]) {
                                widths[columnIndex] = widthPerColumn;
                            }
                        }
                    }

                    logicalColumn += columnSpan;
                }
            }

            TableLayout layout = new(rows, widths);
            _cache.Add(table, layout);
            return layout;
        }

        private static int ResolveColumnCount(WordTable table, List<IReadOnlyList<WordTableCell>> rows) {
            List<int> gridColumnWidths = table.GridColumnWidth;
            if (gridColumnWidths.Count > 0) {
                return gridColumnWidths.Count;
            }

            int columnCount = 0;
            foreach (IReadOnlyList<WordTableCell> row in rows) {
                int rowColumns = 0;
                foreach (WordTableCell cell in row) {
                    if (cell.HorizontalMerge == MergedCellValues.Continue) {
                        continue;
                    }

                    rowColumns += System.Math.Max(1, cell.ColumnSpan);
                }

                if (rowColumns > columnCount) {
                    columnCount = rowColumns;
                }
            }

            return columnCount;
        }
    }
}

