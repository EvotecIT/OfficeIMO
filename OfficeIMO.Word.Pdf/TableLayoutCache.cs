using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;

namespace OfficeIMO.Word.Pdf {
    internal static class TableLayoutCache {
        private const int MaxTableGridColumns = 16_384;
        private const int MaxTableNestingDepth = 128;
        private static readonly ConditionalWeakTable<WordTable, TableLayout> _cache = new();

        internal static TableLayout GetLayout(WordTable table) => GetLayout(table, depth: 0);

        private static TableLayout GetLayout(WordTable table, int depth) {
            if (depth >= MaxTableNestingDepth) {
                throw new InvalidDataException($"Table nesting exceeds the supported limit of {MaxTableNestingDepth} levels.");
            }

            if (_cache.TryGetValue(table, out TableLayout? existingLayout) && existingLayout != null) {
                return existingLayout;
            }

            List<IReadOnlyList<WordTableCell>> rows = WordTableMatrix.Map(table).ToList();
            int[] rowStartColumns = ResolveRowGridOffsets(table, rows.Count, before: true);
            int[] rowTrailingColumns = ResolveRowGridOffsets(table, rows.Count, before: false);
            int columnCount = ResolveColumnCount(table, rows, rowStartColumns, rowTrailingColumns);
            EnsureSupportedColumnCount(columnCount);
            float[] widths = new float[columnCount];
            float[] gridWidths = new float[columnCount];
            bool[] explicitCellWidthColumns = new bool[columnCount];

            List<int> gridColumnWidths = table.GridColumnWidth;
            if (gridColumnWidths.Count > 0) {
                for (int i = 0; i < gridWidths.Length && i < gridColumnWidths.Count; i++) {
                    gridWidths[i] = gridColumnWidths[i] / 20f;
                    if (gridWidths[i] > 0f) {
                        widths[i] = gridWidths[i];
                    }
                }
            }

            for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
                IReadOnlyList<WordTableCell> row = rows[rowIndex];
                int logicalColumn = rowStartColumns[rowIndex];
                for (int i = 0; i < row.Count && logicalColumn < widths.Length; i++) {
                    WordTableCell cell = row[i];
                    if (cell.HorizontalMerge == MergedCellValues.Continue) {
                        continue;
                    }

                    int columnSpan = System.Math.Max(1, cell.ColumnSpan);
                    float width = 0f;
                    bool hasExplicitCellWidth = false;
                    if (IsExplicitDxaCellWidth(cell)) {
                        width = cell.Width!.Value / 20f;
                        hasExplicitCellWidth = true;
                    }

                    foreach (WordTable nested in cell.DirectNestedTables) {
                        TableLayout nestedLayout = GetLayout(nested, depth + 1);
                        float nestedWidth = nestedLayout.ColumnWidths.Sum();
                        if (nestedWidth > width) {
                            width = nestedWidth;
                        }
                    }

                    if (width > 0f) {
                        float widthPerColumn = width / columnSpan;
                        for (int columnIndex = logicalColumn; columnIndex < logicalColumn + columnSpan && columnIndex < widths.Length; columnIndex++) {
                            if (hasExplicitCellWidth) {
                                if (!explicitCellWidthColumns[columnIndex] || widthPerColumn > widths[columnIndex]) {
                                    widths[columnIndex] = widthPerColumn;
                                }

                                explicitCellWidthColumns[columnIndex] = true;
                            } else if (!explicitCellWidthColumns[columnIndex] && widthPerColumn > widths[columnIndex]) {
                                widths[columnIndex] = widthPerColumn;
                            }
                        }
                    }

                    logicalColumn += columnSpan;
                }
            }

            for (int i = 0; i < widths.Length; i++) {
                if (widths[i] <= 0f) {
                    widths[i] = gridWidths[i];
                }
            }

            TableLayout layout = new(rows, widths, rowStartColumns, rowTrailingColumns);
            _cache.Add(table, layout);
            return layout;
        }

        private static int ResolveColumnCount(WordTable table, List<IReadOnlyList<WordTableCell>> rows, int[] rowStartColumns, int[] rowTrailingColumns) {
            List<int> gridColumnWidths = table.GridColumnWidth;
            if (gridColumnWidths.Count > 0) {
                EnsureSupportedColumnCount(gridColumnWidths.Count);
                return gridColumnWidths.Count;
            }

            int columnCount = 0;
            for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
                IReadOnlyList<WordTableCell> row = rows[rowIndex];
                long rowColumns = (long)rowStartColumns[rowIndex] + rowTrailingColumns[rowIndex];
                foreach (WordTableCell cell in row) {
                    if (cell.HorizontalMerge == MergedCellValues.Continue) {
                        continue;
                    }

                    rowColumns += System.Math.Max(1, cell.ColumnSpan);
                    if (rowColumns > MaxTableGridColumns) {
                        throw new InvalidDataException($"The table grid exceeds the supported limit of {MaxTableGridColumns} columns.");
                    }
                }

                if (rowColumns > columnCount) {
                    columnCount = (int)rowColumns;
                }
            }

            return columnCount;
        }

        private static int[] ResolveRowGridOffsets(WordTable table, int rowCount, bool before) {
            int[] offsets = new int[rowCount];
            for (int rowIndex = 0; rowIndex < rowCount && rowIndex < table.Rows.Count; rowIndex++) {
                TableRowProperties? properties = table.Rows[rowIndex]._tableRow.TableRowProperties;
                offsets[rowIndex] = before
                    ? ToNonNegativeInt(properties?.GetFirstChild<GridBefore>()?.Val?.Value)
                    : ToNonNegativeInt(properties?.GetFirstChild<GridAfter>()?.Val?.Value);
            }

            return offsets;
        }

        private static int ToNonNegativeInt(int? value) {
            if (!value.HasValue || value.Value <= 0) {
                return 0;
            }

            EnsureSupportedColumnCount(value.Value);
            return value.Value;
        }

        private static void EnsureSupportedColumnCount(int columnCount) {
            if (columnCount > MaxTableGridColumns) {
                throw new InvalidDataException($"The table grid exceeds the supported limit of {MaxTableGridColumns} columns.");
            }
        }

        private static bool IsExplicitDxaCellWidth(WordTableCell cell) =>
            cell.Width.HasValue &&
            cell.WidthType == TableWidthUnitValues.Dxa &&
            cell.Width.Value > 0;
    }
}
