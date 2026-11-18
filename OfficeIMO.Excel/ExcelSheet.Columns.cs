using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Concurrent;
using System.Globalization;
using System.Xml;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private double CalculateColumnWidth(int columnIndex) {
            return CalculateColumnWidths([columnIndex], CancellationToken.None)[0];
        }

        private const double MaxExcelColumnWidth = 255.0;

        /// <summary>
        /// Returns the stored width, padding included, of a column that keeps the worksheet's default width.
        /// </summary>
        /// <remarks>
        /// Excel reads a <c>col</c> element without <c>width</c> as a zero-width, hidden column, and the XLSB and XLS
        /// writers carry that width into their column records. A column created only to carry a style, an outline
        /// level or visibility therefore states the default width explicitly, without <c>customWidth</c>.
        /// </remarks>
        internal double GetDefaultStoredColumnWidth() {
            SheetFormatProperties? format = WorksheetRoot.GetFirstChild<SheetFormatProperties>();
            if (format?.DefaultColumnWidth?.Value is double configured && configured > 0D) {
                return configured;
            }

            double digit = OfficeIMO.Excel.Utilities.ExcelColumnGeometry.GetMaximumDigitWidth(WorksheetPart);
            int pixels = OfficeIMO.Excel.Utilities.ExcelColumnGeometry.GetWidthPixels(null, format, digit);
            return OfficeIMO.Excel.Utilities.ExcelColumnGeometry.ToStoredWidth(pixels, digit);
        }

        /// <summary>
        /// Applies a change to the worksheet default width and moves the columns that only carry the previous default,
        /// such as those created for visibility or outline levels, to the new default. Custom and best-fit widths stay.
        /// </summary>
        private void ChangeDefaultColumnWidth(Action change) {
            Columns? columns = WorksheetRoot.GetFirstChild<Columns>();
            double previous = columns == null ? 0D : GetDefaultStoredColumnWidth();
            change();
            if (columns == null) return;

            double current = GetDefaultStoredColumnWidth();
            if (current == previous) return;
            foreach (Column column in columns.Elements<Column>()) {
                if (IsDefaultWidthColumn(column) && column.Width?.Value == previous) {
                    column.Width = current;
                }
            }
        }

        /// <summary>Creates a single-column definition that keeps the worksheet's default width.</summary>
        private Column CreateDefaultWidthColumn(uint columnIndex) {
            var column = new Column { Min = columnIndex, Max = columnIndex, Width = GetDefaultStoredColumnWidth() };
            _excelDocument.DefaultWidthColumns.Add(column, column);
            return column;
        }

        /// <summary>Whether the column is one this sheet created with the default width and nobody has resized.</summary>
        private bool IsDefaultWidthColumn(Column column) =>
            _excelDocument.DefaultWidthColumns.TryGetValue(column, out _) && column.CustomWidth?.Value != true && column.BestFit == null;

        /// <summary>Whether a column's width is only the default width this sheet gives to columns it creates.</summary>
        private bool HasOnlyDefaultWidth(Column column) =>
            column.CustomWidth == null && column.BestFit == null
            && (column.Width == null || (IsDefaultWidthColumn(column) && column.Width.Value == GetDefaultStoredColumnWidth()));

        private static double NormalizeColumnWidth(double width) {
            if (double.IsNaN(width) || double.IsInfinity(width)) {
                return 0;
            }

            if (width <= 0) {
                return 0;
            }

            return Math.Min(width, MaxExcelColumnWidth);
        }

        private void SetColumnWidthCore(int columnIndex, double width) {
            var worksheet = WorksheetRoot;
            var columns = worksheet.GetFirstChild<Columns>();
            if (columns == null) {
                columns = worksheet.InsertAt(new Columns(), 0);
            }

            SetColumnWidthCore(columns, columnIndex, width);

            if (columns.Elements<Column>().Any()) {
                ReorderColumns(columns);
            } else {
                columns.Remove();
            }
        }

        private void SetColumnWidthsCore(IReadOnlyList<int> columnIndexes, double[] widths, bool bestFit = true) {
            var worksheet = WorksheetRoot;
            var columns = worksheet.GetFirstChild<Columns>();
            if (columns == null) {
                columns = worksheet.InsertAt(new Columns(), 0);
            }

            SetColumnWidthsCore(columns, columnIndexes, widths, bestFit);
        }

        private static void SetColumnWidthsCore(Columns columns, IReadOnlyList<int> columnIndexes, double[] widths, bool bestFit = true) {
            if (!columns.Elements<Column>().Any()) {
                bool appendedColumn = false;
                for (int i = 0; i < columnIndexes.Count; i++) {
                    double width = NormalizeColumnWidth(widths[i]);
                    if (width <= 0) {
                        continue;
                    }

                    columns.Append(new Column {
                        Min = (uint)columnIndexes[i],
                        Max = (uint)columnIndexes[i],
                        Width = width,
                        CustomWidth = true,
                        BestFit = bestFit ? true : (bool?)null
                    });
                    appendedColumn = true;
                }

                if (appendedColumn) {
                    ReorderColumns(columns);
                } else {
                    columns.Remove();
                }

                return;
            }

            // Auto-fit creates individual column definitions. Index that common
            // case once instead of scanning the growing collection per width.
            var individualColumns = new Dictionary<uint, Column>();
            bool allIndividual = true;
            foreach (Column column in columns.Elements<Column>()) {
                if (column.Min == null || column.Max == null || column.Min.Value != column.Max.Value
                    || individualColumns.ContainsKey(column.Min.Value)) {
                    allIndividual = false;
                    break;
                }
                individualColumns.Add(column.Min.Value, column);
            }

            for (int i = 0; i < columnIndexes.Count; i++) {
                if (allIndividual) {
                    uint index = (uint)columnIndexes[i];
                    individualColumns.TryGetValue(index, out Column? column);
                    SetColumnWidthValue(columns, index, widths[i], column, bestFit);
                } else {
                    SetColumnWidthCore(columns, columnIndexes[i], widths[i], bestFit);
                }
            }

            if (columns.Elements<Column>().Any()) {
                ReorderColumns(columns);
            } else {
                columns.Remove();
            }
        }


        private static void SetColumnWidthCore(Columns columns, int columnIndex, double width, bool bestFit = true) {
            Column? column = columns.Elements<Column>()
                .FirstOrDefault(c => c.Min != null && c.Max != null && c.Min.Value <= (uint)columnIndex && c.Max.Value >= (uint)columnIndex);

            if (column != null) {
                column = SplitColumn(columns, column, (uint)columnIndex);
            }

            SetColumnWidthValue(columns, (uint)columnIndex, width, column, bestFit);
        }

        private static void SetColumnWidthValue(Columns columns, uint columnIndex, double width, Column? column, bool bestFit = true) {
            width = NormalizeColumnWidth(width);

            if (width > 0) {
                if (column == null) {
                    column = new Column { Min = columnIndex, Max = columnIndex };
                    columns.Append(column);
                }
                column.Width = width;
                column.CustomWidth = true;
                column.BestFit = bestFit ? true : (bool?)null;
            } else if (column != null) {
                column.Remove();
            }
        }


        private static Column SplitColumn(Columns columns, Column column, uint index) {
            if (column.Min!.Value == index && column.Max!.Value == index) {
                return column;
            }

            uint min = column.Min!.Value;
            uint max = column.Max!.Value;
            var template = (Column)column.CloneNode(true);
            column.Remove();

            if (min < index) {
                var left = (Column)template.CloneNode(true);
                left.Min = min;
                left.Max = index - 1;
                columns.Append(left);
            }

            var middle = (Column)template.CloneNode(true);
            middle.Min = index;
            middle.Max = index;
            columns.Append(middle);

            if (index < max) {
                var right = (Column)template.CloneNode(true);
                right.Min = index + 1;
                right.Max = max;
                columns.Append(right);
            }

            return middle;
        }

        private static void ReorderColumns(Columns columns) {
            var ordered = columns.Elements<Column>().OrderBy(c => c.Min!.Value).ToList();
            columns.RemoveAllChildren<Column>();
            Column? previous = null;
            foreach (var col in ordered) {
                if (previous != null && col.Min!.Value <= previous.Max!.Value) {
                    if (col.Max!.Value <= previous.Max!.Value) {
                        continue;
                    }
                    col.Min = previous.Max!.Value + 1;
                }
                columns.Append(col);
                previous = col;
            }

        }


        /// <summary>
        /// Sets the width of the specified column.
        /// </summary>
        /// <param name="columnIndex">1-based column index.</param>
        /// <param name="width">The column width.</param>
        public void SetColumnWidth(int columnIndex, ExcelColumnWidth width) {
            double value = NormalizeColumnWidth(width.ToCharacters(this));
            if (_excelDocument.HasUnmaterializedDirectDataSetRows) {
                _excelDocument.MaterializeDeferredDataSetImport();
            }

            WriteLock(() => {
                var worksheet = WorksheetRoot;
                var columns = worksheet.GetFirstChild<Columns>();
                if (columns == null) {
                    columns = worksheet.InsertAt(new Columns(), 0);
                }
                var column = columns.Elements<Column>()
                    .FirstOrDefault(c => c.Min != null && c.Max != null && c.Min.Value <= (uint)columnIndex && c.Max.Value >= (uint)columnIndex);
                if (column == null) {
                    column = new Column { Min = (uint)columnIndex, Max = (uint)columnIndex };
                    columns.Append(column);
                } else {
                    column = SplitColumn(columns, column, (uint)columnIndex);
                }
                if (value > 0) {
                    column.Width = value;
                    column.CustomWidth = true;
                    column.BestFit = null;
                } else {
                    column.Width = null;
                    column.CustomWidth = false;
                    column.BestFit = null;
                }
                ReorderColumns(columns);
            });
        }

        /// <summary>
        /// Hides or shows the specified column.
        /// </summary>
        /// <param name="columnIndex">1-based column index.</param>
        /// <param name="hidden">True to hide the column; false to show it.</param>
        public void SetColumnHidden(int columnIndex, bool hidden) {
            _excelDocument.MaterializeDeferredDataSetImport();
            WriteLock(() => {
                var worksheet = WorksheetRoot;
                var columns = worksheet.GetFirstChild<Columns>();
                if (columns == null) {
                    columns = worksheet.InsertAt(new Columns(), 0);
                }
                var column = columns.Elements<Column>()
                    .FirstOrDefault(c => c.Min != null && c.Max != null && c.Min.Value <= (uint)columnIndex && c.Max.Value >= (uint)columnIndex);
                if (column == null) {
                    column = CreateDefaultWidthColumn((uint)columnIndex);
                    columns.Append(column);
                }
                column.Hidden = hidden ? true : (bool?)null;
            });
        }
    }
}
