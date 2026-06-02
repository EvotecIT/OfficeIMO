using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private double CalculateColumnWidth(int columnIndex) {
            return CalculateColumnWidths([columnIndex], CancellationToken.None)[0];
        }

        private const double MaxExcelColumnWidth = 255.0;

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

        private void SetColumnWidthsCore(IReadOnlyList<int> columnIndexes, double[] widths) {
            var worksheet = WorksheetRoot;
            var columns = worksheet.GetFirstChild<Columns>();
            if (columns == null) {
                columns = worksheet.InsertAt(new Columns(), 0);
            }

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
                        BestFit = true
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

            for (int i = 0; i < columnIndexes.Count; i++) {
                SetColumnWidthCore(columns, columnIndexes[i], widths[i]);
            }

            if (columns.Elements<Column>().Any()) {
                ReorderColumns(columns);
            } else {
                columns.Remove();
            }
        }


        private static void SetColumnWidthCore(Columns columns, int columnIndex, double width) {
            Column? column = columns.Elements<Column>()
                .FirstOrDefault(c => c.Min != null && c.Max != null && c.Min.Value <= (uint)columnIndex && c.Max.Value >= (uint)columnIndex);

            if (column != null) {
                column = SplitColumn(columns, column, (uint)columnIndex);
            }

            width = NormalizeColumnWidth(width);

            if (width > 0) {
                if (column == null) {
                    column = new Column { Min = (uint)columnIndex, Max = (uint)columnIndex };
                    columns.Append(column);
                }
                column.Width = width;
                column.CustomWidth = true;
                column.BestFit = true;
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
        public void SetColumnWidth(int columnIndex, double width) {
            width = NormalizeColumnWidth(width);
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
                    column = new Column { Min = (uint)columnIndex, Max = (uint)columnIndex };
                    columns.Append(column);
                }
                if (width > 0) {
                    column.Width = width;
                    column.CustomWidth = true;
                } else {
                    column.Width = null;
                    column.CustomWidth = false;
                    column.BestFit = null;
                }
                worksheet.Save();
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
                    column = new Column { Min = (uint)columnIndex, Max = (uint)columnIndex };
                    columns.Append(column);
                }
                column.Hidden = hidden ? true : (bool?)null;
                worksheet.Save();
            });
        }
    }
}
