using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using SixLabors.Fonts;
using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using SixLaborsColor = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Automatically fits all columns based on their content.
        /// </summary>
        public void AutoFitColumns(ExecutionMode? mode = null, CancellationToken ct = default) {
            var columnIndexes = GetAllColumnIndices();
            if (columnIndexes.Count == 0) return;

            var columnsList = columnIndexes.OrderBy(i => i).ToList();
            double[] computed = new double[columnsList.Count];

            ExecuteWithPolicy(
                opName: "AutoFitColumns",
                itemCount: columnsList.Count,
                overrideMode: mode,
                sequentialCore: () =>
                {
                    // Sequential path with NoLock
                    var worksheet = _worksheetPart.Worksheet;
                    SheetData sheetData = worksheet.GetFirstChild<SheetData>();
                    if (sheetData == null) return;

                    for (int i = 0; i < columnsList.Count; i++)
                    {
                        computed[i] = CalculateColumnWidth(columnsList[i]);
                    }
                    
                    for (int i = 0; i < columnsList.Count; i++)
                    {
                        SetColumnWidthCore(columnsList[i], computed[i]);
                    }
                    
                    worksheet.Save();
                },
                computeParallel: () =>
                {
                    // Parallel compute phase - calculate widths without DOM mutation
                    Parallel.For(0, columnsList.Count, new ParallelOptions 
                    {
                        CancellationToken = ct,
                        MaxDegreeOfParallelism = EffectiveExecution.MaxDegreeOfParallelism ?? -1
                    }, i =>
                    {
                        computed[i] = CalculateColumnWidth(columnsList[i]);
                    });
                },
                applySequential: () =>
                {
                    // Apply phase - write all column widths to DOM
                    var worksheet = _worksheetPart.Worksheet;
                    for (int i = 0; i < columnsList.Count; i++)
                    {
                        SetColumnWidthCore(columnsList[i], computed[i]);
                    }
                    worksheet.Save();
                },
                ct: ct
            );
        }

        private HashSet<int> GetAllColumnIndices()
        {
            var worksheet = _worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData == null) return new HashSet<int>();

            var columns = worksheet.GetFirstChild<Columns>();
            HashSet<int> columnIndexes = new HashSet<int>();

            foreach (var row in sheetData.Elements<Row>())
            {
                foreach (var cell in row.Elements<Cell>())
                {
                    if (cell.CellReference == null) continue;
                    columnIndexes.Add(GetColumnIndex(cell.CellReference.Value));
                }
            }

            if (columns != null)
            {
                foreach (var column in columns.Elements<Column>())
                {
                    uint min = column.Min?.Value ?? 0;
                    uint max = column.Max?.Value ?? 0;
                    for (uint i = min; i <= max; i++)
                    {
                        columnIndexes.Add((int)i);
                    }
                }
            }

            return columnIndexes;
        }

        private double CalculateColumnWidth(int columnIndex)
        {
            var worksheet = _worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData == null) return 0;

            double width = 0;

            foreach (var row in sheetData.Elements<Row>())
            {
                var cell = row.Elements<Cell>()
                    .FirstOrDefault(c => c.CellReference != null && GetColumnIndex(c.CellReference.Value) == columnIndex);
                if (cell == null) continue;
                string text = GetCellText(cell);
                if (string.IsNullOrWhiteSpace(text)) continue;
                var font = GetCellFont(cell);
                var options = new TextOptions(font);
                float zeroWidth = TextMeasurer.MeasureSize("0", options).Width;
                var size = TextMeasurer.MeasureSize(text, options);
                double cellWidth = zeroWidth == 0 ? 0 : size.Width / zeroWidth + 1;
                if (cellWidth > width) width = cellWidth;
            }

            return width;
        }

        private void SetColumnWidthCore(int columnIndex, double width)
        {
            var worksheet = _worksheetPart.Worksheet;
            var columns = worksheet.GetFirstChild<Columns>();
            if (columns == null)
            {
                columns = worksheet.InsertAt(new Columns(), 0);
            }

            Column column = columns.Elements<Column>()
                .FirstOrDefault(c => c.Min != null && c.Max != null && c.Min.Value <= (uint)columnIndex && c.Max.Value >= (uint)columnIndex);

            if (column != null)
            {
                column = SplitColumn(columns, column, (uint)columnIndex);
            }

            if (width > 0)
            {
                if (column == null)
                {
                    column = new Column { Min = (uint)columnIndex, Max = (uint)columnIndex };
                    columns.Append(column);
                }
                column.Width = width;
                column.CustomWidth = true;
                column.BestFit = true;
            }
            else if (column != null)
            {
                column.Remove();
            }

            ReorderColumns(columns);
        }

        /// <summary>
        /// Automatically fits all rows based on their content.
        /// </summary>
        public void AutoFitRows(ExecutionMode? mode = null, CancellationToken ct = default) {
            var worksheet = _worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData == null) return;

            var rowIndexes = sheetData.Elements<Row>()
                .Select(r => (int)r.RowIndex!.Value)
                .ToList();

            if (rowIndexes.Count == 0) return;

            double[] computed = new double[rowIndexes.Count];

            ExecuteWithPolicy(
                opName: "AutoFitRows",
                itemCount: rowIndexes.Count,
                overrideMode: mode,
                sequentialCore: () =>
                {
                    // Sequential path with NoLock
                    for (int i = 0; i < rowIndexes.Count; i++)
                    {
                        computed[i] = CalculateRowHeight(rowIndexes[i]);
                    }
                    
                    for (int i = 0; i < rowIndexes.Count; i++)
                    {
                        SetRowHeightCore(rowIndexes[i], computed[i]);
                    }
                    
                    UpdateSheetFormat();
                    worksheet.Save();
                },
                computeParallel: () =>
                {
                    // Parallel compute phase - calculate heights without DOM mutation
                    Parallel.For(0, rowIndexes.Count, new ParallelOptions 
                    {
                        CancellationToken = ct,
                        MaxDegreeOfParallelism = EffectiveExecution.MaxDegreeOfParallelism ?? -1
                    }, i =>
                    {
                        computed[i] = CalculateRowHeight(rowIndexes[i]);
                    });
                },
                applySequential: () =>
                {
                    // Apply phase - write all row heights to DOM
                    for (int i = 0; i < rowIndexes.Count; i++)
                    {
                        SetRowHeightCore(rowIndexes[i], computed[i]);
                    }
                    UpdateSheetFormat();
                    worksheet.Save();
                },
                ct: ct
            );
        }

        private void UpdateSheetFormat()
        {
            var worksheet = _worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            var sheetFormat = worksheet.GetFirstChild<SheetFormatProperties>();
            
            bool anyCustom = sheetData.Elements<Row>()
                .Any(r => r.CustomHeight != null && r.CustomHeight.Value);

            if (anyCustom)
            {
                if (sheetFormat == null)
                {
                    sheetFormat = worksheet.InsertAt(new SheetFormatProperties(), 0);
                }
                sheetFormat.DefaultRowHeight = 15;
                sheetFormat.CustomHeight = true;
            }
            else if (sheetFormat != null)
            {
                sheetFormat.Remove();
            }
        }

        private double CalculateRowHeight(int rowIndex)
        {
            var worksheet = _worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData == null) return 0;

            Row row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == (uint)rowIndex);
            if (row == null) return 0;

            const double defaultHeight = 15;
            const double pointsPerInch = 72.0;

            double maxHeight = 0;
            foreach (var cell in row.Elements<Cell>())
            {
                string text = GetCellText(cell);
                if (string.IsNullOrWhiteSpace(text)) continue;
                var font = GetCellFont(cell);
                var options = new TextOptions(font);
                var lines = text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
                double lineHeight = lines.Max(line => TextMeasurer.MeasureSize(line, options).Height * pointsPerInch / options.Dpi);
                double cellHeight = lineHeight * lines.Length;
                if (cellHeight > maxHeight) maxHeight = cellHeight;
            }

            return maxHeight > 0 ? maxHeight + 2 : 0;
        }

        private void SetRowHeightCore(int rowIndex, double height)
        {
            var worksheet = _worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData == null) return;

            Row row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == (uint)rowIndex);
            if (row == null) return;

            if (height > 0)
            {
                row.Height = height;
                row.CustomHeight = true;
            }
            else
            {
                row.Height = null;
                row.CustomHeight = null;
            }
        }

        public void AutoFitColumn(int columnIndex) {
            WriteLockConditional(() => {
                var width = CalculateColumnWidth(columnIndex);
                SetColumnWidthCore(columnIndex, width);
                _worksheetPart.Worksheet.Save();
            });
        }

        private static Column SplitColumn(Columns columns, Column column, uint index) {
            if (column.Min!.Value == index && column.Max!.Value == index) {
                return column;
            }

            uint min = column.Min.Value;
            uint max = column.Max.Value;
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

        public void SetColumnWidth(int columnIndex, double width) {
            WriteLock(() => {
                var worksheet = _worksheetPart.Worksheet;
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
                column.Width = width;
                column.CustomWidth = true;
                worksheet.Save();
            });
        }

        public void SetColumnHidden(int columnIndex, bool hidden) {
            WriteLock(() => {
                var worksheet = _worksheetPart.Worksheet;
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

        public void AutoFitRow(int rowIndex) {
            WriteLockConditional(() => {
                var height = CalculateRowHeight(rowIndex);
                SetRowHeightCore(rowIndex, height);
                UpdateSheetFormat();
                _worksheetPart.Worksheet.Save();
            });
        }

        /// <summary>
        /// Freezes panes on the worksheet.
        /// </summary>
        /// <param name="topRows">Number of rows at the top to freeze.</param>
        /// <param name="leftCols">Number of columns on the left to freeze.</param>
        public void Freeze(int topRows = 0, int leftCols = 0) {
            WriteLock(() => {
                Worksheet worksheet = _worksheetPart.Worksheet;
                SheetViews sheetViews = worksheet.GetFirstChild<SheetViews>();

                if (topRows == 0 && leftCols == 0) {
                    if (sheetViews != null) {
                        worksheet.RemoveChild(sheetViews);
                    }
                    worksheet.Save();
                    return;
                }

                if (sheetViews == null) {
                    sheetViews = new SheetViews();
                    
                    // Remove SheetData temporarily if it exists
                    var sheetData = worksheet.GetFirstChild<SheetData>();
                    if (sheetData != null) {
                        worksheet.RemoveChild(sheetData);
                    } else {
                        sheetData = new SheetData();
                    }
                    
                    // Add sheetViews first
                    worksheet.AppendChild(sheetViews);
                    
                    // Then add SheetData after sheetViews
                    worksheet.AppendChild(sheetData);
                }

                SheetView sheetView = sheetViews.GetFirstChild<SheetView>();
                if (sheetView == null) {
                    sheetView = new SheetView { WorkbookViewId = 0U };
                    sheetViews.Append(sheetView);
                }

                sheetView.RemoveAllChildren<Pane>();
                sheetView.RemoveAllChildren<Selection>();

                Pane pane = new Pane { State = PaneStateValues.Frozen };
                if (topRows > 0) {
                    pane.VerticalSplit = topRows;  // VerticalSplit = number of rows to freeze
                }
                if (leftCols > 0) {
                    pane.HorizontalSplit = leftCols;  // HorizontalSplit = number of columns to freeze
                }

                pane.TopLeftCell = GetColumnName(leftCols + 1) + (topRows + 1).ToString(CultureInfo.InvariantCulture);

                if (topRows > 0 && leftCols > 0) {
                    pane.ActivePane = PaneValues.BottomRight;
                    sheetView.Append(pane);
                    sheetView.Append(new Selection {
                        Pane = PaneValues.TopRight,
                        ActiveCell = pane.TopLeftCell,
                        SequenceOfReferences = new ListValue<StringValue> { InnerText = pane.TopLeftCell }
                    });
                    sheetView.Append(new Selection {
                        Pane = PaneValues.BottomLeft,
                        ActiveCell = pane.TopLeftCell,
                        SequenceOfReferences = new ListValue<StringValue> { InnerText = pane.TopLeftCell }
                    });
                    sheetView.Append(new Selection {
                        Pane = PaneValues.BottomRight,
                        ActiveCell = pane.TopLeftCell,
                        SequenceOfReferences = new ListValue<StringValue> { InnerText = pane.TopLeftCell }
                    });
                } else if (topRows > 0) {
                    pane.ActivePane = PaneValues.BottomLeft;
                    sheetView.Append(pane);
                    sheetView.Append(new Selection {
                        Pane = PaneValues.BottomLeft,
                        ActiveCell = pane.TopLeftCell,
                        SequenceOfReferences = new ListValue<StringValue> { InnerText = pane.TopLeftCell }
                    });
                } else {
                    pane.ActivePane = PaneValues.TopRight;
                    sheetView.Append(pane);
                    sheetView.Append(new Selection {
                        Pane = PaneValues.TopRight,
                        ActiveCell = pane.TopLeftCell,
                        SequenceOfReferences = new ListValue<StringValue> { InnerText = pane.TopLeftCell }
                    });
                }

                sheetView.Append(new Selection {
                    ActiveCell = "A1",
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1" }
                });

                worksheet.Save();
            });
        }


    }
}

