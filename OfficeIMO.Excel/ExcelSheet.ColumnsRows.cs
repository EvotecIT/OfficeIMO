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
                      SheetData? sheetData = worksheet.GetFirstChild<SheetData>();
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
                    var failures = new System.Collections.Concurrent.ConcurrentBag<int>();
                    Parallel.For(0, columnsList.Count, new ParallelOptions
                    {
                        CancellationToken = ct,
                        MaxDegreeOfParallelism = EffectiveExecution.MaxDegreeOfParallelism ?? -1
                    }, i =>
                    {
                        try
                        {
                            computed[i] = CalculateColumnWidth(columnsList[i]);
                        }
                        catch
                        {
                            failures.Add(i);
                        }
                    });
                    if (!failures.IsEmpty)
                    {
                        foreach (var idx in failures)
                        {
                            try { computed[idx] = CalculateColumnWidth(columnsList[idx]); }
                            catch { computed[idx] = 0; }
                        }
                    }
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
              SheetData? sheetData = worksheet.GetFirstChild<SheetData>();
              if (sheetData == null) return new HashSet<int>();

            var columns = worksheet.GetFirstChild<Columns>();
            HashSet<int> columnIndexes = new HashSet<int>();

              foreach (var row in sheetData.Elements<Row>())
              {
                  foreach (var cell in row.Elements<Cell>())
                  {
                      var cellRef = cell.CellReference?.Value;
                      if (string.IsNullOrEmpty(cellRef)) continue;
                      columnIndexes.Add(GetColumnIndex(cellRef));
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
              SheetData? sheetData = worksheet.GetFirstChild<SheetData>();
              if (sheetData == null) return 0;

            double maxWidth = 0;
            
            // Get the default font for MDW calculation
            var defaultFont = GetDefaultFont();
            var defaultOptions = new TextOptions(defaultFont) { Dpi = 96 };
            float mdw = TextMeasurer.MeasureSize("0", defaultOptions).Width;
            if (mdw <= 0.0001f) return 0;
            
            // Pixel Padding (PP) - extra pixels at left and right border (ClosedXML uses 2)
            const double pixelPadding = 2.0;

            foreach (var row in sheetData.Elements<Row>())
            {
                  var cell = row.Elements<Cell>()
                      .FirstOrDefault(c =>
                      {
                          string? reference = c.CellReference?.Value;
                          return reference != null && GetColumnIndex(reference) == columnIndex;
                      });
                if (cell == null) continue;
                
                string text = GetCellText(cell);
                if (string.IsNullOrWhiteSpace(text)) continue;
                
                // Check if cell has wrap text enabled
                bool hasWrapText = HasWrapText(cell);
                
                var font = GetCellFont(cell);
                var options = new TextOptions(font) { Dpi = 96 };

                float textWidthPx;
                if (text.Contains('\n') || text.Contains('\r'))
                {
                    // For text with newlines, measure the longest line (regardless of wrap setting)
                    string[] lines = text.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None);
                    textWidthPx = 0;
                    foreach (var line in lines)
                    {
                        if (!string.IsNullOrEmpty(line))
                        {
                            float lineWidth = TextMeasurer.MeasureSize(line, options).Width;
                            if (lineWidth > textWidthPx) textWidthPx = lineWidth;
                        }
                    }
                }
                else
                {
                    // Measure full text as single line
                    textWidthPx = TextMeasurer.MeasureSize(text, options).Width;
                }
                
                // ClosedXML formula: Add 2 * padding + 1 pixel for cell border
                double cellWidthPx = textWidthPx + (2 * pixelPadding) + 1;
                
                // Convert pixels to Excel column width using ClosedXML's formula
                // width = Truncate(pixels / MDW * 256) / 256
                double columnWidth = Math.Truncate(cellWidthPx / mdw * 256.0) / 256.0;
                
                if (columnWidth > maxWidth)
                {
                    maxWidth = columnWidth;
                }
            }

            return maxWidth;
        }

        private void SetColumnWidthCore(int columnIndex, double width)
        {
            var worksheet = _worksheetPart.Worksheet;
            var columns = worksheet.GetFirstChild<Columns>();
            if (columns == null)
            {
                columns = worksheet.InsertAt(new Columns(), 0);
            }

              Column? column = columns.Elements<Column>()
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

        private double GetDefaultRowHeightPoints()
        {
            var sheetFormat = _worksheetPart.Worksheet.GetFirstChild<SheetFormatProperties>();
            if (sheetFormat?.DefaultRowHeight != null && sheetFormat.DefaultRowHeight.Value > 0)
            {
                return sheetFormat.DefaultRowHeight.Value;
            }
            return 15.0; // Excel default for Calibri 11pt
        }

        private bool HasWrapText(Cell cell)
        {
            if (cell.StyleIndex == null) return false;
            
            var stylesPart = _excelDocument._spreadSheetDocument.WorkbookPart?.WorkbookStylesPart;
            var stylesheet = stylesPart?.Stylesheet;
            var cellFormats = stylesheet?.CellFormats;
            
            if (cellFormats == null) return false;
            
            var cellFormat = cellFormats.Elements<CellFormat>().ElementAtOrDefault((int)cell.StyleIndex.Value);
            if (cellFormat?.Alignment == null) return false;
            
            return cellFormat.Alignment.WrapText?.Value == true;
        }

        private double CalculateRowHeight(int rowIndex)
        {
            var worksheet = _worksheetPart.Worksheet;
              SheetData? sheetData = worksheet.GetFirstChild<SheetData>();
              if (sheetData == null) return 0;

              Row? row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex != null && r.RowIndex.Value == (uint)rowIndex);
              if (row == null) return 0;

            double defaultHeight = GetDefaultRowHeightPoints();
            double maxHeight = defaultHeight; // Start with default as minimum
            
            // Pre-calc default font metrics and MDW for pixel conversions
            var defaultFont = GetDefaultFont();
            var defaultOptions = new TextOptions(defaultFont) { Dpi = 96 };
            float mdw = TextMeasurer.MeasureSize("0", defaultOptions).Width;
            if (mdw <= 0.0001f) return defaultHeight;

            // Helper to get available content width in pixels for a given cell's column span
            double GetAvailableWidthPx(Cell c)
            {
                const double pixelPadding = 2.0; // both sides added by Excel grid
                // Determine merged span width
                  string reference = c.CellReference?.Value ?? throw new InvalidOperationException("CellReference is null");
                  (int fromCol, int toCol) = GetCellMergeSpan(c) ?? (GetColumnIndex(reference), GetColumnIndex(reference));
                double totalPx = 0;
                for (int col = fromCol; col <= toCol; col++)
                {
                    totalPx += GetColumnWidthPixels(col, mdw);
                }
                // subtract small inner padding for content
                double contentPx = Math.Max(0, totalPx - 2 * pixelPadding);
                return contentPx;
            }

            foreach (var cell in row.Elements<Cell>())
            {
                string text = GetCellText(cell);
                if (string.IsNullOrEmpty(text)) continue;
                
                var font = GetCellFont(cell);
                var options = new TextOptions(font) { Dpi = 96 };
                
                // Measure a consistent line height using representative glyphs, but never below default row height
                // ClosedXML effectively uses a line box that’s slightly taller than raw metrics; add a small pixel fudge
                float measuredPx = TextMeasurer.MeasureSize("Xg", options).Height; // representative ascender/descender
                double lineHeightPx = Math.Ceiling(measuredPx + 2); // add 2px to align with Excel/ClosedXML appearance
                double baseLineHeightPt = Math.Max(defaultHeight, lineHeightPx * 72.0 / 96.0);

                // Determine line count considering explicit newlines and wrapping
                // Compute line count: hard breaks always add lines; wrapping adds more within each segment
                var hardLines = text.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None);
                int totalLines = Math.Max(1, hardLines.Length);
                bool hasExplicitBreaks = hardLines.Length > 1;
                bool wrap = HasWrapText(cell) || hasExplicitBreaks; // Excel treats explicit breaks as wrapped content

                // Ensure WrapText for visual parity when explicit breaks exist
                if (hasExplicitBreaks && !HasWrapText(cell)) ApplyWrapText(cell);

                if (wrap)
                {
                    // Available width in pixels for this cell (span-aware)
                    double availPx = GetAvailableWidthPx(cell);
                    if (availPx > 0)
                    {
                        int linesCount = 0;
                        foreach (var hard in hardLines)
                        {
                            // At minimum, each hard segment is one line, even if empty
                            linesCount += CountWrappedLines(hard, availPx, options);
                        }
                        // Ensure we never undercount hard breaks
                        totalLines = Math.Max(totalLines, linesCount);
                    }
                }

                // Excel behavior roughly aligns to: height = baseLineHeight * lines + small padding
                // Increase padding slightly for multi-line to avoid clipping
                double paddingPt = totalLines > 1 ? 2.5 : 0.0;
                double cellHeight = baseLineHeightPt * totalLines + paddingPt;

                // Ensure Excel wraps when our calculation indicates multiple lines
                if (totalLines > 1 && !HasWrapText(cell))
                {
                    ApplyWrapText(cell);
                }

                if (cellHeight > maxHeight)
                {
                    maxHeight = cellHeight;
                }
            }

            // Round to reasonable precision and return desired height
            return Math.Round(maxHeight, 2);
        }

        private int CountWrappedLines(string text, double maxWidthPx, TextOptions options)
        {
            // Empty line still occupies one visual line
            if (string.IsNullOrEmpty(text)) return 1;

            // Quick accept if whole text fits
            float fullWidth = TextMeasurer.MeasureSize(text, options).Width;
            if (fullWidth <= maxWidthPx) return 1;

            // Word-based greedy wrap
            var words = SplitIntoWords(text);
            int lines = 1;
            double current = 0;
            for (int i = 0; i < words.Count; i++)
            {
                string token = words[i];
                bool isSpace = token == " ";
                if (isSpace)
                {
                    // Defer space addition until next word to avoid trailing space width issues
                    continue;
                }

                string segment = token;
                float w = TextMeasurer.MeasureSize(segment, options).Width;
                // If we had a previous nonempty segment on the line, consider a space before this word
                if (current > 0)
                {
                    float spaceW = TextMeasurer.MeasureSize(" ", options).Width;
                    w += spaceW;
                }

                if (w > maxWidthPx)
                {
                    // Word itself too long: split by characters
                    var chars = token.ToCharArray();
                    var sb = new StringBuilder();
                    for (int c = 0; c < chars.Length; c++)
                    {
                        string candidate = (current > 0 ? " " : string.Empty) + sb.ToString() + chars[c];
                        float cw = TextMeasurer.MeasureSize(candidate, options).Width;
                        if (cw > maxWidthPx)
                        {
                            // break before this char
                            lines++;
                            sb.Clear();
                            current = 0;
                            candidate = chars[c].ToString();
                            cw = TextMeasurer.MeasureSize(candidate, options).Width;
                        }
                        sb.Append(chars[c]);
                        current = cw;
                    }
                    continue;
                }

                if (current + w > maxWidthPx + 0.1)
                {
                    // Move word to next line
                    lines++;
                    current = TextMeasurer.MeasureSize(token, options).Width; // start with word only on new line
                }
                else
                {
                    current += w;
                }
            }

            return Math.Max(1, lines);
        }

        private List<string> SplitIntoWords(string text)
        {
            var list = new List<string>();
            int i = 0;
            while (i < text.Length)
            {
                if (char.IsWhiteSpace(text[i]))
                {
                    // normalize all whitespace to single space for width measuring consistency
                    list.Add(" ");
                    while (i < text.Length && char.IsWhiteSpace(text[i])) i++;
                }
                else
                {
                    int start = i;
                    while (i < text.Length && !char.IsWhiteSpace(text[i])) i++;
                    list.Add(text.Substring(start, i - start));
                }
            }
            return list;
        }

        private (int fromCol, int toCol)? GetCellMergeSpan(Cell cell)
        {
            var ws = _worksheetPart.Worksheet;
            var merges = ws.Elements<MergeCells>().FirstOrDefault();
            if (merges == null) return null;
            var r = cell.CellReference?.Value;
            if (string.IsNullOrEmpty(r)) return null;
            int selfCol = GetColumnIndex(r);
            int selfRow = GetRowIndex(r);
            foreach (var mc in merges.Elements<MergeCell>())
            {
                var refAttr = mc.Reference?.Value; // e.g. "A1:C1"
                if (string.IsNullOrEmpty(refAttr)) continue;
                var parts = refAttr.Split(':');
                if (parts.Length != 2) continue;
                int fromRow = GetRowIndex(parts[0]);
                int toRow = GetRowIndex(parts[1]);
                if (selfRow < fromRow || selfRow > toRow) continue;
                int fromCol = GetColumnIndex(parts[0]);
                int toCol = GetColumnIndex(parts[1]);
                if (selfCol < fromCol || selfCol > toCol) continue;
                return (fromCol, toCol);
            }
            return null;
        }

        private double GetColumnWidthPixels(int columnIndex, float mdw)
        {
            // Find explicit column width if present; else use default width
            double width = GetColumnWidthUnits(columnIndex);
            // Convert Excel width to pixels using MDW
            double pixels = Math.Truncate((256.0 * width + Math.Truncate(128.0 / mdw)) / 256.0 * mdw);
            return pixels;
        }

        private double GetColumnWidthUnits(int columnIndex)
        {
            var ws = _worksheetPart.Worksheet;
            var columns = ws.GetFirstChild<Columns>();
            var col = columns?.Elements<Column>()
                .FirstOrDefault(c => c.Min != null && c.Max != null && c.Min.Value <= (uint)columnIndex && c.Max.Value >= (uint)columnIndex);
            if (col?.Width != null && col.CustomWidth != null && col.CustomWidth.Value)
            {
                return col.Width.Value;
            }
            // Fallback to sheet default or Excel default 8.43
            var sf = ws.GetFirstChild<SheetFormatProperties>();
            if (sf?.DefaultColumnWidth != null && sf.DefaultColumnWidth.Value > 0)
                return sf.DefaultColumnWidth.Value;
            return 8.43; // Excel's default width for Calibri 11
        }

        private void SetRowHeightCore(int rowIndex, double height)
        {
            var worksheet = _worksheetPart.Worksheet;
              SheetData? sheetData = worksheet.GetFirstChild<SheetData>();
              if (sheetData == null) return;
              Row? row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex != null && r.RowIndex.Value == (uint)rowIndex);
              if (row == null) return;

            double defaultHeight = GetDefaultRowHeightPoints();
            if (height > defaultHeight)
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

        private void UpdateSheetFormat()
        {
            var worksheet = _worksheetPart.Worksheet;
              SheetData? sheetData = worksheet.GetFirstChild<SheetData>();
            var sheetFormat = worksheet.GetFirstChild<SheetFormatProperties>();

            bool anyCustom = sheetData?.Elements<Row>()
                .Any(r => r.CustomHeight != null && r.CustomHeight.Value) == true;

            if (anyCustom)
            {
                if (sheetFormat == null)
                {
                    sheetFormat = worksheet.InsertAt(new SheetFormatProperties(), 0);
                }
                if (sheetFormat.DefaultRowHeight == null || sheetFormat.DefaultRowHeight.Value <= 0)
                {
                    sheetFormat.DefaultRowHeight = 15D;
                }
                // Do not set CustomHeight here; it's for default height semantics, not per-row
            }
        }

        /// <summary>
        /// Automatically fits all rows based on their content.
        /// </summary>
        public void AutoFitRows(ExecutionMode? mode = null, CancellationToken ct = default) {
            var worksheet = _worksheetPart.Worksheet;
              SheetData? sheetData = worksheet.GetFirstChild<SheetData>();
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
                    var failures = new System.Collections.Concurrent.ConcurrentBag<int>();
                    Parallel.For(0, rowIndexes.Count, new ParallelOptions
                    {
                        CancellationToken = ct,
                        MaxDegreeOfParallelism = EffectiveExecution.MaxDegreeOfParallelism ?? -1
                    }, i =>
                    {
                        try
                        {
                            computed[i] = CalculateRowHeight(rowIndexes[i]);
                        }
                        catch
                        {
                            failures.Add(i);
                        }
                    });
                    if (!failures.IsEmpty)
                    {
                        foreach (var idx in failures)
                        {
                            try { computed[idx] = CalculateRowHeight(rowIndexes[idx]); }
                            catch { computed[idx] = 0; }
                        }
                    }
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

        

        /// <summary>
        /// Auto-fits the width of the specified column based on its contents.
        /// </summary>
        /// <param name="columnIndex">1-based column index.</param>
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

        /// <summary>
        /// Hides or shows the specified column.
        /// </summary>
        /// <param name="columnIndex">1-based column index.</param>
        /// <param name="hidden">True to hide the column; false to show it.</param>
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

        /// <summary>
        /// Auto-fits the height of the specified row based on its contents.
        /// </summary>
        /// <param name="rowIndex">1-based row index.</param>
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
              SheetViews? sheetViews = worksheet.GetFirstChild<SheetViews>();

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

                  SheetView? sheetView = sheetViews.GetFirstChild<SheetView>();
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

