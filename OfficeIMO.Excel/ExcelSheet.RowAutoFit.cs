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
        private double GetDefaultRowHeightPoints() {
            var sheetFormat = WorksheetRoot.GetFirstChild<SheetFormatProperties>();
            if (sheetFormat?.DefaultRowHeight != null && sheetFormat.DefaultRowHeight.Value > 0) {
                return sheetFormat.DefaultRowHeight.Value;
            }
            return 15.0; // Excel default for Calibri 11pt
        }

        private bool HasWrapText(Cell cell) {
            if (cell.StyleIndex == null) return false;

            var stylesPart = _excelDocument.WorkbookPartRoot?.WorkbookStylesPart;
            var stylesheet = stylesPart?.Stylesheet;
            var cellFormats = stylesheet?.CellFormats;

            if (cellFormats == null) return false;

            var cellFormat = cellFormats.Elements<CellFormat>().ElementAtOrDefault((int)cell.StyleIndex.Value);
            if (cellFormat?.Alignment == null) return false;

            return cellFormat.Alignment.WrapText?.Value == true;
        }

        private double CalculateRowHeight(int rowIndex) {
            var worksheet = WorksheetRoot;
            SheetData? sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData == null) return 0;

            Row? row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex != null && r.RowIndex.Value == (uint)rowIndex);
            if (row == null) return 0;

            double defaultHeight = GetDefaultRowHeightPoints();
            double maxHeight = defaultHeight; // Start with default as minimum
            bool hasContent = false;

            // Pre-calc default font metrics and MDW for pixel conversions.
            var textMeasurer = ExcelTextMeasurer.Create(GetWorkbookDefaultFontInfo());
            var defaultStyle = textMeasurer.CreateDefaultStyle(96);
            float mdw = defaultStyle.MaximumDigitWidth;
            if (mdw <= 0.0001f) return defaultHeight;

            // Helper to get available content width in pixels for a given cell's column span
            double GetAvailableWidthPx(Cell c) {
                const double pixelPadding = 2.0; // both sides added by Excel grid
                                                 // Determine merged span width
                string reference = c.CellReference?.Value ?? throw new InvalidOperationException("CellReference is null");
                (int fromCol, int toCol) = GetCellMergeSpan(c) ?? (GetColumnIndex(reference), GetColumnIndex(reference));
                double totalPx = 0;
                for (int col = fromCol; col <= toCol; col++) {
                    totalPx += GetColumnWidthPixels(col, mdw);
                }
                // subtract small inner padding for content
                double contentPx = Math.Max(0, totalPx - 2 * pixelPadding);
                return contentPx;
            }

            foreach (var cell in row.Elements<Cell>()) {
                string text = GetCellAutoFitText(cell);
                if (string.IsNullOrWhiteSpace(text)) continue;
                hasContent = true;

                var fontInfo = GetCellFontInfo(cell, textMeasurer.FallbackFontInfo);
                var style = textMeasurer.CreateStyle(fontInfo, 96);

                // Measure a consistent line height using representative glyphs, but never below default row height
                // ClosedXML effectively uses a line box slightly taller than raw metrics; add a small pixel fudge.
                float measuredPx = textMeasurer.MeasureHeightOrDefault("Xg", style, 0); // representative ascender/descender
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

                if (wrap) {
                    // Available width in pixels for this cell (span-aware)
                    double availPx = GetAvailableWidthPx(cell);
                    if (availPx > 0) {
                        int linesCount = 0;
                        foreach (var hard in hardLines) {
                            // At minimum, each hard segment is one line, even if empty
                            linesCount += CountWrappedLines(hard, availPx, textMeasurer, style);
                        }
                        // Ensure we never undercount hard breaks
                        totalLines = Math.Max(totalLines, linesCount);
                    }
                }

                // Excel behavior roughly aligns to: height = baseLineHeight * lines + small padding
                // Increase padding slightly for multi-line to avoid clipping
                double paddingPt = totalLines > 1 ? 2.5 : 0.0;
                double cellHeight = baseLineHeightPt * totalLines + paddingPt;
                if (totalLines > 1) {
                    cellHeight *= 1.20;
                }

                // Ensure Excel wraps when our calculation indicates multiple lines
                if (totalLines > 1 && !HasWrapText(cell)) {
                    ApplyWrapText(cell);
                }

                if (cellHeight > maxHeight) {
                    maxHeight = cellHeight;
                }
            }

            // Round to reasonable precision and return desired height
            return hasContent ? Math.Round(maxHeight, 2) : 0;
        }

        private int CountWrappedLines(string text, double maxWidthPx, ExcelTextMeasurer textMeasurer, ExcelTextMeasurer.Style style) {
            // Empty line still occupies one visual line
            if (string.IsNullOrEmpty(text)) return 1;

            // Quick accept if whole text fits
            float fullWidth = textMeasurer.MeasureWidthOrDefault(text, style, 0);
            if (fullWidth <= maxWidthPx) return 1;

            // Word-based greedy wrap
            var words = SplitIntoWords(text);
            int lines = 1;
            double current = 0;
            for (int i = 0; i < words.Count; i++) {
                string token = words[i];
                bool isSpace = token == " ";
                if (isSpace) {
                    // Defer space addition until next word to avoid trailing space width issues
                    continue;
                }

                string segment = token;
                float w = textMeasurer.MeasureWidthOrDefault(segment, style, 0);
                // If we had a previous nonempty segment on the line, consider a space before this word
                if (current > 0) {
                    float spaceW = textMeasurer.MeasureWidthOrDefault(" ", style, 0);
                    w += spaceW;
                }

                if (w > maxWidthPx) {
                    // Word itself too long: split by characters
                    var chars = token.ToCharArray();
                    var sb = new StringBuilder();
                    for (int c = 0; c < chars.Length; c++) {
                        string candidate = (current > 0 ? " " : string.Empty) + sb.ToString() + chars[c];
                        float cw = textMeasurer.MeasureWidthOrDefault(candidate, style, 0);
                        if (cw > maxWidthPx) {
                            // break before this char
                            lines++;
                            sb.Clear();
                            current = 0;
                            candidate = chars[c].ToString();
                            cw = textMeasurer.MeasureWidthOrDefault(candidate, style, 0);
                        }
                        sb.Append(chars[c]);
                        current = cw;
                    }
                    continue;
                }

                if (current + w > maxWidthPx + 0.1) {
                    // Move word to next line
                    lines++;
                    current = textMeasurer.MeasureWidthOrDefault(token, style, 0); // start with word only on new line
                } else {
                    current += w;
                }
            }

            return Math.Max(1, lines);
        }

        private List<string> SplitIntoWords(string text) {
            var list = new List<string>();
            int i = 0;
            while (i < text.Length) {
                if (char.IsWhiteSpace(text[i])) {
                    // normalize all whitespace to single space for width measuring consistency
                    list.Add(" ");
                    while (i < text.Length && char.IsWhiteSpace(text[i])) i++;
                } else {
                    int start = i;
                    while (i < text.Length && !char.IsWhiteSpace(text[i])) i++;
                    list.Add(text.Substring(start, i - start));
                }
            }
            return list;
        }

        private (int fromCol, int toCol)? GetCellMergeSpan(Cell cell) {
            var ws = WorksheetRoot;
            var merges = ws.Elements<MergeCells>().FirstOrDefault();
            if (merges == null) return null;
            var r = cell.CellReference?.Value;
            if (string.IsNullOrEmpty(r)) return null;
            int selfCol = GetColumnIndex(r!);
            int selfRow = GetRowIndex(r!);
            foreach (var mc in merges.Elements<MergeCell>()) {
                var refAttr = mc.Reference?.Value; // e.g. "A1:C1"
                if (string.IsNullOrEmpty(refAttr)) continue;
                var parts = refAttr!.Split(':');
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

        private double GetColumnWidthPixels(int columnIndex, float mdw) {
            // Find explicit column width if present; else use default width
            double width = GetColumnWidthUnits(columnIndex);
            // Convert Excel width to pixels using MDW
            double pixels = Math.Truncate((256.0 * width + Math.Truncate(128.0 / mdw)) / 256.0 * mdw);
            return pixels;
        }

        private double GetColumnWidthUnits(int columnIndex) {
            var ws = WorksheetRoot;
            var columns = ws.GetFirstChild<Columns>();
            var col = columns?.Elements<Column>()
                .FirstOrDefault(c => c.Min != null && c.Max != null && c.Min.Value <= (uint)columnIndex && c.Max.Value >= (uint)columnIndex);
            if (col?.Width != null && col.CustomWidth != null && col.CustomWidth.Value) {
                return col.Width.Value;
            }
            // Fallback to sheet default or Excel default 8.43
            var sf = ws.GetFirstChild<SheetFormatProperties>();
            if (sf?.DefaultColumnWidth != null && sf.DefaultColumnWidth.Value > 0)
                return sf.DefaultColumnWidth.Value;
            return 8.43; // Excel's default width for Calibri 11
        }

        private static double NormalizeRowHeight(double height) {
            if (double.IsNaN(height) || double.IsInfinity(height)) {
                return 0;
            }

            if (height <= 0) {
                return 0;
            }

            return Math.Min(height, 409D);
        }

        private void SetRowHeightCore(int rowIndex, double height, bool normalizeForExcelVisibleHeight = false) {
            var worksheet = WorksheetRoot;
            SheetData? sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData == null) return;
            Row? row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex != null && r.RowIndex.Value == (uint)rowIndex);
            if (row == null) return;

            height = NormalizeRowHeight(height);
            if (height > 0) {
                double storedHeight = normalizeForExcelVisibleHeight
                    ? height * 1.5
                    : height;
                row.Height = Math.Round(storedHeight, 2);
                row.CustomHeight = true;
            } else {
                row.Height = null;
                row.CustomHeight = null;
            }
        }

        private void UpdateSheetFormat() {
            var worksheet = WorksheetRoot;
            SheetData? sheetData = worksheet.GetFirstChild<SheetData>();
            var sheetFormat = worksheet.GetFirstChild<SheetFormatProperties>();

            bool anyCustom = sheetData?.Elements<Row>()
                .Any(r => r.CustomHeight != null && r.CustomHeight.Value) == true;

            if (anyCustom) {
                if (sheetFormat == null) {
                    sheetFormat = worksheet.InsertAt(new SheetFormatProperties(), 0);
                }
                if (sheetFormat.DefaultRowHeight == null || sheetFormat.DefaultRowHeight.Value <= 0) {
                    sheetFormat.DefaultRowHeight = 15D;
                }
                // Do not set CustomHeight here; it's for default height semantics, not per-row
            }
        }

        /// <summary>
        /// Automatically fits all rows based on their content.
        /// </summary>
        /// <param name="mode">Overrides how the auto-fit work is scheduled across rows.</param>
        /// <param name="ct">Cancels the row auto-fit pass while heights are being calculated or applied.</param>
        public void AutoFitRows(ExecutionMode? mode = null, CancellationToken ct = default) {
            _excelDocument.MaterializeDeferredDataSetImport();
            var worksheet = WorksheetRoot;
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
                sequentialCore: () => {
                    // Sequential path with NoLock
                    for (int i = 0; i < rowIndexes.Count; i++) {
                        computed[i] = CalculateRowHeight(rowIndexes[i]);
                    }

                    for (int i = 0; i < rowIndexes.Count; i++) {
                        // Excel normalizes OfficeIMO-authored auto-fit row heights down on open/save; serialize a
                        // pixel-equivalent height so the visible Excel row height matches the measured value.
                        SetRowHeightCore(rowIndexes[i], computed[i], normalizeForExcelVisibleHeight: true);
                    }

                    UpdateSheetFormat();
                    if (EffectiveExecution.SaveWorksheetAfterAutoFit) {
                        worksheet.Save();
                    }
                },
                computeParallel: () => {
                    // Parallel compute phase - calculate heights without DOM mutation
                    var failures = new System.Collections.Concurrent.ConcurrentBag<int>();
                    Parallel.For(0, rowIndexes.Count, new ParallelOptions {
                        CancellationToken = ct,
                        MaxDegreeOfParallelism = EffectiveExecution.MaxDegreeOfParallelism ?? -1
                    }, i => {
                        try {
                            computed[i] = CalculateRowHeight(rowIndexes[i]);
                        } catch {
                            failures.Add(i);
                        }
                    });
                    if (!failures.IsEmpty) {
                        foreach (var idx in failures) {
                            try { computed[idx] = CalculateRowHeight(rowIndexes[idx]); } catch { computed[idx] = 0; }
                        }
                    }
                },
                applySequential: () => {
                    // Apply phase - write all row heights to DOM
                    for (int i = 0; i < rowIndexes.Count; i++) {
                        // Excel normalizes OfficeIMO-authored auto-fit row heights down on open/save; serialize a
                        // pixel-equivalent height so the visible Excel row height matches the measured value.
                        SetRowHeightCore(rowIndexes[i], computed[i], normalizeForExcelVisibleHeight: true);
                    }
                    UpdateSheetFormat();
                    if (EffectiveExecution.SaveWorksheetAfterAutoFit) {
                        worksheet.Save();
                    }
                },
                ct: ct
            );
        }




        /// <summary>
        /// Sets whether the specified row is hidden.
        /// </summary>
        /// <param name="rowIndex">1-based row index.</param>
        /// <param name="hidden">True to hide the row; false to show it.</param>
        public void SetRowHidden(int rowIndex, bool hidden) {
            if (rowIndex <= 0) {
                return;
            }

            _excelDocument.MaterializeDeferredDataSetImport();
            WriteLock(() => {
                SheetData sheetData = GetOrCreateSheetData();
                Row row = GetOrCreateRowElement(sheetData, rowIndex);
                row.Hidden = hidden ? true : (bool?)null;
                WorksheetRoot.Save();
            });
        }

        /// <summary>
        /// Auto-fits the height of the specified row based on its contents.
        /// </summary>
        /// <param name="rowIndex">1-based row index.</param>
        public void AutoFitRow(int rowIndex) {
            WriteLockConditional(() => {
                var height = CalculateRowHeight(rowIndex);
                // Excel normalizes OfficeIMO-authored auto-fit row heights down on open/save; serialize a
                // pixel-equivalent height so the visible Excel row height matches the measured value.
                SetRowHeightCore(rowIndex, height, normalizeForExcelVisibleHeight: true);
                UpdateSheetFormat();
                if (EffectiveExecution.SaveWorksheetAfterAutoFit) {
                    WorksheetRoot.Save();
                }
            });
        }

        /// <summary>
        /// Sets the explicit height of the specified row in points. Use a non-positive height to clear the custom row height.
        /// </summary>
        /// <param name="rowIndex">1-based row index.</param>
        /// <param name="height">Row height in points.</param>
        public void SetRowHeight(int rowIndex, double height) {
            if (rowIndex <= 0) {
                return;
            }

            height = NormalizeRowHeight(height);
            _excelDocument.MaterializeDeferredDataSetImport();
            WriteLock(() => {
                SheetData sheetData = GetOrCreateSheetData();
                GetOrCreateRowElement(sheetData, rowIndex);
                SetRowHeightCore(rowIndex, height);
                UpdateSheetFormat();
                WorksheetRoot.Save();
            });
        }
    }
}
