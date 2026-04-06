using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SixLabors.Fonts;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Automatically fits all columns based on their content.
        /// </summary>
        /// <param name="mode">Overrides how the auto-fit work is scheduled across columns.</param>
        /// <param name="ct">Cancels the auto-fit pass while widths are being measured or applied.</param>
        public void AutoFitColumns(ExecutionMode? mode = null, CancellationToken ct = default) {
            var columns = GetAllColumnIndices();
            if (columns.Count == 0) return;
            AutoFitColumnsInternal(columns.OrderBy(i => i).ToList(), mode, ct);
        }

        /// <summary>
        /// Automatically fits the supplied set of column indexes.
        /// </summary>
        /// <param name="columnIndexes">1-based column indexes that should be resized to fit their content.</param>
        /// <param name="mode">Overrides how the auto-fit work is scheduled across the selected columns.</param>
        /// <param name="ct">Cancels the auto-fit pass for the selected columns.</param>
        public void AutoFitColumnsFor(IEnumerable<int> columnIndexes, ExecutionMode? mode = null, CancellationToken ct = default) {
            if (columnIndexes == null) return;
            var list = columnIndexes.Where(i => i > 0).Distinct().OrderBy(i => i).ToList();
            if (list.Count == 0) return;
            AutoFitColumnsInternal(list, mode, ct);
        }

        /// <summary>
        /// Automatically fits all columns except the supplied indexes.
        /// </summary>
        /// <param name="columnsToSkip">1-based column indexes that should not be resized.</param>
        /// <param name="mode">Overrides how the auto-fit work is scheduled for the remaining columns.</param>
        /// <param name="ct">Cancels the auto-fit pass before it completes.</param>
        public void AutoFitColumnsExcept(IEnumerable<int> columnsToSkip, ExecutionMode? mode = null, CancellationToken ct = default) {
            var skip = new HashSet<int>(columnsToSkip ?? Array.Empty<int>());
            var remaining = GetAllColumnIndices().Where(i => i > 0 && !skip.Contains(i)).OrderBy(i => i).ToList();
            if (remaining.Count == 0) return;
            AutoFitColumnsInternal(remaining, mode, ct);
        }

        private void AutoFitColumnsInternal(IReadOnlyList<int> columnsList, ExecutionMode? mode, CancellationToken ct) {
            if (columnsList.Count == 0) return;
            double[] computed = new double[columnsList.Count];
            var measurementPlan = BuildAutoFitMeasurementPlan(columnsList, ct);
            int workload = Math.Max(columnsList.Count, measurementPlan.Measurements.Count);

            ExecuteWithPolicy(
                opName: "AutoFitColumns",
                itemCount: workload,
                overrideMode: mode,
                sequentialCore: () => {
                    var worksheet = WorksheetRoot;
                    var sheetData = worksheet.GetFirstChild<SheetData>();
                    if (sheetData == null) return;

                    computed = CalculateColumnWidths(measurementPlan, ct, parallel: false);

                    for (int i = 0; i < columnsList.Count; i++) {
                        SetColumnWidthCore(columnsList[i], computed[i]);
                    }

                    worksheet.Save();
                },
                computeParallel: () => {
                    computed = CalculateColumnWidths(measurementPlan, ct, parallel: true);
                },
                applySequential: () => {
                    var worksheet = WorksheetRoot;
                    for (int i = 0; i < columnsList.Count; i++) {
                        SetColumnWidthCore(columnsList[i], computed[i]);
                    }
                    worksheet.Save();
                },
                ct: ct
            );
        }

        private HashSet<int> GetAllColumnIndices() {
            var worksheet = WorksheetRoot;
            SheetData? sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData == null) return new HashSet<int>();

            var columns = worksheet.GetFirstChild<Columns>();
            HashSet<int> columnIndexes = new HashSet<int>();

            foreach (var row in sheetData.Elements<Row>()) {
                foreach (var cell in row.Elements<Cell>()) {
                    var cellRef = cell.CellReference?.Value;
                    if (string.IsNullOrEmpty(cellRef)) continue;
                    columnIndexes.Add(GetColumnIndex(cellRef!));
                }
            }

            if (columns != null) {
                foreach (var column in columns.Elements<Column>()) {
                    uint min = column.Min?.Value ?? 0;
                    uint max = column.Max?.Value ?? 0;
                    for (uint i = min; i <= max; i++) {
                        columnIndexes.Add((int)i);
                    }
                }
            }

            return columnIndexes;
        }

        private AutoFitMeasurementPlan BuildAutoFitMeasurementPlan(IReadOnlyList<int> columnsList, CancellationToken ct) {
            var worksheet = WorksheetRoot;
            SheetData? sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData == null || columnsList.Count == 0) {
                return new AutoFitMeasurementPlan(columnsList, new List<AutoFitMeasurement>());
            }

            var targetColumns = new Dictionary<int, int>(columnsList.Count);
            for (int i = 0; i < columnsList.Count; i++) {
                targetColumns[columnsList[i]] = i;
            }

            var measurements = new List<AutoFitMeasurement>();
            var uniqueMeasurements = new HashSet<AutoFitMeasurementKey>();
            var sharedStrings = _excelDocument.SharedStringTablePart?.SharedStringTable?.Elements<SharedStringItem>().ToList();
            var sharedStringTextCache = new Dictionary<int, string>();

            string GetCellTextFast(Cell cell) {
                if (cell.DataType?.Value == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString) {
                    var raw = cell.CellValue?.InnerText;
                    if (!string.IsNullOrEmpty(raw)
                        && int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out int id)
                        && sharedStrings != null
                        && id >= 0
                        && id < sharedStrings.Count) {
                        if (sharedStringTextCache.TryGetValue(id, out var cached) && cached != null) {
                            return cached;
                        }

                        string resolved = GetSharedStringText(sharedStrings[id]);
                        sharedStringTextCache[id] = resolved;
                        return resolved;
                    }

                    return string.Empty;
                }

                if (cell.DataType?.Value == DocumentFormat.OpenXml.Spreadsheet.CellValues.InlineString) {
                    return GetInlineStringText(cell.InlineString);
                }

                return cell.CellValue?.InnerText ?? string.Empty;
            }

            foreach (var row in sheetData.Elements<Row>()) {
                ct.ThrowIfCancellationRequested();

                foreach (var cell in row.Elements<Cell>()) {
                    string? reference = cell.CellReference?.Value;
                    if (string.IsNullOrEmpty(reference)) {
                        continue;
                    }

                    int columnIndex = GetColumnIndex(reference!);
                    if (!targetColumns.TryGetValue(columnIndex, out int targetIndex)) {
                        continue;
                    }

                    string text = GetCellTextFast(cell);
                    if (string.IsNullOrWhiteSpace(text)) {
                        continue;
                    }

                    uint styleIndex = cell.StyleIndex?.Value ?? 0U;
                    if (uniqueMeasurements.Add(new AutoFitMeasurementKey(targetIndex, styleIndex, text))) {
                        measurements.Add(new AutoFitMeasurement(targetIndex, styleIndex, text));
                    }
                }
            }

            return new AutoFitMeasurementPlan(columnsList, measurements);
        }

        private double[] CalculateColumnWidths(IReadOnlyList<int> columnsList, CancellationToken ct) {
            var plan = BuildAutoFitMeasurementPlan(columnsList, ct);
            return CalculateColumnWidths(plan, ct, parallel: false);
        }

        private double[] CalculateColumnWidths(AutoFitMeasurementPlan plan, CancellationToken ct, bool parallel) {
            double[] widths = new double[plan.Columns.Count];
            if (plan.Measurements.Count == 0 || plan.Columns.Count == 0) {
                return widths;
            }

            var defaultFont = GetDefaultFont();
            var defaultOptions = new TextOptions(defaultFont);
            float defaultMdw = MeasureWidthOrDefault("0", defaultOptions, fallback: 0);
            if (defaultMdw <= 0.0001f) {
                return widths;
            }

            var workbookPart = WorkbookPartRoot;
            var stylesheet = workbookPart?.WorkbookStylesPart?.Stylesheet;
            var cellFormats = stylesheet?.CellFormats?.Elements<CellFormat>().ToList();
            var fonts = stylesheet?.Fonts?.Elements<DocumentFormat.OpenXml.Spreadsheet.Font>().ToList();

            AutoFitStyleInfo ResolveStyleInfo(uint styleIndex, Dictionary<uint, AutoFitStyleInfo> styleCache) {
                if (styleCache.TryGetValue(styleIndex, out var cached)) {
                    return cached;
                }

                SixLabors.Fonts.Font font = defaultFont;
                if (cellFormats != null && fonts != null) {
                    var cellFormat = styleIndex < cellFormats.Count ? cellFormats[(int)styleIndex] : null;
                    if (cellFormat?.FontId != null) {
                        uint fontId = cellFormat.FontId.Value;
                        if (fontId < fonts.Count) {
                            font = CreateFontFromOpenXml(fonts[(int)fontId], defaultFont);
                        }
                    }
                }

                var options = new TextOptions(font);
                float mdw = MeasureWidthOrDefault("0", options, defaultMdw);
                var info = new AutoFitStyleInfo(font, options, mdw);
                styleCache[styleIndex] = info;
                return info;
            }

            float MeasureTextWidth(
                string text,
                uint styleIndex,
                AutoFitStyleInfo styleInfo,
                Dictionary<(uint styleIndex, string text), float> textWidthCache,
                Dictionary<uint, Dictionary<char, float>> charWidthCache) {
                if (textWidthCache.TryGetValue((styleIndex, text), out float cached)) {
                    return cached;
                }

                float measured;
                if (text.Contains('\n') || text.Contains('\r')) {
                    measured = 0;
                    string[] lines = text.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None);
                    foreach (string line in lines) {
                        if (string.IsNullOrEmpty(line)) {
                            continue;
                        }

                        float lineWidth = MeasureWidthOrDefault(line, styleInfo.Options, 0);
                        if (lineWidth > measured) {
                            measured = lineWidth;
                        }
                    }
                } else if (TryMeasureSimpleAutoFitTextWidth(text, styleIndex, styleInfo, charWidthCache, out float fastMeasured)) {
                    measured = fastMeasured;
                } else {
                    measured = MeasureWidthOrDefault(text, styleInfo.Options, 0);
                }

                textWidthCache[(styleIndex, text)] = measured;
                return measured;
            }

            const double pixelPadding = 2.0;

            void ApplyMeasurement(
                AutoFitMeasurement measurement,
                double[] localWidths,
                Dictionary<uint, AutoFitStyleInfo> styleCache,
                Dictionary<(uint styleIndex, string text), float> textWidthCache,
                Dictionary<uint, Dictionary<char, float>> charWidthCache) {
                    var styleInfo = ResolveStyleInfo(measurement.StyleIndex, styleCache);
                    float textWidthPx = MeasureTextWidth(measurement.Text, measurement.StyleIndex, styleInfo, textWidthCache, charWidthCache);
                    double cellWidthPx = textWidthPx + (2 * pixelPadding) + 1;
                    double columnWidth = Math.Truncate(cellWidthPx / styleInfo.MaximumDigitWidth * 256.0) / 256.0;

                    if (columnWidth > localWidths[measurement.TargetIndex]) {
                        localWidths[measurement.TargetIndex] = columnWidth;
                    }
                }

            if (!parallel || plan.Measurements.Count < 2) {
                var styleCache = new Dictionary<uint, AutoFitStyleInfo>();
                var textWidthCache = new Dictionary<(uint styleIndex, string text), float>();
                var charWidthCache = new Dictionary<uint, Dictionary<char, float>>();

                for (int i = 0; i < plan.Measurements.Count; i++) {
                    ct.ThrowIfCancellationRequested();
                    ApplyMeasurement(plan.Measurements[i], widths, styleCache, textWidthCache, charWidthCache);
                }

                return widths;
            }

            object mergeLock = new object();
            var options = new ParallelOptions {
                CancellationToken = ct,
                MaxDegreeOfParallelism = EffectiveExecution.MaxDegreeOfParallelism ?? -1
            };

            Parallel.ForEach(Partitioner.Create(0, plan.Measurements.Count), options,
                () => new AutoFitParallelState(plan.Columns.Count),
                (range, _, localState) => {
                    for (int i = range.Item1; i < range.Item2; i++) {
                        ct.ThrowIfCancellationRequested();
                        ApplyMeasurement(plan.Measurements[i], localState.Widths, localState.StyleCache, localState.TextWidthCache, localState.CharWidthCache);
                    }

                    return localState;
                },
                localState => {
                    lock (mergeLock) {
                        for (int i = 0; i < widths.Length; i++) {
                            if (localState.Widths[i] > widths[i]) {
                                widths[i] = localState.Widths[i];
                            }
                        }
                    }
                });

            return widths;
        }

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

            ReorderColumns(columns);
        }

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

            // Pre-calc default font metrics and MDW for pixel conversions
            var defaultFont = GetDefaultFont();
            var defaultOptions = new TextOptions(defaultFont) { Dpi = 96 };
            float mdw = TextMeasurer.MeasureSize("0", defaultOptions).Width;
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

                if (wrap) {
                    // Available width in pixels for this cell (span-aware)
                    double availPx = GetAvailableWidthPx(cell);
                    if (availPx > 0) {
                        int linesCount = 0;
                        foreach (var hard in hardLines) {
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
                if (totalLines > 1 && !HasWrapText(cell)) {
                    ApplyWrapText(cell);
                }

                if (cellHeight > maxHeight) {
                    maxHeight = cellHeight;
                }
            }

            // Round to reasonable precision and return desired height
            return Math.Round(maxHeight, 2);
        }

        private int CountWrappedLines(string text, double maxWidthPx, TextOptions options) {
            // Empty line still occupies one visual line
            if (string.IsNullOrEmpty(text)) return 1;

            // Quick accept if whole text fits
            float fullWidth = TextMeasurer.MeasureSize(text, options).Width;
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
                float w = TextMeasurer.MeasureSize(segment, options).Width;
                // If we had a previous nonempty segment on the line, consider a space before this word
                if (current > 0) {
                    float spaceW = TextMeasurer.MeasureSize(" ", options).Width;
                    w += spaceW;
                }

                if (w > maxWidthPx) {
                    // Word itself too long: split by characters
                    var chars = token.ToCharArray();
                    var sb = new StringBuilder();
                    for (int c = 0; c < chars.Length; c++) {
                        string candidate = (current > 0 ? " " : string.Empty) + sb.ToString() + chars[c];
                        float cw = TextMeasurer.MeasureSize(candidate, options).Width;
                        if (cw > maxWidthPx) {
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

                if (current + w > maxWidthPx + 0.1) {
                    // Move word to next line
                    lines++;
                    current = TextMeasurer.MeasureSize(token, options).Width; // start with word only on new line
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

        private void SetRowHeightCore(int rowIndex, double height) {
            var worksheet = WorksheetRoot;
            SheetData? sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData == null) return;
            Row? row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex != null && r.RowIndex.Value == (uint)rowIndex);
            if (row == null) return;

            double defaultHeight = GetDefaultRowHeightPoints();
            if (height > defaultHeight) {
                row.Height = height;
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
                        SetRowHeightCore(rowIndexes[i], computed[i]);
                    }

                    UpdateSheetFormat();
                    worksheet.Save();
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
                var width = CalculateColumnWidths([columnIndex], CancellationToken.None)[0];
                SetColumnWidthCore(columnIndex, width);
                WorksheetRoot.Save();
            });
        }

        private static float MeasureWidthOrDefault(string text, TextOptions options, float fallback) {
            try {
                float measured = TextMeasurer.MeasureSize(text, options).Width;
                return measured > 0.0001f ? measured : fallback;
            } catch {
                return fallback;
            }
        }

        private static bool TryMeasureSimpleAutoFitTextWidth(
            string text,
            uint styleIndex,
            AutoFitStyleInfo styleInfo,
            Dictionary<uint, Dictionary<char, float>> charWidthCache,
            out float measured) {
            measured = 0;
            if (string.IsNullOrEmpty(text)) {
                return false;
            }

            for (int i = 0; i < text.Length; i++) {
                if (!IsSimpleAutoFitCharacter(text[i])) {
                    return false;
                }
            }

            if (!charWidthCache.TryGetValue(styleIndex, out var perCharWidths)) {
                perCharWidths = new Dictionary<char, float>();
                charWidthCache[styleIndex] = perCharWidths;
            }

            float total = 0;
            for (int i = 0; i < text.Length; i++) {
                char current = text[i];
                if (!perCharWidths.TryGetValue(current, out float width)) {
                    width = MeasureWidthOrDefault(current.ToString(), styleInfo.Options, styleInfo.MaximumDigitWidth);
                    perCharWidths[current] = width;
                }

                total += width;
            }

            // Single-glyph summation can undercount string-level layout slightly on some fonts,
            // so bias upward by roughly one digit width to stay safely on the generous side.
            measured = total + styleInfo.MaximumDigitWidth;
            return true;
        }

        private static bool IsSimpleAutoFitCharacter(char value)
            => (value >= '0' && value <= '9')
            || value == '.'
            || value == ','
            || value == '-'
            || value == '+'
            || value == '/'
            || value == ':'
            || value == ' '
            || value == '%';

        private static string GetSharedStringText(SharedStringItem item) {
            if (item.Text != null) {
                return item.Text.Text ?? string.Empty;
            }

            var sb = new StringBuilder();
            foreach (var text in item.Descendants<Text>()) {
                sb.Append(text.Text);
            }

            return sb.ToString();
        }

        private static string GetInlineStringText(InlineString? inlineString) {
            if (inlineString == null) {
                return string.Empty;
            }

            if (inlineString.Text != null) {
                return inlineString.Text.Text ?? string.Empty;
            }

            var sb = new StringBuilder();
            foreach (var run in inlineString.Elements<Run>()) {
                if (run.Text != null) {
                    sb.Append(run.Text.Text);
                }
            }

            return sb.ToString();
        }

        private static SixLabors.Fonts.Font CreateFontFromOpenXml(DocumentFormat.OpenXml.Spreadsheet.Font fontElement, SixLabors.Fonts.Font fallbackFont) {
            var fontName = fontElement.GetFirstChild<FontName>()?.Val?.Value;
            var fontSize = fontElement.GetFirstChild<FontSize>()?.Val?.Value ?? fallbackFont.Size;
            bool bold = fontElement.GetFirstChild<Bold>() != null;
            bool italic = fontElement.GetFirstChild<Italic>() != null;

            try {
                var style = bold && italic ? FontStyle.BoldItalic : bold ? FontStyle.Bold : italic ? FontStyle.Italic : FontStyle.Regular;
                if (!string.IsNullOrEmpty(fontName)) {
                    return SystemFonts.CreateFont(fontName!, (float)fontSize, style);
                }

                return fallbackFont.Family.CreateFont((float)fontSize, style);
            } catch (FontFamilyNotFoundException) {
                var fallbackStyle = bold && italic ? FontStyle.BoldItalic : bold ? FontStyle.Bold : italic ? FontStyle.Italic : FontStyle.Regular;
                return fallbackFont.Family.CreateFont((float)fontSize, fallbackStyle);
            }
        }

        private readonly struct AutoFitStyleInfo {
            internal AutoFitStyleInfo(SixLabors.Fonts.Font font, TextOptions options, float maximumDigitWidth) {
                _font = font;
                _options = options;
                _maximumDigitWidth = maximumDigitWidth;
            }

            private readonly SixLabors.Fonts.Font _font;
            private readonly TextOptions _options;
            private readonly float _maximumDigitWidth;

            internal SixLabors.Fonts.Font Font => _font;
            internal TextOptions Options => _options;
            internal float MaximumDigitWidth => _maximumDigitWidth;
        }

        private readonly struct AutoFitMeasurement {
            internal AutoFitMeasurement(int targetIndex, uint styleIndex, string text) {
                _targetIndex = targetIndex;
                _styleIndex = styleIndex;
                _text = text;
            }

            private readonly int _targetIndex;
            private readonly uint _styleIndex;
            private readonly string _text;

            internal int TargetIndex => _targetIndex;
            internal uint StyleIndex => _styleIndex;
            internal string Text => _text;
        }

        private readonly struct AutoFitMeasurementKey : IEquatable<AutoFitMeasurementKey> {
            internal AutoFitMeasurementKey(int targetIndex, uint styleIndex, string text) {
                _targetIndex = targetIndex;
                _styleIndex = styleIndex;
                _text = text;
            }

            private readonly int _targetIndex;
            private readonly uint _styleIndex;
            private readonly string _text;

            public bool Equals(AutoFitMeasurementKey other)
                => _targetIndex == other._targetIndex
                && _styleIndex == other._styleIndex
                && string.Equals(_text, other._text, StringComparison.Ordinal);

            public override bool Equals(object? obj)
                => obj is AutoFitMeasurementKey other && Equals(other);

            public override int GetHashCode() {
                unchecked {
                    int hash = _targetIndex;
                    hash = (hash * 397) ^ (int)_styleIndex;
                    hash = (hash * 397) ^ StringComparer.Ordinal.GetHashCode(_text);
                    return hash;
                }
            }
        }

        private sealed class AutoFitMeasurementPlan {
            internal AutoFitMeasurementPlan(IReadOnlyList<int> columns, List<AutoFitMeasurement> measurements) {
                Columns = columns;
                Measurements = measurements;
            }

            internal IReadOnlyList<int> Columns { get; }
            internal List<AutoFitMeasurement> Measurements { get; }
        }

        private sealed class AutoFitParallelState {
            internal AutoFitParallelState(int columnCount) {
                Widths = new double[columnCount];
                StyleCache = new Dictionary<uint, AutoFitStyleInfo>();
                TextWidthCache = new Dictionary<(uint styleIndex, string text), float>();
                CharWidthCache = new Dictionary<uint, Dictionary<char, float>>();
            }

            internal double[] Widths { get; }
            internal Dictionary<uint, AutoFitStyleInfo> StyleCache { get; }
            internal Dictionary<(uint styleIndex, string text), float> TextWidthCache { get; }
            internal Dictionary<uint, Dictionary<char, float>> CharWidthCache { get; }
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

        /// <summary>
        /// Auto-fits the height of the specified row based on its contents.
        /// </summary>
        /// <param name="rowIndex">1-based row index.</param>
        public void AutoFitRow(int rowIndex) {
            WriteLockConditional(() => {
                var height = CalculateRowHeight(rowIndex);
                SetRowHeightCore(rowIndex, height);
                UpdateSheetFormat();
                WorksheetRoot.Save();
            });
        }

        /// <summary>
        /// Freezes panes on the worksheet.
        /// </summary>
        /// <param name="topRows">Number of rows at the top to freeze.</param>
        /// <param name="leftCols">Number of columns on the left to freeze.</param>
        public void Freeze(int topRows = 0, int leftCols = 0) {
            WriteLock(() => {
                Worksheet worksheet = WorksheetRoot;
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

        /// <summary>
        /// Shows or hides gridlines on the current sheet (view-level setting).
        /// </summary>
        public void SetGridlinesVisible(bool visible) {
            WriteLock(() => {
                var worksheet = WorksheetRoot;
                var sheetViews = worksheet.GetFirstChild<SheetViews>();
                if (sheetViews == null) {
                    sheetViews = new SheetViews();
                    worksheet.InsertAt(sheetViews, 0);
                }
                var view = sheetViews.GetFirstChild<SheetView>();
                if (view == null) {
                    view = new SheetView { WorkbookViewId = 0U };
                    sheetViews.Append(view);
                }
                view.ShowGridLines = visible;
                worksheet.Save();
            });
        }

        /// <summary>
        /// Configures basic print/page setup for the sheet.
        /// </summary>
        /// <param name="fitToWidth">Number of pages to fit horizontally (1 = fit to one page).</param>
        /// <param name="fitToHeight">Number of pages to fit vertically (0 = unlimited).</param>
        /// <param name="scale">Manual scale (10-400). Ignored if FitToWidth/Height are specified.</param>
        public void SetPageSetup(uint? fitToWidth = null, uint? fitToHeight = null, uint? scale = null) {
            WriteLock(() => {
                var ws = WorksheetRoot;
                var pageSetup = ws.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.PageSetup>();
                if (pageSetup == null) {
                    pageSetup = new DocumentFormat.OpenXml.Spreadsheet.PageSetup();
                    // Insert after PageMargins when present, else at end
                    var margins = ws.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.PageMargins>();
                    if (margins != null) ws.InsertAfter(pageSetup, margins); else ws.Append(pageSetup);
                }

                if (fitToWidth != null) pageSetup.FitToWidth = fitToWidth.Value;
                if (fitToHeight != null) pageSetup.FitToHeight = fitToHeight.Value;
                if (scale != null) pageSetup.Scale = scale.Value;

                ws.Save();
            });
        }


    }
}
