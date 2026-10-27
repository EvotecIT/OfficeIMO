using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
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
        /// <summary>
        /// Automatically fits all columns based on their content.
        /// </summary>
        /// <param name="mode">Overrides how the auto-fit work is scheduled across columns.</param>
        /// <param name="ct">Cancels the auto-fit pass while widths are being measured or applied.</param>
        public void AutoFitColumns(ExcelExecutionMode? mode = null, CancellationToken ct = default) {
            ct.ThrowIfCancellationRequested();
            if (_excelDocument.TryEnableDirectTabularSaveCandidateAutoFit(this, ct)) {
                return;
            }

            _excelDocument.MaterializeDeferredDataSetImport(ct);
            if (CanSkipStableAutoFitColumns(null)) {
                ct.ThrowIfCancellationRequested();
                return;
            }

            var planWatch = EffectiveExecution.OnTiming == null ? null : System.Diagnostics.Stopwatch.StartNew();
            var measurementPlan = BuildAutoFitMeasurementPlanForAllColumns(ct);
            if (planWatch != null) {
                planWatch.Stop();
                EffectiveExecution.ReportTiming("AutoFitColumns.BuildPlan", planWatch.Elapsed);
            }
            ct.ThrowIfCancellationRequested();
            if (measurementPlan.Columns.Count == 0) return;
            AutoFitColumnsInternal(measurementPlan, mode, ct);
            RefreshTablePrintDefaults();
        }

        /// <summary>
        /// Automatically fits the supplied set of column indexes.
        /// </summary>
        /// <param name="columnIndexes">1-based column indexes that should be resized to fit their content.</param>
        /// <param name="mode">Overrides how the auto-fit work is scheduled across the selected columns.</param>
        /// <param name="ct">Cancels the auto-fit pass for the selected columns.</param>
        public void AutoFitColumnsFor(IEnumerable<int> columnIndexes, ExcelExecutionMode? mode = null, CancellationToken ct = default) {
            ct.ThrowIfCancellationRequested();
            if (columnIndexes == null) return;
            var list = columnIndexes.Where(i => i > 0).Distinct().OrderBy(i => i).ToList();
            ct.ThrowIfCancellationRequested();
            if (list.Count == 0) return;
            if (_excelDocument.TryEnableDirectTabularSaveCandidateAutoFit(this, list, ct)) {
                return;
            }

            _excelDocument.MaterializeDeferredDataSetImport(ct);
            if (CanSkipStableAutoFitColumns(list)) {
                ct.ThrowIfCancellationRequested();
                return;
            }

            AutoFitColumnsInternal(list, mode, ct);
        }

        private void AutoFitContiguousColumns(int startColumn, int columnCount, ExcelExecutionMode? mode = null, CancellationToken ct = default) {
            ct.ThrowIfCancellationRequested();
            if (startColumn <= 0 || columnCount <= 0) return;
            if (startColumn == 1) {
                int[] directCandidateColumns = new int[columnCount];
                for (int i = 0; i < directCandidateColumns.Length; i++) {
                    directCandidateColumns[i] = i + 1;
                }

                if (_excelDocument.TryEnableDirectTabularSaveCandidateAutoFit(this, directCandidateColumns, ct)) {
                    return;
                }
            }

            _excelDocument.MaterializeDeferredDataSetImport(ct);

            var columns = new int[columnCount];
            for (int i = 0; i < columns.Length; i++) {
                columns[i] = startColumn + i;
            }

            if (CanSkipStableAutoFitColumns(columns)) {
                ct.ThrowIfCancellationRequested();
                return;
            }

            AutoFitColumnsInternal(columns, mode, ct);
        }

        /// <summary>
        /// Automatically fits all columns except the supplied indexes.
        /// </summary>
        /// <param name="columnsToSkip">1-based column indexes that should not be resized.</param>
        /// <param name="mode">Overrides how the auto-fit work is scheduled for the remaining columns.</param>
        /// <param name="ct">Cancels the auto-fit pass before it completes.</param>
        public void AutoFitColumnsExcept(IEnumerable<int> columnsToSkip, ExcelExecutionMode? mode = null, CancellationToken ct = default) {
            ct.ThrowIfCancellationRequested();
            var skip = new HashSet<int>(columnsToSkip ?? Array.Empty<int>());
            ct.ThrowIfCancellationRequested();
            if (_excelDocument.TryGetDirectTabularSaveCandidateColumnCount(this, out int directColumnCount)) {
                var directRemaining = Enumerable.Range(1, directColumnCount).Where(i => !skip.Contains(i)).ToList();
                ct.ThrowIfCancellationRequested();
                if (directRemaining.Count == 0 || _excelDocument.TryEnableDirectTabularSaveCandidateAutoFit(this, directRemaining, ct)) {
                    return;
                }
            }

            _excelDocument.MaterializeDeferredDataSetImport(ct);
            var remaining = GetAllColumnIndices().Where(i => i > 0 && !skip.Contains(i)).OrderBy(i => i).ToList();
            ct.ThrowIfCancellationRequested();
            if (remaining.Count == 0) return;
            if (CanSkipStableAutoFitColumns(remaining)) {
                ct.ThrowIfCancellationRequested();
                return;
            }

            AutoFitColumnsInternal(remaining, mode, ct);
        }


        private void AutoFitColumnsInternal(IReadOnlyList<int> columnsList, ExcelExecutionMode? mode, CancellationToken ct) {
            ct.ThrowIfCancellationRequested();
            if (columnsList.Count == 0) return;
            var planWatch = EffectiveExecution.OnTiming == null ? null : System.Diagnostics.Stopwatch.StartNew();
            var measurementPlan = BuildAutoFitMeasurementPlan(columnsList, ct);
            if (planWatch != null) {
                planWatch.Stop();
                EffectiveExecution.ReportTiming("AutoFitColumns.BuildPlan", planWatch.Elapsed);
            }
            AutoFitColumnsInternal(measurementPlan, mode, ct);
        }

        private void AutoFitColumnsInternal(AutoFitMeasurementPlan measurementPlan, ExcelExecutionMode? mode, CancellationToken ct) {
            ct.ThrowIfCancellationRequested();
            var columnsList = measurementPlan.Columns;
            if (columnsList.Count == 0) return;
            double[] computed = new double[columnsList.Count];
            int workload = Math.Max(columnsList.Count, measurementPlan.Measurements.Count);

            ExecuteWithPolicy(
                opName: "AutoFitColumns",
                itemCount: workload,
                overrideMode: mode,
                sequentialCore: () => {
                    var worksheet = WorksheetRoot;
                    var sheetData = worksheet.GetFirstChild<SheetData>();
                    if (sheetData == null) return;

                    var calculateWatch = EffectiveExecution.OnTiming == null ? null : System.Diagnostics.Stopwatch.StartNew();
                    computed = CalculateColumnWidths(measurementPlan, ct, parallel: false);
                    if (calculateWatch != null) {
                        calculateWatch.Stop();
                        EffectiveExecution.ReportTiming("AutoFitColumns.CalculateWidths", calculateWatch.Elapsed);
                    }

                    ct.ThrowIfCancellationRequested();
                    var applyWatch = EffectiveExecution.OnTiming == null ? null : System.Diagnostics.Stopwatch.StartNew();
                    SetColumnWidthsCore(columnsList, computed);

                    if (EffectiveExecution.SaveWorksheetAfterAutoFit) {
                        worksheet.Save();
                    }
                    MarkRequiresSavePreparation();
                    _hasWorksheetMutations = false;
                    if (applyWatch != null) {
                        applyWatch.Stop();
                        EffectiveExecution.ReportTiming("AutoFitColumns.ApplyWidths", applyWatch.Elapsed);
                    }
                },
                computeParallel: () => {
                    var calculateWatch = EffectiveExecution.OnTiming == null ? null : System.Diagnostics.Stopwatch.StartNew();
                    computed = CalculateColumnWidths(measurementPlan, ct, parallel: true);
                    if (calculateWatch != null) {
                        calculateWatch.Stop();
                        EffectiveExecution.ReportTiming("AutoFitColumns.CalculateWidths", calculateWatch.Elapsed);
                    }
                },
                applySequential: () => {
                    ct.ThrowIfCancellationRequested();
                    var worksheet = WorksheetRoot;
                    var applyWatch = EffectiveExecution.OnTiming == null ? null : System.Diagnostics.Stopwatch.StartNew();
                    SetColumnWidthsCore(columnsList, computed);
                    if (EffectiveExecution.SaveWorksheetAfterAutoFit) {
                        worksheet.Save();
                    }
                    MarkRequiresSavePreparation();
                    _hasWorksheetMutations = false;
                    if (applyWatch != null) {
                        applyWatch.Stop();
                        EffectiveExecution.ReportTiming("AutoFitColumns.ApplyWidths", applyWatch.Elapsed);
                    }
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


        private double[] CalculateColumnWidths(IReadOnlyList<int> columnsList, CancellationToken ct) {
            var plan = BuildAutoFitMeasurementPlan(columnsList, ct);
            return CalculateColumnWidths(plan, ct, parallel: false);
        }

        private double[] CalculateColumnWidths(AutoFitMeasurementPlan plan, CancellationToken ct, bool parallel) {
            double[] widths = new double[plan.Columns.Count];
            if (plan.Measurements.Count == 0 || plan.Columns.Count == 0) {
                return widths;
            }

            var textMeasurer = ExcelTextMeasurer.Create(GetWorkbookDefaultFontInfo());
            float defaultMdw = textMeasurer.DefaultStyle.MaximumDigitWidth;
            if (defaultMdw <= 0.0001f) {
                return widths;
            }

            var workbookPart = WorkbookPartRoot;
            var stylesheet = workbookPart?.WorkbookStylesPart?.Stylesheet;
            var cellFormats = stylesheet?.CellFormats?.Elements<CellFormat>().ToList();
            var fonts = stylesheet?.Fonts?.Elements<DocumentFormat.OpenXml.Spreadsheet.Font>().ToList();

            ExcelTextMeasurer.Style ResolveStyleInfo(uint styleIndex, Dictionary<uint, ExcelTextMeasurer.Style> styleCache) {
                if (styleCache.TryGetValue(styleIndex, out var cached)) {
                    return cached;
                }

                var info = textMeasurer.DefaultStyle;
                if (cellFormats != null && fonts != null) {
                    var cellFormat = styleIndex < cellFormats.Count ? cellFormats[(int)styleIndex] : null;
                    if (cellFormat?.FontId != null) {
                        uint fontId = cellFormat.FontId.Value;
                        if (fontId < fonts.Count) {
                            var fontInfo = CreateFontInfoFromOpenXml(fonts[(int)fontId], textMeasurer.DefaultFontSize);
                            info = textMeasurer.CreateStyle(fontInfo);
                        }
                    }
                }

                styleCache[styleIndex] = info;
                return info;
            }

            float MeasureTextWidth(
                string text,
                uint styleIndex,
                ExcelTextMeasurer.Style styleInfo,
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

                        float lineWidth = textMeasurer.MeasureWidthOrDefault(line, styleInfo, 0);
                        if (lineWidth > measured) {
                            measured = lineWidth;
                        }
                    }
                } else if (TryMeasureSimpleAutoFitTextWidth(text, styleIndex, styleInfo, textMeasurer, charWidthCache, out float fastMeasured)) {
                    measured = fastMeasured;
                } else {
                    measured = textMeasurer.MeasureWidthOrDefault(text, styleInfo, 0);
                }

                textWidthCache[(styleIndex, text)] = measured;
                return measured;
            }

            float MeasureRichTextWidth(IReadOnlyList<AutoFitTextRun> runs, ExcelTextMeasurer.Style baseStyle) {
                float maxWidth = 0;
                float currentWidth = 0;

                foreach (var run in runs) {
                    var fontInfo = run.CreateFontInfo(baseStyle.FontInfo);
                    var runStyle = textMeasurer.CreateStyle(fontInfo);
                    string[] parts = run.Text.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None);

                    for (int i = 0; i < parts.Length; i++) {
                        if (i > 0) {
                            if (currentWidth > maxWidth) {
                                maxWidth = currentWidth;
                            }
                            currentWidth = 0;
                        }

                        if (parts[i].Length > 0) {
                            currentWidth += textMeasurer.MeasureWidthOrDefault(parts[i], runStyle, 0);
                        }
                    }
                }

                return Math.Max(maxWidth, currentWidth);
            }

            const double pixelPadding = 2.0;

            void ApplyMeasurement(
                AutoFitMeasurement measurement,
                double[] localWidths,
                Dictionary<uint, ExcelTextMeasurer.Style> styleCache,
                Dictionary<(uint styleIndex, string text), float> textWidthCache,
                Dictionary<uint, Dictionary<char, float>> charWidthCache) {
                    var styleInfo = ResolveStyleInfo(measurement.StyleIndex, styleCache);
                    float textWidthPx = measurement.RichTextRuns != null && measurement.RichTextRuns.Count > 0
                        ? MeasureRichTextWidth(measurement.RichTextRuns, styleInfo)
                        : MeasureTextWidth(measurement.Text, measurement.StyleIndex, styleInfo, textWidthCache, charWidthCache);
                    double cellWidthPx = (textWidthPx * AutoFitTextWidthSafetyFactor) + (2 * pixelPadding) + 1;
                    double columnWidth = Math.Truncate(cellWidthPx / defaultMdw * 256.0) / 256.0;

                    if (columnWidth > localWidths[measurement.TargetIndex]) {
                        localWidths[measurement.TargetIndex] = columnWidth;
                    }
                }

            // Table and AutoFilter header cells are painted bold by the table style and
            // carry a filter button, neither of which their cell format records.
            void ApplyHeaderMeasurements() {
                var headerStyleCache = new Dictionary<uint, ExcelTextMeasurer.Style>();
                var textContext = CreateAutoFitTextContext();
                foreach (AutoFitHeaderCell header in GetAutoFitHeaderCells(plan.Columns)) {
                    ct.ThrowIfCancellationRequested();
                    Cell? cell = TryGetExistingCell(header.Row, header.Column);
                    if (cell == null) continue;
                    string text = GetCellAutoFitText(cell, textContext);
                    if (string.IsNullOrWhiteSpace(text)) continue;
                    uint styleIndex = cell.StyleIndex?.Value ?? 0U;
                    var baseStyle = ResolveStyleInfo(styleIndex, headerStyleCache);
                    var boldStyle = header.Bold && !baseStyle.FontInfo.IsBold
                        ? textMeasurer.CreateStyle(baseStyle.FontInfo.WithStyle(baseStyle.FontInfo.Style | OfficeFontStyle.Bold))
                        : baseStyle;
                    float textWidthPx = 0F;
                    foreach (string line in text.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None)) {
                        if (line.Length > 0) textWidthPx = Math.Max(textWidthPx, textMeasurer.MeasureWidthOrDefault(line, boldStyle, 0));
                    }

                    double buttonPx = header.FilterButton ? AutoFitFilterButtonPixels * boldStyle.Dpi / 96D : 0D;
                    double cellWidthPx = (textWidthPx * AutoFitTextWidthSafetyFactor) + (2 * pixelPadding) + 1 + buttonPx;
                    double columnWidth = Math.Truncate(cellWidthPx / defaultMdw * 256.0) / 256.0;
                    if (columnWidth > widths[header.TargetIndex]) widths[header.TargetIndex] = columnWidth;
                }
            }

            if (!parallel || plan.Measurements.Count < 2) {
                var styleCache = new Dictionary<uint, ExcelTextMeasurer.Style>();
                var textWidthCache = new Dictionary<(uint styleIndex, string text), float>();
                var charWidthCache = new Dictionary<uint, Dictionary<char, float>>();

                for (int i = 0; i < plan.Measurements.Count; i++) {
                    ct.ThrowIfCancellationRequested();
                    ApplyMeasurement(plan.Measurements[i], widths, styleCache, textWidthCache, charWidthCache);
                }

                ApplyHeaderMeasurements();
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

            ApplyHeaderMeasurements();
            return widths;
        }


        /// <summary>Width Excel reserves for the AutoFilter drop-down button, in 96 DPI pixels.</summary>
        private const double AutoFitFilterButtonPixels = 17D;

        private readonly struct AutoFitHeaderCell {
            internal AutoFitHeaderCell(int row, int column, int targetIndex, bool bold, bool filterButton) {
                Row = row;
                Column = column;
                TargetIndex = targetIndex;
                Bold = bold;
                FilterButton = filterButton;
            }

            internal int Row { get; }
            internal int Column { get; }
            internal int TargetIndex { get; }
            internal bool Bold { get; }
            internal bool FilterButton { get; }
        }

        private List<AutoFitHeaderCell> GetAutoFitHeaderCells(IReadOnlyList<int> columns) {
            var targets = new Dictionary<int, int>(columns.Count);
            for (int index = 0; index < columns.Count; index++) targets[columns[index]] = index;
            var headers = new Dictionary<(int Row, int Column), AutoFitHeaderCell>();
            void Add(string? reference, bool bold, bool filterButton) {
                if (!ExcelA1.TryParseRange(reference ?? string.Empty, out int firstRow, out int firstColumn, out _, out int lastColumn)) return;
                for (int column = firstColumn; column <= lastColumn; column++) {
                    if (!targets.TryGetValue(column, out int target)) continue;
                    if (headers.TryGetValue((firstRow, column), out AutoFitHeaderCell existing)) {
                        bold |= existing.Bold;
                        filterButton |= existing.FilterButton;
                    }

                    headers[(firstRow, column)] = new AutoFitHeaderCell(firstRow, column, target, bold, filterButton);
                }
            }

            foreach (TableDefinitionPart part in _worksheetPart.TableDefinitionParts) {
                Table? table = part.Table;
                if (table == null || (table.HeaderRowCount?.Value ?? 1U) == 0U) continue;
                Add(table.Reference?.Value, bold: true, filterButton: table.GetFirstChild<AutoFilter>() != null);
            }

            AutoFilter? sheetFilter = WorksheetRoot.GetFirstChild<AutoFilter>();
            if (sheetFilter != null) Add(sheetFilter.Reference?.Value, bold: false, filterButton: true);
            return headers.Values.ToList();
        }

        internal void ApplyAutoFitColumnWidthsForDeferredMaterialization(double[] widths) {
            if (widths == null || widths.Length == 0) {
                return;
            }

            int[] columnIndexes = new int[widths.Length];
            for (int i = 0; i < columnIndexes.Length; i++) {
                columnIndexes[i] = i + 1;
            }

            SetColumnWidthsCore(columnIndexes, widths);
            MarkRequiresSavePreparation();
            _hasWorksheetMutations = false;
        }


        /// <summary>
        /// Auto-fits the width of the specified column based on its contents.
        /// </summary>
        /// <param name="columnIndex">1-based column index.</param>
        public void AutoFitColumn(int columnIndex) {
            if (columnIndex <= 0) {
                return;
            }

            if (_excelDocument.TryEnableDirectTabularSaveCandidateAutoFit(this, [columnIndex])) {
                return;
            }

            _excelDocument.MaterializeDeferredDataSetImport();
            if (CanSkipStableAutoFitColumns([columnIndex])) {
                return;
            }

            WriteLockConditional(() => {
                var width = CalculateColumnWidths([columnIndex], CancellationToken.None)[0];
                SetColumnWidthCore(columnIndex, width);
                if (EffectiveExecution.SaveWorksheetAfterAutoFit) {
                    WorksheetRoot.Save();
                }
            });
        }
    }
}
