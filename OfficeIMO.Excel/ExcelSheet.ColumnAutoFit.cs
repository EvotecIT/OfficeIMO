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
        /// <summary>
        /// Automatically fits all columns based on their content.
        /// </summary>
        /// <param name="mode">Overrides how the auto-fit work is scheduled across columns.</param>
        /// <param name="ct">Cancels the auto-fit pass while widths are being measured or applied.</param>
        public void AutoFitColumns(ExecutionMode? mode = null, CancellationToken ct = default) {
            if (_excelDocument.TryEnableDirectTabularSaveCandidateAutoFit(this)) {
                return;
            }

            _excelDocument.MaterializeDeferredDataSetImport();
            if (CanSkipStableAutoFitColumns(null)) {
                return;
            }

            var planWatch = EffectiveExecution.OnTiming == null ? null : System.Diagnostics.Stopwatch.StartNew();
            var measurementPlan = BuildAutoFitMeasurementPlanForAllColumns(ct);
            if (planWatch != null) {
                planWatch.Stop();
                EffectiveExecution.ReportTiming("AutoFitColumns.BuildPlan", planWatch.Elapsed);
            }
            if (measurementPlan.Columns.Count == 0) return;
            AutoFitColumnsInternal(measurementPlan, mode, ct);
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
            if (_excelDocument.TryEnableDirectTabularSaveCandidateAutoFit(this, list)) {
                return;
            }

            _excelDocument.MaterializeDeferredDataSetImport();
            if (CanSkipStableAutoFitColumns(list)) {
                return;
            }

            AutoFitColumnsInternal(list, mode, ct);
        }

        private void AutoFitContiguousColumns(int startColumn, int columnCount, ExecutionMode? mode = null, CancellationToken ct = default) {
            if (startColumn <= 0 || columnCount <= 0) return;
            if (startColumn == 1) {
                int[] directCandidateColumns = new int[columnCount];
                for (int i = 0; i < directCandidateColumns.Length; i++) {
                    directCandidateColumns[i] = i + 1;
                }

                if (_excelDocument.TryEnableDirectTabularSaveCandidateAutoFit(this, directCandidateColumns)) {
                    return;
                }
            }

            _excelDocument.MaterializeDeferredDataSetImport();

            var columns = new int[columnCount];
            for (int i = 0; i < columns.Length; i++) {
                columns[i] = startColumn + i;
            }

            if (CanSkipStableAutoFitColumns(columns)) {
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
        public void AutoFitColumnsExcept(IEnumerable<int> columnsToSkip, ExecutionMode? mode = null, CancellationToken ct = default) {
            var skip = new HashSet<int>(columnsToSkip ?? Array.Empty<int>());
            if (_excelDocument.TryGetDirectTabularSaveCandidateColumnCount(this, out int directColumnCount)) {
                var directRemaining = Enumerable.Range(1, directColumnCount).Where(i => !skip.Contains(i)).ToList();
                if (directRemaining.Count == 0 || _excelDocument.TryEnableDirectTabularSaveCandidateAutoFit(this, directRemaining)) {
                    return;
                }
            }

            _excelDocument.MaterializeDeferredDataSetImport();
            var remaining = GetAllColumnIndices().Where(i => i > 0 && !skip.Contains(i)).OrderBy(i => i).ToList();
            if (remaining.Count == 0) return;
            if (CanSkipStableAutoFitColumns(remaining)) {
                return;
            }

            AutoFitColumnsInternal(remaining, mode, ct);
        }


        private void AutoFitColumnsInternal(IReadOnlyList<int> columnsList, ExecutionMode? mode, CancellationToken ct) {
            if (columnsList.Count == 0) return;
            var planWatch = EffectiveExecution.OnTiming == null ? null : System.Diagnostics.Stopwatch.StartNew();
            var measurementPlan = BuildAutoFitMeasurementPlan(columnsList, ct);
            if (planWatch != null) {
                planWatch.Stop();
                EffectiveExecution.ReportTiming("AutoFitColumns.BuildPlan", planWatch.Elapsed);
            }
            AutoFitColumnsInternal(measurementPlan, mode, ct);
        }

        private void AutoFitColumnsInternal(AutoFitMeasurementPlan measurementPlan, ExecutionMode? mode, CancellationToken ct) {
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
            const double columnWidthSafetyFactor = 1.22;

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
                    double cellWidthPx = (textWidthPx * columnWidthSafetyFactor) + (2 * pixelPadding) + 1;
                    double columnWidth = Math.Truncate(cellWidthPx / defaultMdw * 256.0) / 256.0;

                    if (columnWidth > localWidths[measurement.TargetIndex]) {
                        localWidths[measurement.TargetIndex] = columnWidth;
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
