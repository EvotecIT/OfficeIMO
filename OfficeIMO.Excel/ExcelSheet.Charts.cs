using System.Diagnostics;
using OfficeIMO.Excel.OpenXml;
using OfficeIMO.Excel.Utilities;
using OfficeIMO.OpenXml.Internal;
using static OfficeIMO.Excel.OpenXml.ExcelChartXml;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Enumerates charts on the worksheet.
        /// </summary>
        public IEnumerable<ExcelChart> Charts {
            get {
                ExcelPartRef drawingPart = ExcelPartRef.WorksheetForMetadata(this).Drawings;
                OfficeXmlNode worksheetDrawing = drawingPart.IsNone ? default : drawingPart.Root;
                if (worksheetDrawing.IsNone) return Enumerable.Empty<ExcelChart>();
                var charts = new List<ExcelChart>();
                foreach (OfficeXmlNode frame in worksheetDrawing.Descendants(Xdr, "graphicFrame")) {
                    if (frame.Child(A, "graphic").Child(A, "graphicData").Child(C, "chart").Exists) {
                        charts.Add(new ExcelChart(frame, drawingPart, this));
                    }
                }
                return charts;
            }
        }

        /// <summary>
        /// Returns a chart by name (non-visual name), or null if not found.
        /// </summary>
        public ExcelChart? GetChart(string name) {
            if (string.IsNullOrWhiteSpace(name)) return null;
            return Charts.FirstOrDefault(c => string.Equals(c.Name, name, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Writes chart data into the worksheet and returns the corresponding data range.
        /// </summary>
        public ExcelChartDataRange WriteChartData(ExcelChartData data, int startRow = 1, int startColumn = 1, string? categoryHeader = null,
            bool includeHeaderRow = true, bool numericCategories = false, ExcelChartDataOrientation orientation = ExcelChartDataOrientation.Vertical) {
            ArgumentNullException.ThrowIfNull(data);
            if (startRow <= 0 || startColumn <= 0) throw new ArgumentOutOfRangeException(nameof(startRow));

            IReadOnlyList<string> categories = data.Categories;
            IReadOnlyList<ExcelChartSeries> series = data.Series;
            int categoryCount = GetChartDataPointCount(data);
            int seriesCount = series.Count;
            int cellCapacity = categoryCount * (seriesCount + 1);
            if (includeHeaderRow) {
                cellCapacity += seriesCount + 1;
            }

            var cells = new List<(int Row, int Column, object Value)>(cellCapacity);
            if (orientation == ExcelChartDataOrientation.Vertical) {
                int rowOffset = includeHeaderRow ? 1 : 0;
                if (includeHeaderRow) {
                    string header = categoryHeader ?? string.Empty;
                    cells.Add((startRow, startColumn, header));

                    for (int s = 0; s < seriesCount; s++) {
                        cells.Add((startRow, startColumn + s + 1, series[s].Name));
                    }
                }

                for (int i = 0; i < categoryCount; i++) {
                    int row = startRow + i + rowOffset;
                    object categoryValue = GetChartCategoryValue(categories, series, i, numericCategories);
                    cells.Add((row, startColumn, categoryValue));
                    for (int s = 0; s < seriesCount; s++) {
                        if (i < series[s].Values.Count) {
                            cells.Add((row, startColumn + s + 1, series[s].Values[i]));
                        }
                    }
                }
            } else {
                int columnOffset = includeHeaderRow ? 1 : 0;
                if (includeHeaderRow) {
                    cells.Add((startRow, startColumn, categoryHeader ?? string.Empty));
                }

                for (int i = 0; i < categoryCount; i++) {
                    cells.Add((startRow, startColumn + i + columnOffset, GetChartCategoryValue(categories, series, i, numericCategories)));
                }

                for (int s = 0; s < seriesCount; s++) {
                    int row = startRow + s + 1;
                    if (includeHeaderRow) {
                        cells.Add((row, startColumn, series[s].Name));
                    }

                    for (int i = 0; i < categoryCount; i++) {
                        if (i < series[s].Values.Count) {
                            cells.Add((row, startColumn + i + columnOffset, series[s].Values[i]));
                        }
                    }
                }
            }

            using (var preserveDirectDataSet = _excelDocument.PreserveDirectDataSetSaveCandidateDuringDirtyMarks()) {
                if (!TryApplyPlainCellsByAppendingRows(cells, CancellationToken.None)) {
                    SetValues(cells, null);
                }
            }

            return new ExcelChartDataRange(Name, startRow, startColumn, categoryCount, seriesCount, includeHeaderRow, orientation);
        }

        private static int GetChartDataPointCount(ExcelChartData data) {
            int pointCount = data.Categories.Count;
            foreach (ExcelChartSeries series in data.Series) {
                pointCount = Math.Max(pointCount, series.Values.Count);
                if (series.XValues != null) {
                    pointCount = Math.Max(pointCount, series.XValues.Count);
                }
            }

            return pointCount;
        }

        private static object GetChartCategoryValue(IReadOnlyList<string> categories, IReadOnlyList<ExcelChartSeries> series, int index, bool numericCategories) {
            if (index < categories.Count) {
                return CoerceChartCategoryValue(categories[index], numericCategories);
            }

            if (numericCategories) {
                foreach (ExcelChartSeries item in series) {
                    if (item.XValues != null && index < item.XValues.Count) {
                        return item.XValues[index];
                    }
                }

                return index + 1D;
            }

            return string.Empty;
        }

        private static object CoerceChartCategoryValue(string categoryText, bool numericCategories) {
            if (!numericCategories) {
                return categoryText;
            }

            if (!OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(categoryText, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double numeric)) {
                throw new ArgumentException($"Category '{categoryText}' is not numeric. Scatter charts require numeric X values. Use AddScatterChartFromRanges for non-numeric categories.");
            }

            return numeric;
        }

        /// <summary>
        /// Adds a chart to the worksheet using the provided data. Data is stored on a hidden chart data sheet.
        /// </summary>
        private ExcelChart AddChart(ExcelChartData data, int row, int column, OfficeLength? width = null, OfficeLength? height = null,
            OfficeChartType type = OfficeChartType.ColumnClustered, string? title = null) {
            width ??= OfficeLength.Pixels(640);
            height ??= OfficeLength.Pixels(360);
            ArgumentNullException.ThrowIfNull(data);
            ValidateChartTypes(type, data);
            ValidateChartPlacement(row, column, width.Value, height.Value);

            using var preserveFastSaveState = _excelDocument.PreserveDirectDataSetFastSaveStateDuringDirtyMarks();
            Stopwatch? stageWatch = EffectiveExecution.OnTiming == null ? null : Stopwatch.StartNew();
            var dataSheet = _excelDocument.GetOrCreateChartDataSheet();
            ReportChartTiming(stageWatch, "AddChart.GetOrCreateChartDataSheet");

            stageWatch?.Restart();
            int startRow = _excelDocument.ReserveChartDataStartRow(dataSheet, GetChartDataPointCount(data) + 1);
            ReportChartTiming(stageWatch, "AddChart.ReserveChartData");

            stageWatch?.Restart();
            bool numericCategories = RequiresNumericScatterCategories(data.Series, type);
            ExcelChartDataRange range = dataSheet.WriteChartData(data, startRow, 1, numericCategories: numericCategories);
            ReportChartTiming(stageWatch, "AddChart.WriteChartData");

            return AddChartInternal(range, row, column, width.Value, height.Value, type, data, title, preserveDeferredDataSetCandidate: true);
        }

        private static bool RequiresNumericScatterCategories(IReadOnlyList<ExcelChartSeries> series,
            OfficeChartType defaultType) {
            for (int i = 0; i < series.Count; i++) {
                ExcelChartSeries item = series[i];
                if ((item.ChartType ?? defaultType) == OfficeChartType.Scatter && item.XValues == null) {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Adds a chart to the worksheet using an existing data range.
        /// </summary>
        private ExcelChart AddChart(ExcelChartDataRange dataRange, int row, int column, OfficeLength? width = null, OfficeLength? height = null,
            OfficeChartType type = OfficeChartType.ColumnClustered, ExcelChartData? cachedData = null, string? title = null) {
            width ??= OfficeLength.Pixels(640);
            height ??= OfficeLength.Pixels(360);
            ArgumentNullException.ThrowIfNull(dataRange);
            ValidateChartPlacement(row, column, width.Value, height.Value);

            return AddChartInternal(dataRange, row, column, width.Value, height.Value, type, cachedData, title);
        }

        /// <summary>
        /// Adds a chart using an A1 range on this sheet.
        /// </summary>
        private ExcelChart AddChartFromRange(string dataRangeA1, int row, int column, OfficeLength? width = null, OfficeLength? height = null,
            OfficeChartType type = OfficeChartType.ColumnClustered, bool hasHeaders = true, string? title = null, bool includeCachedData = true,
            IReadOnlyList<string>? nativeSeriesNames = null) {
            width ??= OfficeLength.Pixels(640);
            height ??= OfficeLength.Pixels(360);
            if (string.IsNullOrWhiteSpace(dataRangeA1)) throw new ArgumentNullException(nameof(dataRangeA1));
            if (!ExcelA1.TryParseRange(dataRangeA1, out int r1, out int c1, out int r2, out int c2)) {
                throw new ArgumentException($"Invalid A1 range '{dataRangeA1}'.", nameof(dataRangeA1));
            }

            int categoryCount = hasHeaders ? r2 - r1 : r2 - r1 + 1;
            int seriesCount = c2 - c1;
            if (categoryCount <= 0 || seriesCount <= 0) {
                throw new InvalidOperationException("Range must include at least one category row and one series column.");
            }

            var range = new ExcelChartDataRange(Name, r1, c1, categoryCount, seriesCount, hasHeaderRow: hasHeaders);
            ExcelChartData? data = includeCachedData ? ExcelChartUtils.TryReadChartData(this, range) : null;
            return AddChartInternal(range, row, column, width.Value, height.Value, type, data, title,
                preserveDeferredDataSetCandidate: true, nativeSeriesNames: nativeSeriesNames);
        }

        /// <summary>
        /// Adds a chart from a table name on this sheet.
        /// </summary>
        private ExcelChart AddChartFromTable(string tableName, int row, int column, OfficeLength? width = null, OfficeLength? height = null,
            OfficeChartType type = OfficeChartType.ColumnClustered, string? title = null, bool includeCachedData = true) {
            width ??= OfficeLength.Pixels(640);
            height ??= OfficeLength.Pixels(360);
            if (string.IsNullOrWhiteSpace(tableName)) throw new ArgumentNullException(nameof(tableName));

            ChartSourceTable? table = FindChartSourceTable(tableName);
            if (table?.Reference == null) {
                throw new InvalidOperationException($"Table '{tableName}' was not found on sheet '{Name}'.");
            }

            bool hasHeaders = (table.HeaderRowCount ?? 1U) > 0;
            var (firstRow, firstColumn, lastRow, lastColumn) = ExcelA1.ParseRange(table.Reference);
            int totals = checked((int)(table.TotalsRowCount ?? 0U));
            string sourceRange = ExcelA1.CellReference(firstRow, firstColumn) + ":" +
                ExcelA1.CellReference(lastRow - totals, lastColumn);
            IReadOnlyList<string>? nativeSeriesNames = hasHeaders && includeCachedData ? null : table.ColumnNames?
                .Skip(1).ToArray();
            return AddChartFromRange(sourceRange, row, column, width.Value, height.Value, type, hasHeaders, title,
                includeCachedData, nativeSeriesNames);
        }

        /// <summary>
        /// Adds a chart using an A1 range and marks it as sourced from an existing pivot table.
        /// </summary>
        public ExcelChart AddPivotChartFromRange(string pivotTableName, string dataRangeA1, int row, int column, OfficeLength? width = null, OfficeLength? height = null,
            OfficeChartType type = OfficeChartType.ColumnClustered, bool hasHeaders = true, string? title = null, bool includeCachedData = true, uint formatId = 0U) {
            width ??= OfficeLength.Pixels(640);
            height ??= OfficeLength.Pixels(360);
            if (string.IsNullOrWhiteSpace(pivotTableName)) throw new ArgumentNullException(nameof(pivotTableName));

            string? resolvedName = GetPivotTables()
                .FirstOrDefault(p => string.Equals(p.Name, pivotTableName, StringComparison.OrdinalIgnoreCase))
                ?.Name;
            if (string.IsNullOrWhiteSpace(resolvedName)) {
                throw new InvalidOperationException($"Pivot table '{pivotTableName}' was not found on sheet '{Name}'.");
            }

            ExcelChart chart = AddChartFromRange(dataRangeA1, row, column, width.Value, height.Value, type, hasHeaders, title, includeCachedData);
            chart.SetPivotSource(resolvedName!, formatId);
            return chart;
        }

        /// <summary>
        /// Adds a scatter chart using explicit X/Y ranges.
        /// </summary>
        public ExcelChart AddScatterChartFromRanges(IEnumerable<ExcelChartSeriesRange> seriesRanges, int row, int column, OfficeLength? width = null,
            OfficeLength? height = null, string? title = null) {
            width ??= OfficeLength.Pixels(640);
            height ??= OfficeLength.Pixels(360);
            return AddChartFromSeriesRanges(seriesRanges, row, column, width.Value, height.Value, OfficeChartType.Scatter, title);
        }

        /// <summary>
        /// Adds a bubble chart using explicit X/Y/size ranges.
        /// </summary>
        public ExcelChart AddBubbleChartFromRanges(IEnumerable<ExcelChartSeriesRange> seriesRanges, int row, int column, OfficeLength? width = null,
            OfficeLength? height = null, string? title = null) {
            width ??= OfficeLength.Pixels(640);
            height ??= OfficeLength.Pixels(360);
            return AddChartFromSeriesRanges(seriesRanges, row, column, width.Value, height.Value, OfficeChartType.Bubble, title);
        }

        private ExcelChart AddChartFromSeriesRanges(IEnumerable<ExcelChartSeriesRange> seriesRanges, int row, int column, OfficeLength width,
            OfficeLength height, OfficeChartType type, string? title) {
            ArgumentNullException.ThrowIfNull(seriesRanges);
            ValidateChartPlacement(row, column, width, height);

            var ranges = seriesRanges as IReadOnlyList<ExcelChartSeriesRange> ?? seriesRanges.ToList();
            if (ranges.Count == 0) {
                throw new ArgumentException("At least one series range is required.", nameof(seriesRanges));
            }
            if (type == OfficeChartType.Bubble && HasMissingBubbleSizeRange(ranges)) {
                throw new ArgumentException("Bubble charts require bubble size ranges for each series.", nameof(seriesRanges));
            }

            OfficeXmlNode frame = default;
            ExcelPartRef drawingPart = default;

            using var preserveDirectDataSet = _excelDocument.PreserveDirectDataSetSaveCandidateDuringDirtyMarks();
            WriteLockWorksheetPreparationOnly(() => {
                drawingPart = GetOrCreateDrawingsPart();

                _excelDocument.EnsureWorkbookThemeAndStyles();

                EnsureChartAxisIds();
                ExcelPartRef chartPart = drawingPart.AddNewPart(ExcelPartKind.Chart);
                PopulateChartFromSeriesRanges(chartPart, type, Name, ranges, title);
                if (_excelDocument.DefaultChartStylePreset == null)
                    ApplyDefaultChartTheme(chartPart);
                if (_excelDocument.DefaultChartStylePreset != null) {
                    ApplyChartStyle(chartPart, _excelDocument.DefaultChartStylePreset);
                }
                string chartRelId = drawingPart.GetIdOfPart(chartPart)!;

                long cx = width.ToEmus();
                long cy = height.ToEmus();

                uint nvId = NextDrawingId(drawingPart);
                frame = AppendGraphicFrameAnchor(drawingPart.Root, row, column, cx, cy, nvId, $"Chart {nvId}",
                    ChartGraphicDataUri, graphicData => WithRelationshipId(graphicData.New(C, "chart"), chartRelId));
                drawingPart.SaveRoot();
            });

            return new ExcelChart(frame!, drawingPart!, this);
        }

        private static bool HasMissingBubbleSizeRange(IReadOnlyList<ExcelChartSeriesRange> ranges) {
            for (int i = 0; i < ranges.Count; i++) {
                if (string.IsNullOrWhiteSpace(ranges[i].BubbleSizeRangeA1)) {
                    return true;
                }
            }

            return false;
        }

        private ExcelChart AddChartInternal(ExcelChartDataRange range, int row, int column, OfficeLength width, OfficeLength height,
            OfficeChartType type, ExcelChartData? data, string? title, bool preserveDeferredDataSetCandidate = false, IReadOnlyList<string>? nativeSeriesNames = null,
            ExcelChartRangeSeries? rangeSeries = null) {
            ValidateChartTypes(type, data);
            ValidateChartPlacement(row, column, width, height);
            if (data != null && !preserveDeferredDataSetCandidate) {
                MaterializeDeferredDataSetImportIfNeeded();
            }

            OfficeXmlNode frame = default;
            ExcelPartRef drawingPart = default;

            using var preserveDirectDataSet = data == null || preserveDeferredDataSetCandidate
                ? _excelDocument.PreserveDirectDataSetFastSaveStateDuringDirtyMarks()
                : null;
            WriteLockWorksheetPreparationOnly(() => {
                Stopwatch? stageWatch = EffectiveExecution.OnTiming == null ? null : Stopwatch.StartNew();
                drawingPart = GetOrCreateDrawingsPart();
                ReportChartTiming(stageWatch, "AddChartInternal.PrepareDrawingPart");

                stageWatch?.Restart();
                _excelDocument.EnsureWorkbookThemeAndStyles();
                ReportChartTiming(stageWatch, "AddChartInternal.EnsureThemeAndStyles");

                stageWatch?.Restart();
                EnsureChartAxisIds();
                ExcelPartRef chartPart = drawingPart.AddNewPart(ExcelPartKind.Chart);
                ReportChartTiming(stageWatch, "AddChartInternal.AddChartPart");

                stageWatch?.Restart();
                PopulateChart(chartPart, type, range, data, title, nativeSeriesNames);
                if (rangeSeries != null) ApplyRangeSeriesFormulas(rangeSeries, chartPart);
                if (_excelDocument.DefaultChartStylePreset == null)
                    ApplyDefaultChartTheme(chartPart);
                ReportChartTiming(stageWatch, "AddChartInternal.PopulateChart");

                if (_excelDocument.DefaultChartStylePreset != null) {
                    stageWatch?.Restart();
                    ApplyChartStyle(chartPart, _excelDocument.DefaultChartStylePreset);
                    ReportChartTiming(stageWatch, "AddChartInternal.ApplyChartStyle");
                }

                stageWatch?.Restart();
                string chartRelId = drawingPart.GetIdOfPart(chartPart)!;

                long cx = width.ToEmus();
                long cy = height.ToEmus();

                uint nvId = NextDrawingId(drawingPart);
                frame = AppendGraphicFrameAnchor(drawingPart.Root, row, column, cx, cy, nvId, $"Chart {nvId}",
                    ChartGraphicDataUri, graphicData => WithRelationshipId(graphicData.New(C, "chart"), chartRelId));
                ReportChartTiming(stageWatch, "AddChartInternal.BuildAnchor");

                stageWatch?.Restart();
                drawingPart.SaveRoot();
                ReportChartTiming(stageWatch, "AddChartInternal.SaveDrawing");
            });

            if (preserveDeferredDataSetCandidate) {
                _excelDocument.PreserveDeferredDataSetFastSaveModelAndClearCandidate();
            }

            return new ExcelChart(frame, drawingPart, this, rangeSeries == null ? range : null);
        }

        private static OfficeXmlNode WithRelationshipId(OfficeXmlNode element, string relationshipId) {
            element.SetAttribute(R, "id", relationshipId);
            return element;
        }

        private const string ChartGraphicDataUri = "http://schemas.openxmlformats.org/drawingml/2006/chart";

        /// <summary>
        /// Appends a one-cell anchor holding a graphic frame to the worksheet drawing and returns the frame as held. The
        /// frame's graphic data gets the uri and the reference <paramref name="createReference"/> makes for it.
        /// </summary>
        private static OfficeXmlNode AppendGraphicFrameAnchor(OfficeXmlNode worksheetDrawing, int row, int column, long cx, long cy,
            uint id, string name, string graphicDataUri, Func<OfficeXmlNode, OfficeXmlNode> createReference) {
            OfficeXmlNode anchor = worksheetDrawing.New(Xdr, "oneCellAnchor");
            OfficeXmlNode frame = anchor.New(Xdr, "graphicFrame");
            OfficeXmlNode nonVisual = frame.New(Xdr, "nvGraphicFramePr");
            OfficeXmlNode frameDrawing = nonVisual.New(Xdr, "cNvGraphicFramePr");
            OfficeXmlNode transform = frame.New(Xdr, "xfrm");
            OfficeXmlNode graphic = frame.New(A, "graphic");
            OfficeXmlNode graphicData = graphic.New(A, "graphicData").Attr("uri", graphicDataUri);
            frame.With(
                nonVisual.With(
                    nonVisual.New(Xdr, "cNvPr").Attr("id", id).Attr("name", name),
                    frameDrawing.With(frameDrawing.New(A, "graphicFrameLocks").Attr("noChangeAspect", true))
                ),
                transform.With(
                    transform.New(A, "off").Attr("x", 0L).Attr("y", 0L),
                    transform.New(A, "ext").Attr("cx", cx).Attr("cy", cy)
                ),
                graphic.With(graphicData.With(createReference(graphicData)))
            );

            anchor.With(
                NewMarker(anchor, "from", column - 1, 0L, row - 1, 0L),
                anchor.New(Xdr, "ext").Attr("cx", cx).Attr("cy", cy),
                frame,
                anchor.New(Xdr, "clientData")
            );

            anchor = worksheetDrawing.Append(anchor);
            return anchor.Child(Xdr, "graphicFrame");
        }

        private static void ValidateChartPlacement(int row, int column, OfficeLength width, OfficeLength height) {
            if (row <= 0 || row > ExcelA1.MaxRows) throw new ArgumentOutOfRangeException(nameof(row));
            if (column <= 0 || column > ExcelA1.MaxColumns) throw new ArgumentOutOfRangeException(nameof(column));
            if (width <= OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(width));
            if (height <= OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(height));
        }

        private void ReportChartTiming(Stopwatch? stopwatch, string operation) {
            if (stopwatch != null) {
                EffectiveExecution.ReportTiming(operation, stopwatch.Elapsed);
            }
        }
    }
}
