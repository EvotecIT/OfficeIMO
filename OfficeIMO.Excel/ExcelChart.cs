using System.Globalization;
using OfficeIMO.Drawing;
using OfficeIMO.Excel.OpenXml;
using OfficeIMO.Excel.Utilities;
using OfficeIMO.OpenXml.Internal;
using static OfficeIMO.Excel.OpenXml.ExcelChartXml;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Represents a chart on a worksheet.
    /// </summary>
    public sealed partial class ExcelChart {
        private readonly ExcelDrawingHandle _drawingHandle;
        /// <summary>The chart's <c>xdr:graphicFrame</c> in the worksheet drawing.</summary>
        private OfficeXmlNode FrameNode => _drawingHandle.Element;
        /// <summary>The worksheet drawing part that holds the frame.</summary>
        private ExcelPartRef DrawingsPartRef => _drawingHandle.Part;
        private readonly ExcelDocument _document;
        private string _sheetName => _drawingHandle.Sheet.Name;
        private string _lastKnownSheetName;
        private ExcelChartDataRange? _dataRange;
        private bool _dataRangeNeedsRefresh;

        internal ExcelChart(OfficeXmlNode frame, ExcelPartRef drawingsPart, ExcelSheet sheet, ExcelChartDataRange? dataRange = null) {
            ArgumentNullException.ThrowIfNull(sheet);
            _document = sheet.Document;
            _lastKnownSheetName = sheet.Name;
            _dataRange = dataRange;
            _drawingHandle = new ExcelDrawingHandle(frame, drawingsPart, sheet, RefreshDataRangeAfterReload);
        }

        /// <summary>
        /// Re-reads the data range from the reloaded chart part on next use, so a save does not parse every chart again.
        /// </summary>
        private void RefreshDataRangeAfterReload() => _dataRangeNeedsRefresh = true;

        /// <summary>
        /// Gets or sets the chart name (non-visual drawing name).
        /// </summary>
        public string Name {
            get => DrawingProperties.Attribute("name") ?? string.Empty;
            set {
                OfficeXmlNode props = DrawingProperties;
                if (props.Exists) {
                    props.SetAttribute("name", value ?? string.Empty);
                    DrawingsPartRef.SaveRoot();
                    MarkChartDirty();
                }
            }
        }

        /// <summary>
        /// Gets the chart data range when it is known, including the worksheet range a saved chart references.
        /// </summary>
        public ExcelChartDataRange? DataRange {
            get {
                // Resolving the part also refreshes a range invalidated by a save.
                ExcelPartRef chartPart = ChartPartRef();
                return _dataRange ?? ExcelChartUtils.TryExtractDataRange(chartPart);
            }
        }

        internal ExcelChartDataRange? GetSourceDataRange(int maximumSeries) {
            ExcelChartDataRange? range = DataRange;
            return range != null
                ? range.SeriesCount <= maximumSeries ? range : null
                : ExcelChartUtils.TryExtractDataRange(ChartPartRef(), maximumSeries);
        }

        internal bool HasCanonicalWorksheetReferences() {
            ExcelChartDataRange? range = DataRange;
            return range != null && ExcelChartUtils.HasCanonicalWorksheetReferences(ChartPartRef(), range);
        }

        internal bool HasCanonicalWorksheetReferences(ExcelChartDataRange range) =>
            ExcelChartUtils.HasCanonicalWorksheetReferences(ChartPartRef(), range);

        /// <summary>
        /// Gets this chart anchor's zero-based order in the worksheet drawing layer.
        /// </summary>
        public int DrawingOrder => GetDrawingOrder();

        internal bool HasAbsoluteAnchor => FrameNode.Ancestor(Xdr, "absoluteAnchor").Exists;

        internal bool TryGetAbsoluteAnchorBounds(out int xPixels, out int yPixels, out int widthPixels, out int heightPixels) {
            xPixels = 0;
            yPixels = 0;
            widthPixels = 0;
            heightPixels = 0;
            OfficeXmlNode absoluteAnchor = FrameNode.Ancestor(Xdr, "absoluteAnchor");
            if (absoluteAnchor.IsNone) {
                return false;
            }

            OfficeXmlNode position = absoluteAnchor.Child(Xdr, "pos");
            OfficeXmlNode extent = absoluteAnchor.Child(Xdr, "ext");
            xPixels = ExcelImageExportLimits.ClampAbsoluteOffsetPixels(EmuOffsetToPixels(position.Int64Attribute("x") ?? 0L));
            yPixels = ExcelImageExportLimits.ClampAbsoluteOffsetPixels(EmuOffsetToPixels(position.Int64Attribute("y") ?? 0L));
            (int fallbackWidthPixels, int fallbackHeightPixels) = GetAnchorSizePixels();
            widthPixels = ExcelImageExportLimits.ClampExtentPixels(EmuToPixels(extent.Int64Attribute("cx"), fallbackWidthPixels));
            heightPixels = ExcelImageExportLimits.ClampExtentPixels(EmuToPixels(extent.Int64Attribute("cy"), fallbackHeightPixels));
            return widthPixels > 0 && heightPixels > 0;
        }

        /// <summary>
        /// Gets the detected chart type.
        /// </summary>
        public OfficeChartType ChartType {
            get {
                OfficeXmlNode plotArea = ChartNode().Child(C, "plotArea");
                return plotArea.IsNone ? OfficeChartType.ColumnClustered : ExcelChartUtils.InferChartType(plotArea);
            }
        }

        /// <summary>
        /// Gets the chart title text when present.
        /// </summary>
        public string? Title => GetChartTitleText(ChartNode());

        /// <summary>
        /// Tries to read the chart data from the chart's source range.
        /// </summary>
        public bool TryGetData(out ExcelChartData data) => TryGetData(out data, out _);

        /// <summary>Reads the chart data and reports why a strict projection rejected the chart while reading it.</summary>
        private bool TryGetData(out ExcelChartData data, out string? unsupportedReason) {
            unsupportedReason = null;
            try {
                ExcelPartRef chartPart = ChartPartRef();
                ExcelChartDataRange? range = _dataRange ?? ExcelChartUtils.TryExtractDataRange(chartPart);
                ExcelSheet sheet = range == null ? _document[_sheetName] : _document[range.SheetName];
                ExcelChartData? chartData = range == null
                    ? ExcelChartUtils.TryReadChartData(chartPart, sheet)
                    : ExcelChartUtils.TryReadChartData(sheet, range, keepBlankCells: true);
                if (chartData == null) {
                    data = null!;
                    return false;
                }

                chartData = ExcelChartUtils.ApplyChartSeriesTypes(chartPart, chartData, ChartType);
                chartData = ExcelChartUtils.ApplyScatterSeriesXValues(chartPart, chartData, sheet);
                bool[]? hiddenSeries = null;
                // Custom error amounts lose the same hidden cells as the values, whenever the values lost any.
                ExcelSheet? selectVisibleAmountsFrom = null;
                if (OfficeOpenXmlChartSourceVisibility.PlotsVisibleCellsOnly(chartPart.Root.Child(C, "chart"))) {
                    ExcelChartData? allPoints = chartData;
                    chartData = ExcelChartUtils.SelectVisiblePoints(chartPart, sheet, chartData, out hiddenSeries);
                    if (chartData == null) {
                        data = null!;
                        return false;
                    }
                    if (!ReferenceEquals(chartData, allPoints)) selectVisibleAmountsFrom = sheet;
                }
                chartData = ApplyImageExportSeriesStyles(chartPart, chartData, ExcelPartRef.Workbook(_document),
                    selectVisibleAmountsFrom == null ? null : ExcelChartUtils.CreateVisibleNodeAmountSelector(selectVisibleAmountsFrom));
                chartData = ExcelChartUtils.RemoveSeries(chartData, hiddenSeries);
                // With every series hidden Excel plots nothing, which no snapshot can stand for.
                if (chartData.Series.Count == 0) {
                    data = null!;
                    return false;
                }
                if (range != null) {
                    _dataRange = range;
                }
                data = chartData;
                return true;
            } catch (Exception exception) {
                if (exception is NotSupportedException) unsupportedReason = exception.Message;
                data = null!;
                return false;
            }
        }

        /// <summary>
        /// Tries to create a dependency-free snapshot for rendering/export consumers.
        /// </summary>
        public bool TryGetSnapshot(out ExcelChartInfo snapshot) =>
            TryGetSnapshotCore(null, out snapshot, out _);

        internal bool TryGetSnapshot(ExcelWorksheetGeometryIndex geometry, out ExcelChartInfo snapshot) =>
            TryGetSnapshotCore(geometry, OfficeChartFidelity.Strict, out snapshot, out _);

        /// <summary>Creates a rendering snapshot at the requested chart fidelity.</summary>
        internal bool TryGetSnapshot(OfficeChartFidelity fidelity, out ExcelChartInfo snapshot) =>
            TryGetSnapshotCore(null, fidelity, out snapshot, out _);

        /// <summary>Creates a worksheet rendering snapshot at the requested chart fidelity.</summary>
        internal bool TryGetSnapshot(ExcelWorksheetGeometryIndex geometry, OfficeChartFidelity fidelity, out ExcelChartInfo snapshot) =>
            TryGetSnapshotCore(geometry, fidelity, out snapshot, out _);

        /// <summary>Reads the snapshot at the requested chart fidelity and reports why a chart whose data resolved cannot be projected.</summary>
        internal bool TryGetSnapshot(OfficeChartFidelity fidelity, out ExcelChartInfo snapshot, out string? unsupportedReason) =>
            TryGetSnapshotCore(null, fidelity, out snapshot, out unsupportedReason);

        private bool TryGetSnapshotCore(ExcelWorksheetGeometryIndex? geometry, out ExcelChartInfo snapshot, out string? unsupportedReason) =>
            TryGetSnapshotCore(geometry, OfficeChartFidelity.Strict, out snapshot, out unsupportedReason);

        private bool TryGetSnapshotCore(ExcelWorksheetGeometryIndex? geometry, OfficeChartFidelity fidelity, out ExcelChartInfo snapshot,
            out string? unsupportedReason) {
            unsupportedReason = null;
            using var scope = OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartFidelityScope.Enter(fidelity);
            if (!TryGetData(out ExcelChartData data, out unsupportedReason)) {
                snapshot = null!;
                return false;
            }

            try {
                (OfficeLength anchorWidth, OfficeLength anchorHeight) = GetAnchorSize(geometry);
                OfficeXmlNode plot = ChartSpaceNode().Child(C, "chart").Child(C, "plotArea");
                OfficeOpenXmlChartSecondaryAxis.QualifyLinearProjection(plot);
                var bubble = ReadBubbleImageSettings(plot, data);
                OfficeChartStyle? style = CreateImageExportStyle(data.Categories.Count);
                OfficeChartLayout? layout = CreateImageExportLayout();
                IReadOnlyList<OfficeImageExportDiagnostic> diagnostics = CreateImageExportDiagnostics().Concat(bubble.Diagnostics)
                    .Concat(scope.Approximations.Select(approximation => new OfficeImageExportDiagnostic(
                        OfficeImageExportDiagnosticSeverity.Warning, OfficeDiagnosticCodes.ChartApproximation, approximation,
                        GetImageExportSource(), OfficeConversionLossKind.Approximation))).ToArray();
                snapshot = new ExcelChartInfo(
                    Name,
                    Title ?? OfficeOpenXmlChartPlotSemantics.ReadAutomaticTitle(ChartNode(),
                        data.Series.Count == 1 ? data.Series[0].Name : null),
                    ChartType,
                    data,
                    GetAnchorRow(),
                    GetAnchorColumn(),
                    GetAnchorOffset(horizontal: true),
                    GetAnchorOffset(horizontal: false),
                    anchorWidth,
                    anchorHeight,
                    style,
                    layout,
                    diagnostics,
                    RadialLayout, bubble.Scale, bubble.Mode, bubble.ShowNegative);
                return true;
            } catch (Exception exception) when (exception is ArgumentException || exception is FormatException ||
                exception is OverflowException || exception is NotSupportedException) {
                snapshot = null!;
                unsupportedReason = exception is NotSupportedException ? exception.Message : exception.GetType().Name + ": " + exception.Message;
                return false;
            }
        }

        /// <summary>
        /// Gets whether the chart declares a pivot table source.
        /// </summary>
        public bool IsPivotChart => !string.IsNullOrWhiteSpace(PivotTableName);

        /// <summary>
        /// Gets the pivot table name used as this chart's pivot source, if present.
        /// </summary>
        public string? PivotTableName => ChartPartRef().Root.Child(C, "pivotSource").ChildText(C, "name");

        /// <summary>
        /// Marks the chart as sourced from a pivot table.
        /// </summary>
        /// <param name="pivotTableName">Pivot table name to assign as the chart's pivot source.</param>
        /// <param name="formatId">Pivot chart format id.</param>
        public ExcelChart SetPivotSource(string pivotTableName, uint formatId = 0U) {
            if (string.IsNullOrWhiteSpace(pivotTableName)) {
                throw new ArgumentNullException(nameof(pivotTableName));
            }

            OfficeXmlNode chartSpace = ChartPartRef().Root;
            if (chartSpace.IsNone) throw new InvalidOperationException("Chart space not found in chart part.");
            chartSpace.RemoveChildren(C, "pivotSource");

            OfficeXmlNode pivotSource = chartSpace.New(C, "pivotSource");
            pivotSource.With(
                pivotSource.NewText(C, "name", pivotTableName.Trim()),
                pivotSource.NewVal(C, "fmtId", formatId));

            OfficeXmlNode chart = chartSpace.Child(C, "chart");
            if (chart.Exists) {
                chartSpace.InsertBefore(pivotSource, chart);
            } else {
                chartSpace.Append(pivotSource);
            }

            Save();
            return this;
        }

        /// <summary>
        /// Updates the chart data (series and categories).
        /// </summary>
        public ExcelChart UpdateData(ExcelChartData data, ExcelChartDataRange? dataRange = null, bool writeToSheet = true) {
            ArgumentNullException.ThrowIfNull(data);

            ExcelPartRef chartPart = ChartPartRef();
            ExcelChartDataRange? resolved = dataRange ?? _dataRange ?? ExcelChartUtils.TryExtractDataRange(chartPart);
            if (resolved == null) {
                throw new InvalidOperationException("Chart data range could not be resolved. Provide a data range explicitly.");
            }

            resolved = resolved.WithSize(data.Categories.Count, data.Series.Count);

            if (writeToSheet) {
                var targetSheet = _document[resolved.SheetName];
                bool numericCategories = chartPart.Root.Child(C, "chart").Child(C, "plotArea").Child(C, "scatterChart").Exists;
                targetSheet.WriteChartData(data, resolved.StartRow, resolved.StartColumn, includeHeaderRow: resolved.HasHeaderRow, numericCategories: numericCategories, orientation: resolved.Orientation);
            }

            ExcelChartUtils.UpdateChartData(chartPart, data, resolved);
            _dataRange = resolved;
            Save();
            return this;
        }

        /// <summary>
        /// Updates the chart data using selectors.
        /// </summary>
        public ExcelChart UpdateData<T>(IEnumerable<T> items, Func<T, string> categorySelector,
            params ExcelChartSeriesDefinition<T>[] seriesDefinitions) {
            var data = ExcelChartData.From(items, categorySelector, seriesDefinitions);
            return UpdateData(data);
        }

        /// <summary>
        /// Sets the chart title text.
        /// </summary>
        public ExcelChart SetTitle(string title) {
            ArgumentNullException.ThrowIfNull(title);

            OfficeXmlNode chart = ChartNode();
            chart.SetOrdered(chart.NewVal(C, "autoTitleDeleted", false), CChartView.Order);

            OfficeXmlNode existingTitle = chart.Child(C, "title");
            OfficeXmlNode chartTitle = existingTitle.Exists ? existingTitle : chart.New(C, "title");
            chartTitle.RemoveChildren(C, "tx");
            chartTitle.AddOrdered(CreateChartText(chartTitle, title), CTitleView.Order);
            if (chartTitle.Child(C, "layout").IsNone) {
                chartTitle.AddOrdered(chartTitle.New(C, "layout"), CTitleView.Order);
            }
            chartTitle.RemoveChildren(C, "overlay");
            chartTitle.AddOrdered(chartTitle.NewVal(C, "overlay", false), CTitleView.Order);

            if (existingTitle.IsNone) {
                OfficeOpenXmlChartThemeWriter.ApplyTitle(chartTitle, OfficeTheme.Default);
                chart.InsertBefore(chartTitle, chart.FirstChild);
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets the chart title text style.
        /// </summary>
        public ExcelChart SetTitleTextStyle(OfficeLength? fontSize = null, bool? bold = null, bool? italic = null,
            OfficeColor? color = null, string? fontFamily = null) {
            ValidateDataLabelTextStyle(fontSize, color, fontFamily);

            OfficeXmlNode chart = ChartNode();
            OfficeXmlNode title = chart.Child(C, "title");
            if (title.IsNone) {
                return this;
            }

            OfficeXmlNode chartText = title.Child(C, "tx");
            if (chartText.IsNone) {
                return this;
            }

            ApplyTextStyle(EnsureChartTextRunProperties(chartText), fontSize, bold, italic, color, fontFamily);
            Save();
            return this;
        }

        /// <summary>
        /// Removes the chart title.
        /// </summary>
        public ExcelChart ClearTitle() {
            OfficeXmlNode chart = ChartNode();
            chart.RemoveChild(C, "title");
            chart.SetOrdered(chart.NewVal(C, "autoTitleDeleted", true), CChartView.Order);
            Save();
            return this;
        }

        /// <summary>
        /// Shows the legend and updates only the supplied position or overlay state.
        /// </summary>
        public ExcelChart SetLegend(OfficeChartLegendPosition? position = null, bool? overlay = null) {
            OfficeXmlNode chart = ChartNode();
            OfficeXmlNode existingLegend = chart.Child(C, "legend");
            OfficeXmlNode legend = existingLegend.Exists ? existingLegend : chart.New(C, "legend");
            if (position.HasValue) {
                string token = position.Value.ToOpenXmlToken();
                OfficeXmlNode legendPosition = legend.Child(C, "legendPos");
                if (legendPosition.Exists) legendPosition.SetVal(token);
                else legend.AddOrdered(legend.NewVal(C, "legendPos", token), CLegendView.Order);
            }
            if (legend.Child(C, "layout").IsNone) {
                legend.AddOrdered(legend.New(C, "layout"), CLegendView.Order);
            }
            if (overlay.HasValue) {
                legend.RemoveChildren(C, "overlay");
                legend.AddOrdered(legend.NewVal(C, "overlay", overlay.Value), CLegendView.Order);
            }

            if (existingLegend.IsNone) {
                OfficeXmlNode plotArea = chart.Child(C, "plotArea");
                if (plotArea.Exists) {
                    plotArea.InsertAfterSelf(legend);
                } else {
                    chart.Append(legend);
                }
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets the legend text style.
        /// </summary>
        public ExcelChart SetLegendTextStyle(OfficeLength? fontSize = null, bool? bold = null, bool? italic = null,
            OfficeColor? color = null, string? fontFamily = null) {
            ValidateDataLabelTextStyle(fontSize, color, fontFamily);

            OfficeXmlNode legend = ChartNode().Child(C, "legend");
            if (legend.IsNone) {
                return this;
            }

            ApplyTextStyle(EnsureTextPropertiesRunProperties(legend), fontSize, bold, italic, color, fontFamily);
            Save();
            return this;
        }

        /// <summary>
        /// Hides the chart legend.
        /// </summary>
        public ExcelChart HideLegend() {
            ChartNode().RemoveChild(C, "legend");
            Save();
            return this;
        }

        /// <summary>
        /// Applies a built-in chart style/color preset.
        /// </summary>
        public ExcelChart ApplyStylePreset(int styleId = 251, int colorStyleId = 10) {
            _document.EnsureWorkbookThemeAndStyles();
            ExcelChartUtils.ApplyChartStyle(ChartPartRef(), styleId, colorStyleId);
            Save();
            return this;
        }

        /// <summary>
        /// Applies a chart style/color preset.
        /// </summary>
        public ExcelChart ApplyStylePreset(ExcelChartStylePreset preset) {
            ArgumentNullException.ThrowIfNull(preset);
            _document.EnsureWorkbookThemeAndStyles();
            ExcelChartUtils.ApplyChartStyle(ChartPartRef(), preset);
            Save();
            return this;
        }

        private static string? GetChartTitleText(OfficeXmlNode chart) {
            OfficeXmlNode chartText = chart.Child(C, "title").Child(C, "tx");
            if (chartText.IsNone) {
                return null;
            }

            string text = string.Concat(chartText.Descendants(A, "t").Select(item => item.InnerText));
            return string.IsNullOrWhiteSpace(text) ? null : text.Trim();
        }

        /// <summary>The <c>xdr:from</c> marker of the frame's one-cell or two-cell anchor, or none.</summary>
        private OfficeXmlNode GetFromMarker() {
            OfficeXmlNode marker = FrameNode.Ancestor(Xdr, "oneCellAnchor").Child(Xdr, "from");
            return marker.Exists ? marker : FrameNode.Ancestor(Xdr, "twoCellAnchor").Child(Xdr, "from");
        }

        private int GetAnchorRow() => ParseOneBasedMarker(GetFromMarker().ChildText(Xdr, "row"), 1048576);

        private int GetAnchorColumn() => ParseOneBasedMarker(GetFromMarker().ChildText(Xdr, "col"), 16384);

        private OfficeLength GetAnchorOffset(bool horizontal) =>
            OfficeLength.Emus(ParseEmuOffset(GetFromMarker().ChildText(Xdr, horizontal ? "colOff" : "rowOff")));

        private (OfficeLength Width, OfficeLength Height) GetAnchorSize(ExcelWorksheetGeometryIndex? geometry) {
            OfficeXmlNode twoCell = FrameNode.Ancestor(Xdr, "twoCellAnchor");
            if (twoCell.Child(Xdr, "from").Exists && twoCell.Child(Xdr, "to").Exists) {
                ExcelPartRef worksheet = DrawingsPartRef.ParentWorksheet;
                ExcelWorksheetGeometryIndex index = geometry ?? ExcelPartRef.CreateGeometryIndex(worksheet);
                return (GetTwoCellLength(twoCell, index, true), GetTwoCellLength(twoCell, index, false));
            }
            OfficeXmlNode extent = FrameNode.Ancestor(Xdr, "absoluteAnchor").Child(Xdr, "ext");
            if (extent.IsNone) extent = FrameNode.Ancestor(Xdr, "oneCellAnchor").Child(Xdr, "ext");
            return (extent.Int64Attribute("cx") is long cx && cx > 0 ? OfficeLength.Emus(cx) : OfficeLength.Pixels(480),
                extent.Int64Attribute("cy") is long cy && cy > 0 ? OfficeLength.Emus(cy) : OfficeLength.Pixels(320));
        }

        private static OfficeLength GetTwoCellLength(OfficeXmlNode anchor, ExcelWorksheetGeometryIndex geometry, bool horizontal) {
            OfficeXmlNode start = anchor.Child(Xdr, "from");
            OfficeXmlNode end = anchor.Child(Xdr, "to");
            if (start.IsNone || end.IsNone) return OfficeLength.Pixels(horizontal ? 480 : 320);
            int from = ParseZeroBasedMarker(start.ChildText(Xdr, horizontal ? "col" : "row"));
            int to = ParseZeroBasedMarker(end.ChildText(Xdr, horizontal ? "col" : "row"));
            if (!ExcelImageExportLimits.IsValidMarkerSpan(from, to, horizontal ? 16384 : 1048576))
                return OfficeLength.Pixels(horizontal ? 480 : 320);
            decimal emus = 0;
            for (int index = from; index < to; index++)
                emus += horizontal ? geometry.GetColumnLength(index + 1).ToEmus() : geometry.GetRowLength(index + 1).ToEmus();
            emus += (decimal)ParseEmuOffset(end.ChildText(Xdr, horizontal ? "colOff" : "rowOff"))
                - ParseEmuOffset(start.ChildText(Xdr, horizontal ? "colOff" : "rowOff"));
            return OfficeLength.Emus(checked((long)Math.Max(1m, emus)));
        }

        private (int WidthPixels, int HeightPixels) GetAnchorSizePixels(ExcelWorksheetGeometryIndex? geometry = null) {
            OfficeXmlNode twoCellAnchor = FrameNode.Ancestor(Xdr, "twoCellAnchor");
            if (twoCellAnchor.Child(Xdr, "from").Exists && twoCellAnchor.Child(Xdr, "to").Exists) {
                ExcelPartRef worksheetPart = DrawingsPartRef.ParentWorksheet;
                ExcelWorksheetGeometryIndex effectiveGeometry = geometry ?? ExcelPartRef.CreateGeometryIndex(worksheetPart);
                bool hasWidth = TryGetTwoCellAnchorSizePixels(twoCellAnchor, worksheetPart, effectiveGeometry, horizontal: true, out int widthPixels);
                bool hasHeight = TryGetTwoCellAnchorSizePixels(twoCellAnchor, worksheetPart, effectiveGeometry, horizontal: false, out int heightPixels);
                return (hasWidth ? widthPixels : 480, hasHeight ? heightPixels : 320);
            }

            OfficeXmlNode absoluteExtent = FrameNode.Ancestor(Xdr, "absoluteAnchor").Child(Xdr, "ext");
            if (absoluteExtent.Exists) {
                return (
                    EmuToPixels(absoluteExtent.Int64Attribute("cx"), 480),
                    EmuToPixels(absoluteExtent.Int64Attribute("cy"), 320));
            }

            OfficeXmlNode extent = FrameNode.Ancestor(Xdr, "oneCellAnchor").Child(Xdr, "ext");
            return (EmuToPixels(extent.Int64Attribute("cx"), 480), EmuToPixels(extent.Int64Attribute("cy"), 320));
        }

        private int GetDrawingOrder() {
            OfficeXmlNode anchor = FrameNode.Ancestor(Xdr, "oneCellAnchor");
            if (anchor.IsNone) anchor = FrameNode.Ancestor(Xdr, "twoCellAnchor");
            if (anchor.IsNone) anchor = FrameNode.Ancestor(Xdr, "absoluteAnchor");
            OfficeXmlNode worksheetDrawing = DrawingsPartRef.Root;
            if (anchor.IsNone || worksheetDrawing.IsNone) {
                return 0;
            }

            int index = 0;
            for (OfficeXmlNode child = worksheetDrawing.FirstChild; !child.IsNone; child = child.NextSibling) {
                if (child == anchor) {
                    return index;
                }
                index++;
            }

            return 0;
        }

        private static bool TryGetTwoCellAnchorSizePixels(OfficeXmlNode anchor, ExcelPartRef worksheetPart, ExcelWorksheetGeometryIndex geometry, bool horizontal, out int pixels) {
            pixels = 0;
            OfficeXmlNode fromMarker = anchor.Child(Xdr, "from");
            OfficeXmlNode toMarker = anchor.Child(Xdr, "to");
            if (fromMarker.IsNone || toMarker.IsNone) {
                return false;
            }

            int from = horizontal ? ParseZeroBasedMarker(fromMarker.ChildText(Xdr, "col")) : ParseZeroBasedMarker(fromMarker.ChildText(Xdr, "row"));
            int to = horizontal ? ParseZeroBasedMarker(toMarker.ChildText(Xdr, "col")) : ParseZeroBasedMarker(toMarker.ChildText(Xdr, "row"));
            int maximumMarker = horizontal ? 16384 : 1048576;
            if (!ExcelImageExportLimits.IsValidMarkerSpan(from, to, maximumMarker)) {
                return false;
            }

            long fromOffset = ParseEmuOffset(fromMarker.ChildText(Xdr, horizontal ? "colOff" : "rowOff"));
            long toOffset = ParseEmuOffset(toMarker.ChildText(Xdr, horizontal ? "colOff" : "rowOff"));
            double maximumDigitWidth = ExcelPartRef.GetMaximumDigitWidth(worksheetPart);
            long basePixels = 0L;
            for (int index = from; index < to; index++) {
                basePixels += horizontal
                    ? geometry.GetColumnWidthPixels(index + 1, maximumDigitWidth)
                    : geometry.GetRowHeightPixels(index + 1);
            }

            int offsetPixels = ExcelImageExportLimits.EmuOffsetDifferencePixels(fromOffset, toOffset);
            long totalPixels = Math.Max(1L, basePixels + offsetPixels);
            pixels = ExcelImageExportLimits.ClampExtentPixels((int)Math.Min(int.MaxValue, totalPixels));
            return pixels > 0;
        }

        private static int ParseOneBasedMarker(string? value, int maximum) {
            return int.TryParse(value, out int zeroBased) && zeroBased >= 0 && zeroBased < maximum ? zeroBased + 1 : 1;
        }

        private static int ParseZeroBasedMarker(string? value) {
            return int.TryParse(value, out int zeroBased) && zeroBased >= 0 ? zeroBased : 0;
        }

        private static long ParseEmuOffset(string? value) {
            return long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long emu) ? emu : 0L;
        }

        private static int EmuOffsetToPixels(long emu) {
            double pixels = Math.Round(emu / 9525D);
            if (pixels <= -int.MaxValue) return -int.MaxValue;
            if (pixels >= int.MaxValue) return int.MaxValue;
            return ExcelImageExportLimits.ClampOffsetPixels((int)pixels);
        }

        private static int EmuToPixels(long? emu, int fallback) {
            if (!emu.HasValue || emu.Value <= 0) {
                return fallback;
            }

            double pixels = Math.Round(emu.Value / 9525D);
            return ExcelImageExportLimits.ClampExtentPixels(pixels >= int.MaxValue ? int.MaxValue : Math.Max(1, (int)pixels));
        }

        /// <summary>
        /// Records a chart edit. The chart part keeps its loaded XML tree, and every package save writes that tree,
        /// so serializing it here as well would build the chart XML a second time for each edit.
        /// </summary>
        private void Save() => MarkChartDirty();

        /// <summary>
        /// Chart and drawing edits leave worksheet cells alone, so rows waiting for the direct worksheet writer stay there.
        /// </summary>
        private void MarkChartDirty() {
            using (_document.PreserveDirectDataSetFastSaveStateDuringDirtyMarks()) {
                _document.MarkPackageDirty();
            }
        }
    }
}
