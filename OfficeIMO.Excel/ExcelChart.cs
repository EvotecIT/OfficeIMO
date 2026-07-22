using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.Excel.Utilities;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Represents a chart on a worksheet.
    /// </summary>
    public sealed partial class ExcelChart {
        private readonly Xdr.GraphicFrame _frame;
        private readonly DrawingsPart _drawingsPart;
        private readonly ExcelDocument _document;
        private readonly string _sheetName;
        private ExcelChartDataRange? _dataRange;

        internal ExcelChart(Xdr.GraphicFrame frame, DrawingsPart drawingsPart, ExcelSheet sheet, ExcelChartDataRange? dataRange = null) {
            _frame = frame ?? throw new ArgumentNullException(nameof(frame));
            _drawingsPart = drawingsPart ?? throw new ArgumentNullException(nameof(drawingsPart));
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            _document = sheet.Document;
            _sheetName = sheet.Name;
            _dataRange = dataRange;
        }

        /// <summary>
        /// Gets or sets the chart name (non-visual drawing name).
        /// </summary>
        public string Name {
            get => _frame.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value ?? string.Empty;
            set {
                var props = _frame.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties;
                if (props != null) {
                    props.Name = value ?? string.Empty;
                }
            }
        }

        /// <summary>
        /// Gets the chart data range when it is known.
        /// </summary>
        public ExcelChartDataRange? DataRange => _dataRange;

        /// <summary>
        /// Gets this chart anchor's zero-based order in the worksheet drawing layer.
        /// </summary>
        public int DrawingOrder => GetDrawingOrder();

        internal bool TryGetAbsoluteAnchorBounds(out int xPixels, out int yPixels, out int widthPixels, out int heightPixels) {
            xPixels = 0;
            yPixels = 0;
            widthPixels = 0;
            heightPixels = 0;
            Xdr.AbsoluteAnchor? absoluteAnchor = _frame.Ancestors<Xdr.AbsoluteAnchor>().FirstOrDefault();
            if (absoluteAnchor == null) {
                return false;
            }

            xPixels = ExcelImageExportLimits.ClampAbsoluteOffsetPixels(EmuOffsetToPixels(absoluteAnchor.Position?.X?.Value ?? 0L));
            yPixels = ExcelImageExportLimits.ClampAbsoluteOffsetPixels(EmuOffsetToPixels(absoluteAnchor.Position?.Y?.Value ?? 0L));
            widthPixels = ExcelImageExportLimits.ClampExtentPixels(EmuToPixels(absoluteAnchor.Extent?.Cx?.Value, GetAnchorWidthPixels()));
            heightPixels = ExcelImageExportLimits.ClampExtentPixels(EmuToPixels(absoluteAnchor.Extent?.Cy?.Value, GetAnchorHeightPixels()));
            return widthPixels > 0 && heightPixels > 0;
        }

        /// <summary>
        /// Gets the detected chart type.
        /// </summary>
        public ExcelChartType ChartType {
            get {
                C.PlotArea? plotArea = GetChart().GetFirstChild<C.PlotArea>();
                return plotArea == null ? ExcelChartType.ColumnClustered : ExcelChartUtils.InferChartType(plotArea);
            }
        }

        /// <summary>
        /// Gets the chart title text when present.
        /// </summary>
        public string? Title => GetChartTitleText(GetChart());

        /// <summary>
        /// Tries to read the chart data from the chart's source range.
        /// </summary>
        public bool TryGetData(out ExcelChartData data) {
            try {
                ChartPart chartPart = GetChartPart();
                ExcelChartDataRange? range = _dataRange ?? ExcelChartUtils.TryExtractDataRange(chartPart);
                ExcelSheet sheet = range == null ? _document[_sheetName] : _document[range.SheetName];
                ExcelChartData? chartData = range == null
                    ? ExcelChartUtils.TryReadChartData(chartPart, sheet)
                    : ExcelChartUtils.TryReadChartData(sheet, range);
                if (chartData == null) {
                    data = null!;
                    return false;
                }

                chartData = ExcelChartUtils.ApplyChartSeriesTypes(chartPart, chartData, ChartType);
                chartData = ExcelChartUtils.ApplyScatterSeriesXValues(chartPart, chartData, sheet);
                chartData = ApplyImageExportSeriesStyles(chartPart, chartData, _document.WorkbookPartRoot);
                if (range != null) {
                    _dataRange = range;
                }
                data = chartData;
                return true;
            } catch {
                data = null!;
                return false;
            }
        }

        /// <summary>
        /// Tries to create a dependency-free snapshot for rendering/export consumers.
        /// </summary>
        public bool TryGetSnapshot(out ExcelChartSnapshot snapshot) {
            if (!TryGetData(out ExcelChartData data)) {
                snapshot = null!;
                return false;
            }

            snapshot = new ExcelChartSnapshot(
                Name,
                Title,
                ChartType,
                data,
                GetAnchorRow(),
                GetAnchorColumn(),
                GetAnchorOffsetXPixels(),
                GetAnchorOffsetYPixels(),
                GetAnchorWidthPixels(),
                GetAnchorHeightPixels(),
                CreateImageExportStyle(),
                CreateImageExportLayout(),
                CreateImageExportDiagnostics());
            return true;
        }

        /// <summary>
        /// Gets whether the chart declares a pivot table source.
        /// </summary>
        public bool IsPivotChart => !string.IsNullOrWhiteSpace(PivotTableName);

        /// <summary>
        /// Gets the pivot table name used as this chart's pivot source, if present.
        /// </summary>
        public string? PivotTableName {
            get {
                ChartPart chartPart = GetChartPart();
                return chartPart.ChartSpace?
                    .GetFirstChild<C.PivotSource>()?
                    .GetFirstChild<C.PivotTableName>()?
                    .Text;
            }
        }

        /// <summary>
        /// Marks the chart as sourced from a pivot table.
        /// </summary>
        /// <param name="pivotTableName">Pivot table name to assign as the chart's pivot source.</param>
        /// <param name="formatId">Pivot chart format id.</param>
        public ExcelChart SetPivotSource(string pivotTableName, uint formatId = 0U) {
            if (string.IsNullOrWhiteSpace(pivotTableName)) {
                throw new ArgumentNullException(nameof(pivotTableName));
            }

            ChartPart chartPart = GetChartPart();
            C.ChartSpace chartSpace = chartPart.ChartSpace ?? throw new InvalidOperationException("Chart space not found in chart part.");
            chartSpace.RemoveAllChildren<C.PivotSource>();

            var pivotSource = new C.PivotSource(
                new C.PivotTableName { Text = pivotTableName.Trim() },
                new C.FormatId { Val = formatId });

            C.Chart? chart = chartSpace.GetFirstChild<C.Chart>();
            if (chart != null) {
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
            if (data == null) {
                throw new ArgumentNullException(nameof(data));
            }

            var chartPart = GetChartPart();
            ExcelChartDataRange? resolved = dataRange ?? _dataRange ?? ExcelChartUtils.TryExtractDataRange(chartPart);
            if (resolved == null) {
                throw new InvalidOperationException("Chart data range could not be resolved. Provide a data range explicitly.");
            }

            resolved = resolved.WithSize(data.Categories.Count, data.Series.Count);

            if (writeToSheet) {
                var targetSheet = _document[resolved.SheetName];
                bool numericCategories = chartPart.ChartSpace?
                    .GetFirstChild<C.Chart>()?
                    .GetFirstChild<C.PlotArea>()?
                    .GetFirstChild<C.ScatterChart>() != null;
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
            if (title == null) {
                throw new ArgumentNullException(nameof(title));
            }

            C.Chart chart = GetChart();
            chart.AutoTitleDeleted = new C.AutoTitleDeleted { Val = false };

            C.Title chartTitle = chart.GetFirstChild<C.Title>() ?? new C.Title();
            chartTitle.RemoveAllChildren<C.ChartText>();
            chartTitle.Append(CreateChartText(title));
            if (chartTitle.GetFirstChild<C.Layout>() == null) {
                chartTitle.Append(new C.Layout());
            }
            chartTitle.RemoveAllChildren<C.Overlay>();
            chartTitle.Append(new C.Overlay { Val = false });

            if (chart.GetFirstChild<C.Title>() == null) {
                chart.InsertAt(chartTitle, 0);
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets the chart title text style.
        /// </summary>
        public ExcelChart SetTitleTextStyle(double? fontSizePoints = null, bool? bold = null, bool? italic = null,
            string? color = null, string? fontName = null) {
            ValidateDataLabelTextStyle(fontSizePoints, color, fontName);

            C.Chart chart = GetChart();
            C.Title? title = chart.GetFirstChild<C.Title>();
            if (title == null) {
                return this;
            }

            C.ChartText? chartText = title.GetFirstChild<C.ChartText>();
            if (chartText == null) {
                return this;
            }

            ApplyTextStyle(EnsureChartTextRunProperties(chartText), fontSizePoints, bold, italic, color, fontName);
            Save();
            return this;
        }

        /// <summary>
        /// Removes the chart title.
        /// </summary>
        public ExcelChart ClearTitle() {
            C.Chart chart = GetChart();
            chart.GetFirstChild<C.Title>()?.Remove();
            chart.AutoTitleDeleted = new C.AutoTitleDeleted { Val = true };
            Save();
            return this;
        }

        /// <summary>
        /// Sets the legend position and visibility.
        /// </summary>
        public ExcelChart SetLegend(C.LegendPositionValues position, bool overlay = false) {
            C.Chart chart = GetChart();
            C.Legend legend = chart.GetFirstChild<C.Legend>() ?? new C.Legend();
            var legendPosition = legend.GetFirstChild<C.LegendPosition>() ?? new C.LegendPosition();
            legendPosition.Val = position;
            if (legendPosition.Parent == null) {
                legend.Append(legendPosition);
            }
            if (legend.GetFirstChild<C.Layout>() == null) {
                legend.Append(new C.Layout());
            }
            legend.RemoveAllChildren<C.Overlay>();
            legend.Append(new C.Overlay { Val = overlay });

            if (chart.GetFirstChild<C.Legend>() == null) {
                C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
                if (plotArea != null) {
                    chart.InsertAfter(legend, plotArea);
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
        public ExcelChart SetLegendTextStyle(double? fontSizePoints = null, bool? bold = null, bool? italic = null,
            string? color = null, string? fontName = null) {
            ValidateDataLabelTextStyle(fontSizePoints, color, fontName);

            C.Chart chart = GetChart();
            C.Legend? legend = chart.GetFirstChild<C.Legend>();
            if (legend == null) {
                return this;
            }

            ApplyTextStyle(EnsureTextPropertiesRunProperties(legend), fontSizePoints, bold, italic, color, fontName);
            Save();
            return this;
        }

        /// <summary>
        /// Hides the chart legend.
        /// </summary>
        public ExcelChart HideLegend() {
            C.Chart chart = GetChart();
            chart.GetFirstChild<C.Legend>()?.Remove();
            Save();
            return this;
        }

        /// <summary>
        /// Applies a built-in chart style/color preset.
        /// </summary>
        public ExcelChart ApplyStylePreset(int styleId = 251, int colorStyleId = 10) {
            _document.EnsureWorkbookThemeAndStyles();
            ExcelChartUtils.ApplyChartStyle(GetChartPart(), styleId, colorStyleId);
            Save();
            return this;
        }

        /// <summary>
        /// Applies a chart style/color preset.
        /// </summary>
        public ExcelChart ApplyStylePreset(ExcelChartStylePreset preset) {
            if (preset == null) {
                throw new ArgumentNullException(nameof(preset));
            }
            _document.EnsureWorkbookThemeAndStyles();
            ExcelChartUtils.ApplyChartStyle(GetChartPart(), preset);
            Save();
            return this;
        }

        private C.Chart GetChart() {
            ChartPart chartPart = GetChartPart();
            C.Chart? chart = chartPart.ChartSpace?.GetFirstChild<C.Chart>();
            if (chart == null) {
                throw new InvalidOperationException("Chart element not found in chart part.");
            }
            return chart;
        }

        private static string? GetChartTitleText(C.Chart chart) {
            C.Title? title = chart.GetFirstChild<C.Title>();
            if (title == null) {
                return null;
            }

            C.ChartText? chartText = title.GetFirstChild<C.ChartText>();
            if (chartText == null) {
                return null;
            }

            string text = string.Concat(chartText.Descendants<A.Text>().Select(item => item.Text));
            return string.IsNullOrWhiteSpace(text) ? null : text.Trim();
        }

        private int GetAnchorRow() {
            Xdr.FromMarker? marker = _frame.Ancestors<Xdr.OneCellAnchor>().FirstOrDefault()?.FromMarker
                ?? _frame.Ancestors<Xdr.TwoCellAnchor>().FirstOrDefault()?.FromMarker;
            return ParseOneBasedMarker(marker?.RowId?.Text, 1048576);
        }

        private int GetAnchorColumn() {
            Xdr.FromMarker? marker = _frame.Ancestors<Xdr.OneCellAnchor>().FirstOrDefault()?.FromMarker
                ?? _frame.Ancestors<Xdr.TwoCellAnchor>().FirstOrDefault()?.FromMarker;
            return ParseOneBasedMarker(marker?.ColumnId?.Text, 16384);
        }

        private int GetAnchorOffsetXPixels() {
            Xdr.FromMarker? marker = _frame.Ancestors<Xdr.OneCellAnchor>().FirstOrDefault()?.FromMarker
                ?? _frame.Ancestors<Xdr.TwoCellAnchor>().FirstOrDefault()?.FromMarker;
            return EmuOffsetToPixels(ParseEmuOffset(marker?.ColumnOffset?.Text));
        }

        private int GetAnchorOffsetYPixels() {
            Xdr.FromMarker? marker = _frame.Ancestors<Xdr.OneCellAnchor>().FirstOrDefault()?.FromMarker
                ?? _frame.Ancestors<Xdr.TwoCellAnchor>().FirstOrDefault()?.FromMarker;
            return EmuOffsetToPixels(ParseEmuOffset(marker?.RowOffset?.Text));
        }

        private int GetAnchorWidthPixels() {
            Xdr.TwoCellAnchor? twoCellAnchor = _frame.Ancestors<Xdr.TwoCellAnchor>().FirstOrDefault();
            WorksheetPart? worksheetPart = _drawingsPart.GetParentParts().OfType<WorksheetPart>().FirstOrDefault();
            if (TryGetTwoCellAnchorSizePixels(twoCellAnchor, worksheetPart, GetWorksheetGeometry(worksheetPart), horizontal: true, out int widthPixels)) {
                return widthPixels;
            }

            long? emu = _frame.Ancestors<Xdr.OneCellAnchor>().FirstOrDefault()?.Extent?.Cx?.Value;
            return EmuToPixels(emu, 480);
        }

        private int GetAnchorHeightPixels() {
            Xdr.TwoCellAnchor? twoCellAnchor = _frame.Ancestors<Xdr.TwoCellAnchor>().FirstOrDefault();
            WorksheetPart? worksheetPart = _drawingsPart.GetParentParts().OfType<WorksheetPart>().FirstOrDefault();
            if (TryGetTwoCellAnchorSizePixels(twoCellAnchor, worksheetPart, GetWorksheetGeometry(worksheetPart), horizontal: false, out int heightPixels)) {
                return heightPixels;
            }

            long? emu = _frame.Ancestors<Xdr.OneCellAnchor>().FirstOrDefault()?.Extent?.Cy?.Value;
            return EmuToPixels(emu, 320);
        }

        private int GetDrawingOrder() {
            OpenXmlElement? anchor = _frame.Ancestors<Xdr.OneCellAnchor>().FirstOrDefault()
                ?? (OpenXmlElement?)_frame.Ancestors<Xdr.TwoCellAnchor>().FirstOrDefault()
                ?? _frame.Ancestors<Xdr.AbsoluteAnchor>().FirstOrDefault();
            Xdr.WorksheetDrawing? worksheetDrawing = _drawingsPart.WorksheetDrawing;
            if (anchor == null || worksheetDrawing == null) {
                return 0;
            }

            OpenXmlElementList children = worksheetDrawing.ChildElements;
            for (int i = 0; i < children.Count; i++) {
                if (ReferenceEquals(children[i], anchor)) {
                    return i;
                }
            }

            return 0;
        }

        private static ExcelWorksheetGeometryIndex GetWorksheetGeometry(WorksheetPart? worksheetPart) =>
            ExcelWorksheetGeometryIndex.Create(worksheetPart);

        private static bool TryGetTwoCellAnchorSizePixels(Xdr.TwoCellAnchor? anchor, WorksheetPart? worksheetPart, ExcelWorksheetGeometryIndex geometry, bool horizontal, out int pixels) {
            pixels = 0;
            if (anchor?.FromMarker == null || anchor.ToMarker == null) {
                return false;
            }

            int from = horizontal ? ParseZeroBasedMarker(anchor.FromMarker.ColumnId?.Text) : ParseZeroBasedMarker(anchor.FromMarker.RowId?.Text);
            int to = horizontal ? ParseZeroBasedMarker(anchor.ToMarker.ColumnId?.Text) : ParseZeroBasedMarker(anchor.ToMarker.RowId?.Text);
            int maximumMarker = horizontal ? 16384 : 1048576;
            if (!ExcelImageExportLimits.IsValidMarkerSpan(from, to, maximumMarker)) {
                return false;
            }

            long fromOffset = ParseEmuOffset(horizontal ? anchor.FromMarker.ColumnOffset?.Text : anchor.FromMarker.RowOffset?.Text);
            long toOffset = ParseEmuOffset(horizontal ? anchor.ToMarker.ColumnOffset?.Text : anchor.ToMarker.RowOffset?.Text);
            double maximumDigitWidth = GetDefaultMaximumDigitWidth(worksheetPart);
            long basePixels = 0L;
            for (int index = from; index < to; index++) {
                basePixels += horizontal
                    ? geometry.GetColumnWidthPixels(index + 1, maximumDigitWidth)
                    : geometry.GetRowHeightPixels(index + 1);
            }

            int offsetPixels = EmuOffsetToPixels(toOffset - fromOffset);
            long totalPixels = Math.Max(1L, basePixels + offsetPixels);
            pixels = ExcelImageExportLimits.ClampExtentPixels((int)Math.Min(int.MaxValue, totalPixels));
            return pixels > 0;
        }

        private static double GetDefaultMaximumDigitWidth(WorksheetPart? worksheetPart) {
            const double fallbackMaximumDigitWidth = 7D;
            try {
                WorkbookPart? workbookPart = worksheetPart?.GetParentParts().OfType<WorkbookPart>().FirstOrDefault();
                DocumentFormat.OpenXml.Spreadsheet.Font? font = workbookPart?.WorkbookStylesPart?.Stylesheet?.Fonts?
                    .Elements<DocumentFormat.OpenXml.Spreadsheet.Font>()
                    .FirstOrDefault();
                string? fontName = font?.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.FontName>()?.Val?.Value;
                if (string.IsNullOrWhiteSpace(fontName)) {
                    return fallbackMaximumDigitWidth;
                }

                double fontSize = font!.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.FontSize>()?.Val?.Value ?? 11D;
                OfficeFontStyle fontStyle = OfficeFontStyle.Regular;
                if (font.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Bold>() != null) {
                    fontStyle |= OfficeFontStyle.Bold;
                }

                if (font.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Italic>() != null) {
                    fontStyle |= OfficeFontStyle.Italic;
                }

                if (font.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Underline>() != null) {
                    fontStyle |= OfficeFontStyle.Underline;
                }

                float maximumDigitWidth = ExcelTextMeasurer.Create(new OfficeFontInfo(fontName, fontSize, fontStyle)).DefaultStyle.MaximumDigitWidth;
                return maximumDigitWidth > 0.0001f ? maximumDigitWidth : fallbackMaximumDigitWidth;
            } catch {
                return fallbackMaximumDigitWidth;
            }
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

        private ChartPart GetChartPart() {
            C.ChartReference? chartReference = _frame.Graphic?.GraphicData?.GetFirstChild<C.ChartReference>();
            StringValue? relationshipId = chartReference?.Id;
            if (relationshipId == null) {
                throw new InvalidOperationException("Chart reference not found for the shape.");
            }

            string relId = relationshipId.Value ?? throw new InvalidOperationException("Chart relationship id is empty.");
            return (ChartPart)_drawingsPart.GetPartById(relId);
        }

        private void Save() {
            ChartPart chartPart = GetChartPart();
            chartPart.ChartSpace?.Save();
            _document.MarkPackageDirty();
        }
    }
}
