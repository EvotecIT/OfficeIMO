using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Represents a chart on a worksheet.
    /// </summary>
    public sealed class ExcelChart {
        private readonly Xdr.GraphicFrame _frame;
        private readonly DrawingsPart _drawingsPart;
        private readonly ExcelDocument _document;
        private ExcelChartDataRange? _dataRange;

        internal ExcelChart(Xdr.GraphicFrame frame, DrawingsPart drawingsPart, ExcelSheet sheet, ExcelChartDataRange? dataRange = null) {
            _frame = frame ?? throw new ArgumentNullException(nameof(frame));
            _drawingsPart = drawingsPart ?? throw new ArgumentNullException(nameof(drawingsPart));
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            _document = sheet.Document;
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
                targetSheet.WriteChartData(data, resolved.StartRow, resolved.StartColumn, includeHeaderRow: resolved.HasHeaderRow, numericCategories: numericCategories);
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
            ExcelChartUtils.ApplyChartStyle(GetChartPart(), preset);
            Save();
            return this;
        }

        /// <summary>
        /// Configures data labels for all supported chart series.
        /// </summary>
        public ExcelChart SetDataLabels(bool showValue = true, bool showCategoryName = false,
            bool showSeriesName = false, bool showLegendKey = false, bool showPercent = false) {
            return SetDataLabels(showValue, showCategoryName, showSeriesName, showLegendKey, showPercent, null, null, false);
        }

        /// <summary>
        /// Configures data labels for all supported chart series with optional formatting.
        /// </summary>
        public ExcelChart SetDataLabels(bool showValue, bool showCategoryName,
            bool showSeriesName, bool showLegendKey, bool showPercent,
            C.DataLabelPositionValues? position, string? numberFormat, bool sourceLinked = false) {
            if (numberFormat != null && string.IsNullOrWhiteSpace(numberFormat)) {
                throw new ArgumentException("Number format cannot be empty.", nameof(numberFormat));
            }
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            foreach (C.BarChart barChart in plotArea.Elements<C.BarChart>()) {
                ApplyDataLabels(barChart, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent,
                    position, numberFormat, sourceLinked);
            }

            foreach (C.LineChart lineChart in plotArea.Elements<C.LineChart>()) {
                ApplyDataLabels(lineChart, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent,
                    position, numberFormat, sourceLinked);
            }

            foreach (C.AreaChart areaChart in plotArea.Elements<C.AreaChart>()) {
                ApplyDataLabels(areaChart, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent,
                    position, numberFormat, sourceLinked);
            }

            foreach (C.PieChart pieChart in plotArea.Elements<C.PieChart>()) {
                ApplyDataLabels(pieChart, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent,
                    position, numberFormat, sourceLinked);
            }

            foreach (C.DoughnutChart doughnutChart in plotArea.Elements<C.DoughnutChart>()) {
                ApplyDataLabels(doughnutChart, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent,
                    position, numberFormat, sourceLinked);
            }

            foreach (C.ScatterChart scatterChart in plotArea.Elements<C.ScatterChart>()) {
                ApplyDataLabels(scatterChart, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent,
                    position, numberFormat, sourceLinked);
            }

            foreach (C.BubbleChart bubbleChart in plotArea.Elements<C.BubbleChart>()) {
                ApplyDataLabels(bubbleChart, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent,
                    position, numberFormat, sourceLinked);
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets data label text style for all chart labels.
        /// </summary>
        public ExcelChart SetDataLabelTextStyle(double? fontSizePoints = null, bool? bold = null, bool? italic = null,
            string? color = null, string? fontName = null) {
            ValidateDataLabelTextStyle(fontSizePoints, color, fontName);

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            foreach (C.BarChart barChart in plotArea.Elements<C.BarChart>()) {
                ApplyDataLabelTextStyle(EnsureDataLabels(barChart), fontSizePoints, bold, italic, color, fontName);
            }
            foreach (C.LineChart lineChart in plotArea.Elements<C.LineChart>()) {
                ApplyDataLabelTextStyle(EnsureDataLabels(lineChart), fontSizePoints, bold, italic, color, fontName);
            }
            foreach (C.AreaChart areaChart in plotArea.Elements<C.AreaChart>()) {
                ApplyDataLabelTextStyle(EnsureDataLabels(areaChart), fontSizePoints, bold, italic, color, fontName);
            }
            foreach (C.PieChart pieChart in plotArea.Elements<C.PieChart>()) {
                ApplyDataLabelTextStyle(EnsureDataLabels(pieChart), fontSizePoints, bold, italic, color, fontName);
            }
            foreach (C.DoughnutChart doughnutChart in plotArea.Elements<C.DoughnutChart>()) {
                ApplyDataLabelTextStyle(EnsureDataLabels(doughnutChart), fontSizePoints, bold, italic, color, fontName);
            }
            foreach (C.ScatterChart scatterChart in plotArea.Elements<C.ScatterChart>()) {
                ApplyDataLabelTextStyle(EnsureDataLabels(scatterChart), fontSizePoints, bold, italic, color, fontName);
            }
            foreach (C.BubbleChart bubbleChart in plotArea.Elements<C.BubbleChart>()) {
                ApplyDataLabelTextStyle(EnsureDataLabels(bubbleChart), fontSizePoints, bold, italic, color, fontName);
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets data label shape styling for all chart labels.
        /// </summary>
        public ExcelChart SetDataLabelShapeStyle(string? fillColor = null, string? lineColor = null, double? lineWidthPoints = null,
            bool noFill = false, bool noLine = false) {
            ValidateDataLabelShapeStyle(fillColor, lineColor, lineWidthPoints, noFill, noLine);

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            foreach (C.BarChart barChart in plotArea.Elements<C.BarChart>()) {
                ApplyDataLabelShapeStyle(EnsureDataLabels(barChart), fillColor, lineColor, lineWidthPoints, noFill, noLine);
            }
            foreach (C.LineChart lineChart in plotArea.Elements<C.LineChart>()) {
                ApplyDataLabelShapeStyle(EnsureDataLabels(lineChart), fillColor, lineColor, lineWidthPoints, noFill, noLine);
            }
            foreach (C.AreaChart areaChart in plotArea.Elements<C.AreaChart>()) {
                ApplyDataLabelShapeStyle(EnsureDataLabels(areaChart), fillColor, lineColor, lineWidthPoints, noFill, noLine);
            }
            foreach (C.PieChart pieChart in plotArea.Elements<C.PieChart>()) {
                ApplyDataLabelShapeStyle(EnsureDataLabels(pieChart), fillColor, lineColor, lineWidthPoints, noFill, noLine);
            }
            foreach (C.DoughnutChart doughnutChart in plotArea.Elements<C.DoughnutChart>()) {
                ApplyDataLabelShapeStyle(EnsureDataLabels(doughnutChart), fillColor, lineColor, lineWidthPoints, noFill, noLine);
            }
            foreach (C.ScatterChart scatterChart in plotArea.Elements<C.ScatterChart>()) {
                ApplyDataLabelShapeStyle(EnsureDataLabels(scatterChart), fillColor, lineColor, lineWidthPoints, noFill, noLine);
            }
            foreach (C.BubbleChart bubbleChart in plotArea.Elements<C.BubbleChart>()) {
                ApplyDataLabelShapeStyle(EnsureDataLabels(bubbleChart), fillColor, lineColor, lineWidthPoints, noFill, noLine);
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets data label shape styling for a series by index.
        /// </summary>
        public ExcelChart SetSeriesDataLabelShapeStyle(int seriesIndex, string? fillColor = null, string? lineColor = null,
            double? lineWidthPoints = null, bool noFill = false, bool noLine = false) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            ValidateDataLabelShapeStyle(fillColor, lineColor, lineWidthPoints, noFill, noLine);

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                ApplyDataLabelShapeStyle(EnsureDataLabels(series), fillColor, lineColor, lineWidthPoints, noFill, noLine);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets data label shape styling for a series by name.
        /// </summary>
        public ExcelChart SetSeriesDataLabelShapeStyle(string seriesName, string? fillColor = null, string? lineColor = null,
            double? lineWidthPoints = null, bool noFill = false, bool noLine = false, bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            ValidateDataLabelShapeStyle(fillColor, lineColor, lineWidthPoints, noFill, noLine);

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                ApplyDataLabelShapeStyle(EnsureDataLabels(series), fillColor, lineColor, lineWidthPoints, noFill, noLine);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Configures data label leader lines for all chart labels.
        /// </summary>
        public ExcelChart SetDataLabelLeaderLines(bool showLeaderLines = true, string? lineColor = null, double? lineWidthPoints = null) {
            ValidateDataLabelLeaderLines(lineColor, lineWidthPoints);

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            foreach (C.BarChart barChart in plotArea.Elements<C.BarChart>()) {
                ApplyDataLabelLeaderLines(EnsureDataLabels(barChart), showLeaderLines, lineColor, lineWidthPoints);
            }
            foreach (C.LineChart lineChart in plotArea.Elements<C.LineChart>()) {
                ApplyDataLabelLeaderLines(EnsureDataLabels(lineChart), showLeaderLines, lineColor, lineWidthPoints);
            }
            foreach (C.AreaChart areaChart in plotArea.Elements<C.AreaChart>()) {
                ApplyDataLabelLeaderLines(EnsureDataLabels(areaChart), showLeaderLines, lineColor, lineWidthPoints);
            }
            foreach (C.PieChart pieChart in plotArea.Elements<C.PieChart>()) {
                ApplyDataLabelLeaderLines(EnsureDataLabels(pieChart), showLeaderLines, lineColor, lineWidthPoints);
            }
            foreach (C.DoughnutChart doughnutChart in plotArea.Elements<C.DoughnutChart>()) {
                ApplyDataLabelLeaderLines(EnsureDataLabels(doughnutChart), showLeaderLines, lineColor, lineWidthPoints);
            }
            foreach (C.ScatterChart scatterChart in plotArea.Elements<C.ScatterChart>()) {
                ApplyDataLabelLeaderLines(EnsureDataLabels(scatterChart), showLeaderLines, lineColor, lineWidthPoints);
            }
            foreach (C.BubbleChart bubbleChart in plotArea.Elements<C.BubbleChart>()) {
                ApplyDataLabelLeaderLines(EnsureDataLabels(bubbleChart), showLeaderLines, lineColor, lineWidthPoints);
            }

            Save();
            return this;
        }

        /// <summary>
        /// Configures data label leader lines for a series by index.
        /// </summary>
        public ExcelChart SetSeriesDataLabelLeaderLines(int seriesIndex, bool showLeaderLines = true, string? lineColor = null, double? lineWidthPoints = null) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            ValidateDataLabelLeaderLines(lineColor, lineWidthPoints);

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                ApplyDataLabelLeaderLines(EnsureDataLabels(series), showLeaderLines, lineColor, lineWidthPoints);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Configures data label leader lines for a series by name.
        /// </summary>
        public ExcelChart SetSeriesDataLabelLeaderLines(string seriesName, bool showLeaderLines = true, string? lineColor = null,
            double? lineWidthPoints = null, bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            ValidateDataLabelLeaderLines(lineColor, lineWidthPoints);

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                ApplyDataLabelLeaderLines(EnsureDataLabels(series), showLeaderLines, lineColor, lineWidthPoints);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Enables callout-style labels by positioning labels outside with leader lines.
        /// </summary>
        public ExcelChart SetDataLabelCallouts(bool enabled = true, C.DataLabelPositionValues? position = null,
            string? lineColor = null, double? lineWidthPoints = null) {
            var resolvedPosition = enabled ? position ?? C.DataLabelPositionValues.OutsideEnd : position;
            SetDataLabels(showValue: enabled, showCategoryName: false, showSeriesName: false, showLegendKey: false,
                showPercent: false, position: resolvedPosition, numberFormat: null, sourceLinked: false);
            return SetDataLabelLeaderLines(enabled, lineColor, lineWidthPoints);
        }

        /// <summary>
        /// Enables callout-style labels for a series by index.
        /// </summary>
        public ExcelChart SetSeriesDataLabelCallouts(int seriesIndex, bool enabled = true, C.DataLabelPositionValues? position = null,
            string? lineColor = null, double? lineWidthPoints = null) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            var resolvedPosition = enabled ? position ?? C.DataLabelPositionValues.OutsideEnd : position;
            SetSeriesDataLabels(seriesIndex, showValue: enabled, showCategoryName: false, showSeriesName: false,
                showLegendKey: false, showPercent: false, position: resolvedPosition, numberFormat: null, sourceLinked: false);
            return SetSeriesDataLabelLeaderLines(seriesIndex, enabled, lineColor, lineWidthPoints);
        }

        /// <summary>
        /// Enables callout-style labels for a series by name.
        /// </summary>
        public ExcelChart SetSeriesDataLabelCallouts(string seriesName, bool enabled = true, C.DataLabelPositionValues? position = null,
            string? lineColor = null, double? lineWidthPoints = null, bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            var resolvedPosition = enabled ? position ?? C.DataLabelPositionValues.OutsideEnd : position;
            SetSeriesDataLabels(seriesName, showValue: enabled, showCategoryName: false, showSeriesName: false,
                showLegendKey: false, showPercent: false, position: resolvedPosition, numberFormat: null, sourceLinked: false,
                ignoreCase: ignoreCase);
            return SetSeriesDataLabelLeaderLines(seriesName, enabled, lineColor, lineWidthPoints, ignoreCase);
        }

        /// <summary>
        /// Sets data label text style for a series by index.
        /// </summary>
        public ExcelChart SetSeriesDataLabelTextStyle(int seriesIndex, double? fontSizePoints = null, bool? bold = null,
            bool? italic = null, string? color = null, string? fontName = null) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            ValidateDataLabelTextStyle(fontSizePoints, color, fontName);

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                ApplyDataLabelTextStyle(EnsureDataLabels(series), fontSizePoints, bold, italic, color, fontName);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets data label text style for a series by name.
        /// </summary>
        public ExcelChart SetSeriesDataLabelTextStyle(string seriesName, double? fontSizePoints = null, bool? bold = null,
            bool? italic = null, string? color = null, string? fontName = null, bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            ValidateDataLabelTextStyle(fontSizePoints, color, fontName);

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                ApplyDataLabelTextStyle(EnsureDataLabels(series), fontSizePoints, bold, italic, color, fontName);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Applies a reusable data label template to a series by index.
        /// </summary>
        public ExcelChart SetSeriesDataLabelTemplate(int seriesIndex, ExcelChartDataLabelTemplate template) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            if (template == null) {
                throw new ArgumentNullException(nameof(template));
            }

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                ApplyDataLabelTemplate(series, template);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Applies a reusable data label template to a series by name.
        /// </summary>
        public ExcelChart SetSeriesDataLabelTemplate(string seriesName, ExcelChartDataLabelTemplate template,
            bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            if (template == null) {
                throw new ArgumentNullException(nameof(template));
            }

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                ApplyDataLabelTemplate(series, template);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Configures a single data label point by series index and point index.
        /// </summary>
        public ExcelChart SetSeriesDataLabelForPoint(int seriesIndex, int pointIndex, bool? showValue = null,
            bool? showCategoryName = null, bool? showSeriesName = null, bool? showLegendKey = null,
            bool? showPercent = null, bool? showBubbleSize = null,
            C.DataLabelPositionValues? position = null, string? numberFormat = null, bool sourceLinked = false) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            if (pointIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(pointIndex));
            }
            if (numberFormat != null && string.IsNullOrWhiteSpace(numberFormat)) {
                throw new ArgumentException("Number format cannot be empty.", nameof(numberFormat));
            }

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                C.DataLabel label = EnsureDataLabel(series, pointIndex);
                ApplyDataLabelOverrides(label, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent,
                    showBubbleSize, position, numberFormat, sourceLinked);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Configures a single data label point by series name and point index.
        /// </summary>
        public ExcelChart SetSeriesDataLabelForPoint(string seriesName, int pointIndex, bool? showValue = null,
            bool? showCategoryName = null, bool? showSeriesName = null, bool? showLegendKey = null,
            bool? showPercent = null, bool? showBubbleSize = null,
            C.DataLabelPositionValues? position = null, string? numberFormat = null, bool sourceLinked = false,
            bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            if (pointIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(pointIndex));
            }
            if (numberFormat != null && string.IsNullOrWhiteSpace(numberFormat)) {
                throw new ArgumentException("Number format cannot be empty.", nameof(numberFormat));
            }

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                C.DataLabel label = EnsureDataLabel(series, pointIndex);
                ApplyDataLabelOverrides(label, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent,
                    showBubbleSize, position, numberFormat, sourceLinked);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets data label text style for a specific point by series index.
        /// </summary>
        public ExcelChart SetSeriesDataLabelTextStyleForPoint(int seriesIndex, int pointIndex,
            double? fontSizePoints = null, bool? bold = null, bool? italic = null,
            string? color = null, string? fontName = null) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            if (pointIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(pointIndex));
            }
            ValidateDataLabelTextStyle(fontSizePoints, color, fontName);

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                C.DataLabel label = EnsureDataLabel(series, pointIndex);
                ApplyDataLabelTextStyle(label, fontSizePoints, bold, italic, color, fontName);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets data label text style for a specific point by series name.
        /// </summary>
        public ExcelChart SetSeriesDataLabelTextStyleForPoint(string seriesName, int pointIndex,
            double? fontSizePoints = null, bool? bold = null, bool? italic = null,
            string? color = null, string? fontName = null, bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            if (pointIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(pointIndex));
            }
            ValidateDataLabelTextStyle(fontSizePoints, color, fontName);

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                C.DataLabel label = EnsureDataLabel(series, pointIndex);
                ApplyDataLabelTextStyle(label, fontSizePoints, bold, italic, color, fontName);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets data label shape styling for a specific point by series index.
        /// </summary>
        public ExcelChart SetSeriesDataLabelShapeStyleForPoint(int seriesIndex, int pointIndex,
            string? fillColor = null, string? lineColor = null,
            double? lineWidthPoints = null, bool noFill = false, bool noLine = false) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            if (pointIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(pointIndex));
            }
            ValidateDataLabelShapeStyle(fillColor, lineColor, lineWidthPoints, noFill, noLine);

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                C.DataLabel label = EnsureDataLabel(series, pointIndex);
                ApplyDataLabelShapeStyle(label, fillColor, lineColor, lineWidthPoints, noFill, noLine);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets data label shape styling for a specific point by series name.
        /// </summary>
        public ExcelChart SetSeriesDataLabelShapeStyleForPoint(string seriesName, int pointIndex,
            string? fillColor = null, string? lineColor = null,
            double? lineWidthPoints = null, bool noFill = false, bool noLine = false,
            bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            if (pointIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(pointIndex));
            }
            ValidateDataLabelShapeStyle(fillColor, lineColor, lineWidthPoints, noFill, noLine);

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                C.DataLabel label = EnsureDataLabel(series, pointIndex);
                ApplyDataLabelShapeStyle(label, fillColor, lineColor, lineWidthPoints, noFill, noLine);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets the separator for all data labels.
        /// </summary>
        public ExcelChart SetDataLabelSeparator(string? separator) {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            foreach (C.BarChart barChart in plotArea.Elements<C.BarChart>()) {
                ApplyDataLabelSeparator(EnsureDataLabels(barChart), separator);
            }
            foreach (C.LineChart lineChart in plotArea.Elements<C.LineChart>()) {
                ApplyDataLabelSeparator(EnsureDataLabels(lineChart), separator);
            }
            foreach (C.AreaChart areaChart in plotArea.Elements<C.AreaChart>()) {
                ApplyDataLabelSeparator(EnsureDataLabels(areaChart), separator);
            }
            foreach (C.PieChart pieChart in plotArea.Elements<C.PieChart>()) {
                ApplyDataLabelSeparator(EnsureDataLabels(pieChart), separator);
            }
            foreach (C.DoughnutChart doughnutChart in plotArea.Elements<C.DoughnutChart>()) {
                ApplyDataLabelSeparator(EnsureDataLabels(doughnutChart), separator);
            }
            foreach (C.ScatterChart scatterChart in plotArea.Elements<C.ScatterChart>()) {
                ApplyDataLabelSeparator(EnsureDataLabels(scatterChart), separator);
            }
            foreach (C.BubbleChart bubbleChart in plotArea.Elements<C.BubbleChart>()) {
                ApplyDataLabelSeparator(EnsureDataLabels(bubbleChart), separator);
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets the separator for data labels in a series by index.
        /// </summary>
        public ExcelChart SetSeriesDataLabelSeparator(int seriesIndex, string? separator) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                ApplyDataLabelSeparator(EnsureDataLabels(series), separator);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets the separator for data labels in a series by name.
        /// </summary>
        public ExcelChart SetSeriesDataLabelSeparator(string seriesName, string? separator, bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                ApplyDataLabelSeparator(EnsureDataLabels(series), separator);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets the separator for a specific data label point by series index.
        /// </summary>
        public ExcelChart SetSeriesDataLabelSeparatorForPoint(int seriesIndex, int pointIndex, string? separator) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            if (pointIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(pointIndex));
            }

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                C.DataLabel label = EnsureDataLabel(series, pointIndex);
                ApplyDataLabelSeparator(label, separator);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets the separator for a specific data label point by series name.
        /// </summary>
        public ExcelChart SetSeriesDataLabelSeparatorForPoint(string seriesName, int pointIndex, string? separator,
            bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            if (pointIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(pointIndex));
            }

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                C.DataLabel label = EnsureDataLabel(series, pointIndex);
                ApplyDataLabelSeparator(label, separator);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Configures data labels for a single series by index.
        /// </summary>
        public ExcelChart SetSeriesDataLabels(int seriesIndex, bool showValue = true, bool showCategoryName = false,
            bool showSeriesName = false, bool showLegendKey = false, bool showPercent = false,
            C.DataLabelPositionValues? position = null, string? numberFormat = null, bool sourceLinked = false) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            if (numberFormat != null && string.IsNullOrWhiteSpace(numberFormat)) {
                throw new ArgumentException("Number format cannot be empty.", nameof(numberFormat));
            }

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                ApplyDataLabels(series, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent,
                    position, numberFormat, sourceLinked);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Configures data labels for a single series by name.
        /// </summary>
        public ExcelChart SetSeriesDataLabels(string seriesName, bool showValue = true, bool showCategoryName = false,
            bool showSeriesName = false, bool showLegendKey = false, bool showPercent = false,
            C.DataLabelPositionValues? position = null, string? numberFormat = null, bool sourceLinked = false,
            bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            if (numberFormat != null && string.IsNullOrWhiteSpace(numberFormat)) {
                throw new ArgumentException("Number format cannot be empty.", nameof(numberFormat));
            }

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                ApplyDataLabels(series, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent,
                    position, numberFormat, sourceLinked);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets the category axis title.
        /// </summary>
        public ExcelChart SetCategoryAxisTitle(string title) {
            return SetCategoryAxisTitle(title, ExcelChartAxisGroup.Primary);
        }

        /// <summary>
        /// Sets the category axis title for the selected axis group.
        /// </summary>
        public ExcelChart SetCategoryAxisTitle(string title, ExcelChartAxisGroup axisGroup) {
            return SetAxisTitle(title, axisGroup, AxisKind.Category);
        }

        /// <summary>
        /// Sets the value axis title.
        /// </summary>
        public ExcelChart SetValueAxisTitle(string title) {
            return SetValueAxisTitle(title, ExcelChartAxisGroup.Primary);
        }

        /// <summary>
        /// Sets the value axis title for the selected axis group.
        /// </summary>
        public ExcelChart SetValueAxisTitle(string title, ExcelChartAxisGroup axisGroup) {
            return SetAxisTitle(title, axisGroup, AxisKind.Value);
        }

        /// <summary>
        /// Sets the category axis title text style.
        /// </summary>
        public ExcelChart SetCategoryAxisTitleTextStyle(double? fontSizePoints = null, bool? bold = null,
            bool? italic = null, string? color = null, string? fontName = null,
            ExcelChartAxisGroup axisGroup = ExcelChartAxisGroup.Primary) {
            return SetAxisTitleTextStyle(axisGroup, AxisKind.Category, fontSizePoints, bold, italic, color, fontName);
        }

        /// <summary>
        /// Sets the value axis title text style.
        /// </summary>
        public ExcelChart SetValueAxisTitleTextStyle(double? fontSizePoints = null, bool? bold = null,
            bool? italic = null, string? color = null, string? fontName = null,
            ExcelChartAxisGroup axisGroup = ExcelChartAxisGroup.Primary) {
            return SetAxisTitleTextStyle(axisGroup, AxisKind.Value, fontSizePoints, bold, italic, color, fontName);
        }

        /// <summary>
        /// Sets category axis gridlines visibility and optional styling.
        /// </summary>
        public ExcelChart SetCategoryAxisGridlines(bool showMajor = true, bool showMinor = false,
            string? lineColor = null, double? lineWidthPoints = null,
            ExcelChartAxisGroup axisGroup = ExcelChartAxisGroup.Primary) {
            return SetAxisGridlines(axisGroup, AxisKind.Category, showMajor, showMinor, lineColor, lineWidthPoints);
        }

        /// <summary>
        /// Sets value axis gridlines visibility and optional styling.
        /// </summary>
        public ExcelChart SetValueAxisGridlines(bool showMajor = true, bool showMinor = false,
            string? lineColor = null, double? lineWidthPoints = null,
            ExcelChartAxisGroup axisGroup = ExcelChartAxisGroup.Primary) {
            return SetAxisGridlines(axisGroup, AxisKind.Value, showMajor, showMinor, lineColor, lineWidthPoints);
        }

        /// <summary>
        /// Sets the category axis label text style.
        /// </summary>
        public ExcelChart SetCategoryAxisLabelTextStyle(double? fontSizePoints = null, bool? bold = null,
            bool? italic = null, string? color = null, string? fontName = null,
            ExcelChartAxisGroup axisGroup = ExcelChartAxisGroup.Primary) {
            return SetAxisLabelTextStyle(axisGroup, AxisKind.Category, fontSizePoints, bold, italic, color, fontName);
        }

        /// <summary>
        /// Sets the value axis label text style.
        /// </summary>
        public ExcelChart SetValueAxisLabelTextStyle(double? fontSizePoints = null, bool? bold = null,
            bool? italic = null, string? color = null, string? fontName = null,
            ExcelChartAxisGroup axisGroup = ExcelChartAxisGroup.Primary) {
            return SetAxisLabelTextStyle(axisGroup, AxisKind.Value, fontSizePoints, bold, italic, color, fontName);
        }

        /// <summary>
        /// Sets the category axis label rotation in degrees (-90..90).
        /// </summary>
        public ExcelChart SetCategoryAxisLabelRotation(double rotationDegrees,
            ExcelChartAxisGroup axisGroup = ExcelChartAxisGroup.Primary) {
            return SetAxisLabelRotation(axisGroup, AxisKind.Category, rotationDegrees);
        }

        /// <summary>
        /// Sets the value axis label rotation in degrees (-90..90).
        /// </summary>
        public ExcelChart SetValueAxisLabelRotation(double rotationDegrees,
            ExcelChartAxisGroup axisGroup = ExcelChartAxisGroup.Primary) {
            return SetAxisLabelRotation(axisGroup, AxisKind.Value, rotationDegrees);
        }

        /// <summary>
        /// Sets the category axis tick label position.
        /// </summary>
        public ExcelChart SetCategoryAxisTickLabelPosition(C.TickLabelPositionValues position,
            ExcelChartAxisGroup axisGroup = ExcelChartAxisGroup.Primary) {
            return SetAxisTickLabelPosition(axisGroup, AxisKind.Category, position);
        }

        /// <summary>
        /// Sets the value axis tick label position.
        /// </summary>
        public ExcelChart SetValueAxisTickLabelPosition(C.TickLabelPositionValues position,
            ExcelChartAxisGroup axisGroup = ExcelChartAxisGroup.Primary) {
            return SetAxisTickLabelPosition(axisGroup, AxisKind.Value, position);
        }

        /// <summary>
        /// Sets how the value axis crosses between categories.
        /// </summary>
        public ExcelChart SetValueAxisCrossBetween(C.CrossBetweenValues between,
            ExcelChartAxisGroup axisGroup = ExcelChartAxisGroup.Primary) {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            C.ValueAxis? axis = ResolveValueAxis(plotArea, axisGroup);
            if (axis == null) {
                return this;
            }

            ReplaceChild(axis, new C.CrossBetween { Val = between });
            Save();
            return this;
        }

        /// <summary>
        /// Sets where the category axis crosses the value axis.
        /// </summary>
        public ExcelChart SetCategoryAxisCrossing(C.CrossesValues crosses, double? crossesAt = null,
            ExcelChartAxisGroup axisGroup = ExcelChartAxisGroup.Primary) {
            if (crossesAt != null && (double.IsNaN(crossesAt.Value) || double.IsInfinity(crossesAt.Value))) {
                throw new ArgumentOutOfRangeException(nameof(crossesAt));
            }

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            C.CategoryAxis? axis = ResolveCategoryAxis(plotArea, axisGroup);
            if (axis == null) {
                return this;
            }

            ApplyAxisCrossing(axis, crosses, crossesAt);
            Save();
            return this;
        }

        /// <summary>
        /// Sets where the value axis crosses the category axis.
        /// </summary>
        public ExcelChart SetValueAxisCrossing(C.CrossesValues crosses, double? crossesAt = null,
            ExcelChartAxisGroup axisGroup = ExcelChartAxisGroup.Primary) {
            if (crossesAt != null && (double.IsNaN(crossesAt.Value) || double.IsInfinity(crossesAt.Value))) {
                throw new ArgumentOutOfRangeException(nameof(crossesAt));
            }

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            C.ValueAxis? axis = ResolveValueAxis(plotArea, axisGroup);
            if (axis == null) {
                return this;
            }

            ValidateCrossesAtForAxis(axis, crossesAt);
            ApplyAxisCrossing(axis, crosses, crossesAt);
            Save();
            return this;
        }

        /// <summary>
        /// Sets where the scatter X-axis crosses the Y-axis.
        /// </summary>
        public ExcelChart SetScatterXAxisCrossing(C.CrossesValues crosses, double? crossesAt = null) {
            if (crossesAt != null && (double.IsNaN(crossesAt.Value) || double.IsInfinity(crossesAt.Value))) {
                throw new ArgumentOutOfRangeException(nameof(crossesAt));
            }

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            C.ValueAxis? axis = ResolveScatterXAxis(plotArea);
            if (axis == null) {
                return this;
            }

            ValidateCrossesAtForAxis(axis, crossesAt);
            ApplyAxisCrossing(axis, crosses, crossesAt);
            Save();
            return this;
        }

        /// <summary>
        /// Sets where the scatter Y-axis crosses the X-axis.
        /// </summary>
        public ExcelChart SetScatterYAxisCrossing(C.CrossesValues crosses, double? crossesAt = null) {
            if (crossesAt != null && (double.IsNaN(crossesAt.Value) || double.IsInfinity(crossesAt.Value))) {
                throw new ArgumentOutOfRangeException(nameof(crossesAt));
            }

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            C.ValueAxis? axis = ResolveScatterYAxis(plotArea);
            if (axis == null) {
                return this;
            }

            ValidateCrossesAtForAxis(axis, crossesAt);
            ApplyAxisCrossing(axis, crosses, crossesAt);
            Save();
            return this;
        }

        /// <summary>
        /// Sets display units for the value axis.
        /// </summary>
        public ExcelChart SetValueAxisDisplayUnits(C.BuiltInUnitValues unit, bool showLabel = true,
            ExcelChartAxisGroup axisGroup = ExcelChartAxisGroup.Primary) {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            C.ValueAxis? axis = ResolveValueAxis(plotArea, axisGroup);
            if (axis == null) {
                return this;
            }

            C.DisplayUnits displayUnits = axis.GetFirstChild<C.DisplayUnits>() ?? new C.DisplayUnits();
            displayUnits.RemoveAllChildren<C.CustomDisplayUnit>();
            displayUnits.RemoveAllChildren<C.BuiltInUnit>();
            displayUnits.Append(new C.BuiltInUnit { Val = unit });
            ApplyDisplayUnitsLabel(displayUnits, showLabel);
            if (displayUnits.Parent == null) {
                axis.Append(displayUnits);
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets display units for the value axis with custom label text.
        /// </summary>
        public ExcelChart SetValueAxisDisplayUnits(C.BuiltInUnitValues unit, string labelText, bool showLabel = true,
            ExcelChartAxisGroup axisGroup = ExcelChartAxisGroup.Primary) {
            if (string.IsNullOrWhiteSpace(labelText)) {
                throw new ArgumentException("Label text cannot be empty.", nameof(labelText));
            }

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            C.ValueAxis? axis = ResolveValueAxis(plotArea, axisGroup);
            if (axis == null) {
                return this;
            }

            C.DisplayUnits displayUnits = axis.GetFirstChild<C.DisplayUnits>() ?? new C.DisplayUnits();
            displayUnits.RemoveAllChildren<C.CustomDisplayUnit>();
            displayUnits.RemoveAllChildren<C.BuiltInUnit>();
            displayUnits.Append(new C.BuiltInUnit { Val = unit });
            ApplyDisplayUnitsLabel(displayUnits, showLabel, labelText);
            if (displayUnits.Parent == null) {
                axis.Append(displayUnits);
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets custom display units for the value axis.
        /// </summary>
        public ExcelChart SetValueAxisDisplayUnits(double customUnit, bool showLabel = true,
            ExcelChartAxisGroup axisGroup = ExcelChartAxisGroup.Primary) {
            if (customUnit <= 0 || double.IsNaN(customUnit) || double.IsInfinity(customUnit)) {
                throw new ArgumentOutOfRangeException(nameof(customUnit));
            }

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            C.ValueAxis? axis = ResolveValueAxis(plotArea, axisGroup);
            if (axis == null) {
                return this;
            }

            C.DisplayUnits displayUnits = axis.GetFirstChild<C.DisplayUnits>() ?? new C.DisplayUnits();
            displayUnits.RemoveAllChildren<C.BuiltInUnit>();
            displayUnits.RemoveAllChildren<C.CustomDisplayUnit>();
            displayUnits.Append(new C.CustomDisplayUnit { Val = customUnit });
            ApplyDisplayUnitsLabel(displayUnits, showLabel);
            if (displayUnits.Parent == null) {
                axis.Append(displayUnits);
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets custom display units for the value axis with custom label text.
        /// </summary>
        public ExcelChart SetValueAxisDisplayUnits(double customUnit, string labelText, bool showLabel = true,
            ExcelChartAxisGroup axisGroup = ExcelChartAxisGroup.Primary) {
            if (customUnit <= 0 || double.IsNaN(customUnit) || double.IsInfinity(customUnit)) {
                throw new ArgumentOutOfRangeException(nameof(customUnit));
            }
            if (string.IsNullOrWhiteSpace(labelText)) {
                throw new ArgumentException("Label text cannot be empty.", nameof(labelText));
            }

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            C.ValueAxis? axis = ResolveValueAxis(plotArea, axisGroup);
            if (axis == null) {
                return this;
            }

            C.DisplayUnits displayUnits = axis.GetFirstChild<C.DisplayUnits>() ?? new C.DisplayUnits();
            displayUnits.RemoveAllChildren<C.BuiltInUnit>();
            displayUnits.RemoveAllChildren<C.CustomDisplayUnit>();
            displayUnits.Append(new C.CustomDisplayUnit { Val = customUnit });
            ApplyDisplayUnitsLabel(displayUnits, showLabel, labelText);
            if (displayUnits.Parent == null) {
                axis.Append(displayUnits);
            }

            Save();
            return this;
        }

        /// <summary>
        /// Clears display units from the value axis.
        /// </summary>
        public ExcelChart ClearValueAxisDisplayUnits(ExcelChartAxisGroup axisGroup = ExcelChartAxisGroup.Primary) {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            C.ValueAxis? axis = ResolveValueAxis(plotArea, axisGroup);
            if (axis == null) {
                return this;
            }

            axis.GetFirstChild<C.DisplayUnits>()?.Remove();
            Save();
            return this;
        }

        /// <summary>
        /// Sets the category axis orientation (normal or reversed order).
        /// </summary>
        public ExcelChart SetCategoryAxisReverseOrder(bool reverseOrder = true,
            ExcelChartAxisGroup axisGroup = ExcelChartAxisGroup.Primary) {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            C.CategoryAxis? axis = ResolveCategoryAxis(plotArea, axisGroup);
            if (axis == null) {
                return this;
            }

            C.Scaling scaling = EnsureScaling(axis);
            ReplaceChild(scaling, new C.Orientation {
                Val = reverseOrder ? C.OrientationValues.MaxMin : C.OrientationValues.MinMax
            });
            Save();
            return this;
        }

        /// <summary>
        /// Sets scatter chart X-axis scale (value axis on the bottom).
        /// </summary>
        public ExcelChart SetScatterXAxisScale(double? minimum = null, double? maximum = null,
            double? majorUnit = null, double? minorUnit = null, bool? reverseOrder = null,
            bool? logScale = null, double? logBase = null) {
            ValidateAxisScale(minimum, maximum, majorUnit, minorUnit, logScale, logBase);

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            C.ValueAxis? axis = ResolveScatterXAxis(plotArea);
            if (axis == null) {
                return this;
            }

            ApplyAxisScale(axis, minimum, maximum, majorUnit, minorUnit, reverseOrder, logScale, logBase);
            Save();
            return this;
        }

        /// <summary>
        /// Sets scatter chart Y-axis scale (value axis on the left).
        /// </summary>
        public ExcelChart SetScatterYAxisScale(double? minimum = null, double? maximum = null,
            double? majorUnit = null, double? minorUnit = null, bool? reverseOrder = null,
            bool? logScale = null, double? logBase = null) {
            ValidateAxisScale(minimum, maximum, majorUnit, minorUnit, logScale, logBase);

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            C.ValueAxis? axis = ResolveScatterYAxis(plotArea);
            if (axis == null) {
                return this;
            }

            ApplyAxisScale(axis, minimum, maximum, majorUnit, minorUnit, reverseOrder, logScale, logBase);
            Save();
            return this;
        }

        /// <summary>
        /// Sets the category axis number format.
        /// </summary>
        public ExcelChart SetCategoryAxisNumberFormat(string formatCode, bool sourceLinked = false) {
            return SetCategoryAxisNumberFormat(formatCode, sourceLinked, ExcelChartAxisGroup.Primary);
        }

        /// <summary>
        /// Sets the category axis number format for the selected axis group.
        /// </summary>
        public ExcelChart SetCategoryAxisNumberFormat(string formatCode, bool sourceLinked, ExcelChartAxisGroup axisGroup) {
            return SetAxisNumberFormat(formatCode, sourceLinked, axisGroup, AxisKind.Category);
        }

        /// <summary>
        /// Sets the value axis number format.
        /// </summary>
        public ExcelChart SetValueAxisNumberFormat(string formatCode, bool sourceLinked = false) {
            return SetValueAxisNumberFormat(formatCode, sourceLinked, ExcelChartAxisGroup.Primary);
        }

        /// <summary>
        /// Sets the value axis number format for the selected axis group.
        /// </summary>
        public ExcelChart SetValueAxisNumberFormat(string formatCode, bool sourceLinked, ExcelChartAxisGroup axisGroup) {
            return SetAxisNumberFormat(formatCode, sourceLinked, axisGroup, AxisKind.Value);
        }

        /// <summary>
        /// Sets value axis scale parameters for the selected axis group.
        /// </summary>
        public ExcelChart SetValueAxisScale(double? minimum = null, double? maximum = null, double? majorUnit = null,
            double? minorUnit = null, double? logBase = null, bool? reverseOrder = null,
            ExcelChartAxisGroup axisGroup = ExcelChartAxisGroup.Primary, bool? logScale = null) {
            ValidateAxisScale(minimum, maximum, majorUnit, minorUnit, logScale, logBase);

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            C.ValueAxis? axis = ResolveValueAxis(plotArea, axisGroup);
            if (axis == null) {
                return this;
            }

            ApplyAxisScale(axis, minimum, maximum, majorUnit, minorUnit, reverseOrder, logScale, logBase);
            Save();
            return this;
        }

        private enum AxisKind {
            Category,
            Value
        }

        private ExcelChart SetAxisTitle(string title, ExcelChartAxisGroup axisGroup, AxisKind axisKind) {
            if (title == null) {
                throw new ArgumentNullException(nameof(title));
            }

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            OpenXmlCompositeElement? axis = axisKind == AxisKind.Category
                ? ResolveCategoryAxis(plotArea, axisGroup)
                : ResolveValueAxis(plotArea, axisGroup);

            if (axis == null && axisKind == AxisKind.Category && axisGroup == ExcelChartAxisGroup.Primary) {
                axis = ResolveScatterXAxis(plotArea);
            }

            if (axis == null) {
                return this;
            }

            axis.RemoveAllChildren<C.Title>();
            axis.Append(CreateAxisTitle(title));
            Save();
            return this;
        }

        private ExcelChart SetAxisTitleTextStyle(ExcelChartAxisGroup axisGroup, AxisKind axisKind,
            double? fontSizePoints, bool? bold, bool? italic, string? color, string? fontName) {
            ValidateDataLabelTextStyle(fontSizePoints, color, fontName);

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            OpenXmlCompositeElement? axis = axisKind == AxisKind.Category
                ? ResolveCategoryAxis(plotArea, axisGroup)
                : ResolveValueAxis(plotArea, axisGroup);

            if (axis == null && axisKind == AxisKind.Category && axisGroup == ExcelChartAxisGroup.Primary) {
                axis = ResolveScatterXAxis(plotArea);
            }

            if (axis == null) {
                return this;
            }

            C.Title? title = axis.GetFirstChild<C.Title>();
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

        private ExcelChart SetAxisLabelTextStyle(ExcelChartAxisGroup axisGroup, AxisKind axisKind,
            double? fontSizePoints, bool? bold, bool? italic, string? color, string? fontName) {
            ValidateDataLabelTextStyle(fontSizePoints, color, fontName);

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            OpenXmlCompositeElement? axis = axisKind == AxisKind.Category
                ? ResolveCategoryAxis(plotArea, axisGroup)
                : ResolveValueAxis(plotArea, axisGroup);

            if (axis == null && axisKind == AxisKind.Category && axisGroup == ExcelChartAxisGroup.Primary) {
                axis = ResolveScatterXAxis(plotArea);
            }

            if (axis == null) {
                return this;
            }

            ApplyTextStyle(EnsureTextPropertiesRunProperties(axis), fontSizePoints, bold, italic, color, fontName);
            Save();
            return this;
        }

        private ExcelChart SetAxisGridlines(ExcelChartAxisGroup axisGroup, AxisKind axisKind,
            bool showMajor, bool showMinor, string? lineColor, double? lineWidthPoints) {
            ValidateAxisGridlinesStyle(lineColor, lineWidthPoints);

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            OpenXmlCompositeElement? axis = axisKind == AxisKind.Category
                ? ResolveCategoryAxis(plotArea, axisGroup)
                : ResolveValueAxis(plotArea, axisGroup);

            if (axis == null && axisKind == AxisKind.Category && axisGroup == ExcelChartAxisGroup.Primary) {
                axis = ResolveScatterXAxis(plotArea);
            }

            if (axis == null) {
                return this;
            }

            ApplyGridlines(axis, showMajor, showMinor, lineColor, lineWidthPoints);
            Save();
            return this;
        }

        private ExcelChart SetAxisLabelRotation(ExcelChartAxisGroup axisGroup, AxisKind axisKind, double rotationDegrees) {
            if (double.IsNaN(rotationDegrees) || double.IsInfinity(rotationDegrees)) {
                throw new ArgumentOutOfRangeException(nameof(rotationDegrees));
            }
            if (rotationDegrees < -90 || rotationDegrees > 90) {
                throw new ArgumentOutOfRangeException(nameof(rotationDegrees), "Rotation must be between -90 and 90 degrees.");
            }

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            OpenXmlCompositeElement? axis = axisKind == AxisKind.Category
                ? ResolveCategoryAxis(plotArea, axisGroup)
                : ResolveValueAxis(plotArea, axisGroup);

            if (axis == null && axisKind == AxisKind.Category && axisGroup == ExcelChartAxisGroup.Primary) {
                axis = ResolveScatterXAxis(plotArea);
            }

            if (axis == null) {
                return this;
            }

            C.TextProperties textProps = axis.GetFirstChild<C.TextProperties>() ?? new C.TextProperties();
            A.BodyProperties body = textProps.GetFirstChild<A.BodyProperties>() ?? new A.BodyProperties();
            body.Rotation = (int)Math.Round(rotationDegrees * 60000d);
            if (body.Parent == null) {
                textProps.Append(body);
            }
            if (textProps.Parent == null) {
                axis.Append(textProps);
            }

            Save();
            return this;
        }

        private ExcelChart SetAxisTickLabelPosition(ExcelChartAxisGroup axisGroup, AxisKind axisKind,
            C.TickLabelPositionValues position) {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            OpenXmlCompositeElement? axis = axisKind == AxisKind.Category
                ? ResolveCategoryAxis(plotArea, axisGroup)
                : ResolveValueAxis(plotArea, axisGroup);

            if (axis == null && axisKind == AxisKind.Category && axisGroup == ExcelChartAxisGroup.Primary) {
                axis = ResolveScatterXAxis(plotArea);
            }

            if (axis == null) {
                return this;
            }

            ReplaceChild(axis, new C.TickLabelPosition { Val = position });
            Save();
            return this;
        }

        private ExcelChart SetAxisNumberFormat(string formatCode, bool sourceLinked,
            ExcelChartAxisGroup axisGroup, AxisKind axisKind) {
            if (string.IsNullOrWhiteSpace(formatCode)) {
                throw new ArgumentException("Format code cannot be null or empty.", nameof(formatCode));
            }

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            OpenXmlCompositeElement? axis = axisKind == AxisKind.Category
                ? ResolveCategoryAxis(plotArea, axisGroup)
                : ResolveValueAxis(plotArea, axisGroup);

            if (axis == null && axisKind == AxisKind.Category && axisGroup == ExcelChartAxisGroup.Primary) {
                axis = ResolveScatterXAxis(plotArea);
            }

            if (axis == null) {
                return this;
            }

            C.NumberingFormat format = axis.GetFirstChild<C.NumberingFormat>() ?? new C.NumberingFormat();
            format.FormatCode = formatCode;
            format.SourceLinked = sourceLinked;
            if (format.Parent == null) {
                axis.InsertAt(format, 0);
            }

            Save();
            return this;
        }

        private static C.CategoryAxis? ResolveCategoryAxis(C.PlotArea plotArea, ExcelChartAxisGroup axisGroup) {
            var axes = plotArea.Elements<C.CategoryAxis>().ToList();
            if (axes.Count == 0) {
                return null;
            }

            bool isBar = HasHorizontalBarChart(plotArea);
            C.AxisPositionValues primaryPosition = isBar ? C.AxisPositionValues.Left : C.AxisPositionValues.Bottom;
            C.AxisPositionValues secondaryPosition = isBar ? C.AxisPositionValues.Right : C.AxisPositionValues.Top;
            C.AxisPositionValues desired = axisGroup == ExcelChartAxisGroup.Primary ? primaryPosition : secondaryPosition;

            C.CategoryAxis? axis = axes.FirstOrDefault(ax => ax.AxisPosition?.Val?.Value == desired);
            if (axis != null) {
                return axis;
            }

            return axisGroup == ExcelChartAxisGroup.Primary
                ? axes.FirstOrDefault()
                : axes.Skip(1).FirstOrDefault() ?? axes.LastOrDefault();
        }

        private static C.ValueAxis? ResolveValueAxis(C.PlotArea plotArea, ExcelChartAxisGroup axisGroup) {
            var axes = plotArea.Elements<C.ValueAxis>().ToList();
            if (axes.Count == 0) {
                return null;
            }

            bool isBar = HasHorizontalBarChart(plotArea);
            C.AxisPositionValues primaryPosition = isBar ? C.AxisPositionValues.Bottom : C.AxisPositionValues.Left;
            C.AxisPositionValues secondaryPosition = isBar ? C.AxisPositionValues.Top : C.AxisPositionValues.Right;
            C.AxisPositionValues desired = axisGroup == ExcelChartAxisGroup.Primary ? primaryPosition : secondaryPosition;

            C.ValueAxis? axis = axes.FirstOrDefault(ax => ax.AxisPosition?.Val?.Value == desired);
            if (axis != null) {
                return axis;
            }

            return axisGroup == ExcelChartAxisGroup.Primary
                ? axes.FirstOrDefault()
                : axes.Skip(1).FirstOrDefault() ?? axes.LastOrDefault();
        }

        private static bool HasHorizontalBarChart(C.PlotArea plotArea) {
            return plotArea.Elements<C.BarChart>()
                .Select(chart => chart.GetFirstChild<C.BarDirection>()?.Val?.Value ?? C.BarDirectionValues.Column)
                .Any(direction => direction == C.BarDirectionValues.Bar);
        }

        private static C.ValueAxis? ResolveScatterXAxis(C.PlotArea plotArea) {
            if (plotArea.Elements<C.CategoryAxis>().Any()) {
                return null;
            }

            return plotArea.Elements<C.ValueAxis>()
                .FirstOrDefault(ax => ax.AxisPosition?.Val?.Value == C.AxisPositionValues.Bottom);
        }

        private static C.ValueAxis? ResolveScatterYAxis(C.PlotArea plotArea) {
            if (plotArea.Elements<C.CategoryAxis>().Any()) {
                return null;
            }

            return plotArea.Elements<C.ValueAxis>()
                .FirstOrDefault(ax => ax.AxisPosition?.Val?.Value == C.AxisPositionValues.Left);
        }

        /// <summary>
        /// Sets the fill color for a chart series by index.
        /// </summary>
        public ExcelChart SetSeriesFillColor(int seriesIndex, string color) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            if (string.IsNullOrWhiteSpace(color)) {
                throw new ArgumentException("Series color cannot be null or empty.", nameof(color));
            }

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                C.ChartShapeProperties props = EnsureChartShapeProperties(series);
                ApplySolidFill(props, NormalizeHexColor(color));
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets the fill color for a chart series by name.
        /// </summary>
        public ExcelChart SetSeriesFillColor(string seriesName, string color, bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            if (string.IsNullOrWhiteSpace(color)) {
                throw new ArgumentException("Series color cannot be null or empty.", nameof(color));
            }

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                C.ChartShapeProperties props = EnsureChartShapeProperties(series);
                ApplySolidFill(props, NormalizeHexColor(color));
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets the line color for a chart series by index.
        /// </summary>
        public ExcelChart SetSeriesLineColor(int seriesIndex, string color, double? widthPoints = null) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            if (string.IsNullOrWhiteSpace(color)) {
                throw new ArgumentException("Series color cannot be null or empty.", nameof(color));
            }

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                C.ChartShapeProperties props = EnsureChartShapeProperties(series);
                ApplyLine(props, NormalizeHexColor(color), widthPoints);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets the line color for a chart series by name.
        /// </summary>
        public ExcelChart SetSeriesLineColor(string seriesName, string color, double? widthPoints = null, bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            if (string.IsNullOrWhiteSpace(color)) {
                throw new ArgumentException("Series color cannot be null or empty.", nameof(color));
            }

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                C.ChartShapeProperties props = EnsureChartShapeProperties(series);
                ApplyLine(props, NormalizeHexColor(color), widthPoints);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets the marker style for a chart series by index.
        /// </summary>
        public ExcelChart SetSeriesMarker(int seriesIndex, C.MarkerStyleValues style, int? size = null, string? fillColor = null, string? lineColor = null, double? lineWidthPoints = null) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            if (size is < 1 or > 72) {
                throw new ArgumentOutOfRangeException(nameof(size), "Marker size must be between 1 and 72.");
            }
            if (fillColor != null && string.IsNullOrWhiteSpace(fillColor)) {
                throw new ArgumentException("Marker fill color cannot be empty.", nameof(fillColor));
            }
            if (lineColor != null && string.IsNullOrWhiteSpace(lineColor)) {
                throw new ArgumentException("Marker line color cannot be empty.", nameof(lineColor));
            }

            bool applied = ApplySeriesMarkerByIndex(seriesIndex, marker => {
                ApplyMarker(marker, style, size, fillColor, lineColor, lineWidthPoints);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets the marker style for a chart series by name.
        /// </summary>
        public ExcelChart SetSeriesMarker(string seriesName, C.MarkerStyleValues style, int? size = null, string? fillColor = null, string? lineColor = null, double? lineWidthPoints = null, bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            if (size is < 1 or > 72) {
                throw new ArgumentOutOfRangeException(nameof(size), "Marker size must be between 1 and 72.");
            }
            if (fillColor != null && string.IsNullOrWhiteSpace(fillColor)) {
                throw new ArgumentException("Marker fill color cannot be empty.", nameof(fillColor));
            }
            if (lineColor != null && string.IsNullOrWhiteSpace(lineColor)) {
                throw new ArgumentException("Marker line color cannot be empty.", nameof(lineColor));
            }

            bool applied = ApplySeriesMarkerByName(seriesName, ignoreCase, marker => {
                ApplyMarker(marker, style, size, fillColor, lineColor, lineWidthPoints);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets chart area fill/line styling.
        /// </summary>
        public ExcelChart SetChartAreaStyle(string? fillColor = null, string? lineColor = null,
            double? lineWidthPoints = null, bool noFill = false, bool noLine = false) {
            ValidateDataLabelShapeStyle(fillColor, lineColor, lineWidthPoints, noFill, noLine);

            ChartPart chartPart = GetChartPart();
            C.ChartSpace? chartSpace = chartPart.ChartSpace;
            if (chartSpace == null) {
                return this;
            }

            C.ShapeProperties props = chartSpace.GetFirstChild<C.ShapeProperties>() ?? new C.ShapeProperties();
            ApplyAreaStyle(props, fillColor, lineColor, lineWidthPoints, noFill, noLine);
            if (props.Parent == null) {
                chartSpace.Append(props);
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets plot area fill/line styling.
        /// </summary>
        public ExcelChart SetPlotAreaStyle(string? fillColor = null, string? lineColor = null,
            double? lineWidthPoints = null, bool noFill = false, bool noLine = false) {
            ValidateDataLabelShapeStyle(fillColor, lineColor, lineWidthPoints, noFill, noLine);

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            C.ShapeProperties props = plotArea.GetFirstChild<C.ShapeProperties>() ?? new C.ShapeProperties();
            ApplyAreaStyle(props, fillColor, lineColor, lineWidthPoints, noFill, noLine);
            if (props.Parent == null) {
                plotArea.Append(props);
            }

            Save();
            return this;
        }

        /// <summary>
        /// Adds or replaces a trendline for a chart series by index.
        /// </summary>
        public ExcelChart SetSeriesTrendline(int seriesIndex, C.TrendlineValues type,
            int? order = null, int? period = null, double? forward = null, double? backward = null, double? intercept = null,
            bool displayEquation = false, bool displayRSquared = false, string? lineColor = null, double? lineWidthPoints = null) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            ValidateTrendline(type, order, period, forward, backward, lineColor, lineWidthPoints);

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                ApplyTrendline(series, type, order, period, forward, backward, intercept, displayEquation, displayRSquared, lineColor, lineWidthPoints);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Adds or replaces a trendline for a chart series by name.
        /// </summary>
        public ExcelChart SetSeriesTrendline(string seriesName, C.TrendlineValues type,
            int? order = null, int? period = null, double? forward = null, double? backward = null, double? intercept = null,
            bool displayEquation = false, bool displayRSquared = false, string? lineColor = null, double? lineWidthPoints = null,
            bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            ValidateTrendline(type, order, period, forward, backward, lineColor, lineWidthPoints);

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                ApplyTrendline(series, type, order, period, forward, backward, intercept, displayEquation, displayRSquared, lineColor, lineWidthPoints);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Removes trendlines from a chart series by index.
        /// </summary>
        public ExcelChart ClearSeriesTrendline(int seriesIndex) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                series.RemoveAllChildren<C.Trendline>();
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Removes trendlines from a chart series by name.
        /// </summary>
        public ExcelChart ClearSeriesTrendline(string seriesName, bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                series.RemoveAllChildren<C.Trendline>();
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        private static void ApplyDataLabels(OpenXmlCompositeElement chartElement, bool showLegendKey, bool showValue,
            bool showCategoryName, bool showSeriesName, bool showPercent) {
            ApplyDataLabels(chartElement, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent,
                null, null, false);
        }

        private static void ApplyDataLabels(OpenXmlCompositeElement chartElement, bool showLegendKey, bool showValue,
            bool showCategoryName, bool showSeriesName, bool showPercent,
            C.DataLabelPositionValues? position, string? numberFormat, bool sourceLinked) {
            C.DataLabels labels = chartElement.GetFirstChild<C.DataLabels>() ?? new C.DataLabels();
            ReplaceChild(labels, new C.ShowLegendKey { Val = showLegendKey });
            ReplaceChild(labels, new C.ShowValue { Val = showValue });
            ReplaceChild(labels, new C.ShowCategoryName { Val = showCategoryName });
            ReplaceChild(labels, new C.ShowSeriesName { Val = showSeriesName });
            ReplaceChild(labels, new C.ShowPercent { Val = showPercent });
            ReplaceChild(labels, new C.ShowBubbleSize { Val = false });

            if (position != null) {
                ReplaceChild(labels, new C.DataLabelPosition { Val = position.Value });
            }

            if (numberFormat != null) {
                ReplaceChild(labels, new C.NumberingFormat {
                    FormatCode = numberFormat,
                    SourceLinked = sourceLinked
                });
            }

            if (chartElement.GetFirstChild<C.DataLabels>() == null) {
                chartElement.Append(labels);
            }
        }

        private static void ApplyTextStyle(A.TextCharacterPropertiesType runProps, double? fontSizePoints, bool? bold,
            bool? italic, string? color, string? fontName) {
            if (fontSizePoints != null) {
                runProps.FontSize = (int)Math.Round(fontSizePoints.Value * 100);
            }
            if (bold != null) {
                runProps.Bold = bold.Value;
            }
            if (italic != null) {
                runProps.Italic = italic.Value;
            }
            if (fontName != null) {
                runProps.RemoveAllChildren<A.LatinFont>();
                runProps.Append(new A.LatinFont { Typeface = fontName });
            }
            if (color != null) {
                ApplySolidFill(runProps, NormalizeHexColor(color));
            }
        }

        private static void ValidateDataLabelTextStyle(double? fontSizePoints, string? color, string? fontName) {
            if (fontSizePoints != null && fontSizePoints <= 0) {
                throw new ArgumentOutOfRangeException(nameof(fontSizePoints));
            }
            if (color != null && string.IsNullOrWhiteSpace(color)) {
                throw new ArgumentException("Label color cannot be empty.", nameof(color));
            }
            if (fontName != null && string.IsNullOrWhiteSpace(fontName)) {
                throw new ArgumentException("Font name cannot be empty.", nameof(fontName));
            }
        }

        private static void ValidateDataLabelShapeStyle(string? fillColor, string? lineColor, double? lineWidthPoints,
            bool noFill, bool noLine) {
            if (fillColor != null && string.IsNullOrWhiteSpace(fillColor)) {
                throw new ArgumentException("Fill color cannot be empty.", nameof(fillColor));
            }
            if (lineColor != null && string.IsNullOrWhiteSpace(lineColor)) {
                throw new ArgumentException("Line color cannot be empty.", nameof(lineColor));
            }
            if (lineWidthPoints != null && lineWidthPoints <= 0) {
                throw new ArgumentOutOfRangeException(nameof(lineWidthPoints));
            }
            if (noFill && fillColor != null) {
                throw new ArgumentException("Cannot set both fill color and noFill.", nameof(noFill));
            }
            if (noLine && (lineColor != null || lineWidthPoints != null)) {
                throw new ArgumentException("Cannot set line color/width when noLine is true.", nameof(noLine));
            }
        }

        private static void ValidateDataLabelLeaderLines(string? lineColor, double? lineWidthPoints) {
            if (lineColor != null && string.IsNullOrWhiteSpace(lineColor)) {
                throw new ArgumentException("Line color cannot be empty.", nameof(lineColor));
            }
            if (lineWidthPoints != null && lineWidthPoints <= 0) {
                throw new ArgumentOutOfRangeException(nameof(lineWidthPoints));
            }
        }

        private static void ValidateAxisGridlinesStyle(string? lineColor, double? lineWidthPoints) {
            if (lineColor != null && string.IsNullOrWhiteSpace(lineColor)) {
                throw new ArgumentException("Gridline color cannot be empty.", nameof(lineColor));
            }
            if (lineWidthPoints != null && lineWidthPoints <= 0) {
                throw new ArgumentOutOfRangeException(nameof(lineWidthPoints));
            }
        }

        private static void ValidateAxisScale(double? minimum, double? maximum, double? majorUnit, double? minorUnit,
            bool? logScale, double? logBase) {
            if (minimum != null && double.IsNaN(minimum.Value)) {
                throw new ArgumentOutOfRangeException(nameof(minimum));
            }
            if (maximum != null && double.IsNaN(maximum.Value)) {
                throw new ArgumentOutOfRangeException(nameof(maximum));
            }
            if (minimum != null && maximum != null && minimum.Value >= maximum.Value) {
                throw new ArgumentException("Minimum must be less than maximum.");
            }
            if (majorUnit != null && majorUnit.Value <= 0) {
                throw new ArgumentOutOfRangeException(nameof(majorUnit));
            }
            if (minorUnit != null && minorUnit.Value <= 0) {
                throw new ArgumentOutOfRangeException(nameof(minorUnit));
            }
            if (logScale == false && logBase != null) {
                throw new ArgumentException("Log base requires logScale to be enabled.", nameof(logBase));
            }

            bool effectiveLog = logScale == true || logBase != null;
            if (effectiveLog) {
                double baseValue = logBase ?? 10d;
                if (baseValue <= 1d) {
                    throw new ArgumentOutOfRangeException(nameof(logBase), "Log base must be greater than 1.");
                }
                if (minimum != null && minimum.Value <= 0) {
                    throw new ArgumentException("Minimum must be greater than 0 for log scale.", nameof(minimum));
                }
                if (maximum != null && maximum.Value <= 0) {
                    throw new ArgumentException("Maximum must be greater than 0 for log scale.", nameof(maximum));
                }
            }
        }

        private static void ValidateCrossesAtForAxis(OpenXmlCompositeElement axis, double? crossesAt) {
            if (crossesAt == null) {
                return;
            }

            C.Scaling? scaling = axis.GetFirstChild<C.Scaling>();
            if (scaling?.GetFirstChild<C.LogBase>() != null && crossesAt.Value <= 0) {
                throw new ArgumentOutOfRangeException(nameof(crossesAt), "Crosses-at value must be greater than 0 for log scale.");
            }
        }

        private static void ApplyAxisScale(OpenXmlCompositeElement axis, double? minimum, double? maximum,
            double? majorUnit, double? minorUnit, bool? reverseOrder, bool? logScale, double? logBase) {
            if (reverseOrder != null || minimum != null || maximum != null || logScale != null || logBase != null) {
                C.Scaling scaling = EnsureScaling(axis);
                if (reverseOrder != null) {
                    ReplaceChild(scaling, new C.Orientation {
                        Val = reverseOrder.Value ? C.OrientationValues.MaxMin : C.OrientationValues.MinMax
                    });
                }
                if (minimum != null) {
                    ReplaceChild(scaling, new C.MinAxisValue { Val = minimum.Value });
                }
                if (maximum != null) {
                    ReplaceChild(scaling, new C.MaxAxisValue { Val = maximum.Value });
                }

                bool effectiveLog = logScale == true || logBase != null;
                if (effectiveLog) {
                    double baseValue = logBase ?? 10d;
                    ReplaceChild(scaling, new C.LogBase { Val = baseValue });
                } else if (logScale == false) {
                    scaling.GetFirstChild<C.LogBase>()?.Remove();
                }
            }

            if (majorUnit != null) {
                ReplaceChild(axis, new C.MajorUnit { Val = majorUnit.Value });
            }
            if (minorUnit != null) {
                ReplaceChild(axis, new C.MinorUnit { Val = minorUnit.Value });
            }
        }

        private static C.Scaling EnsureScaling(OpenXmlCompositeElement axis) {
            C.Scaling scaling = axis.GetFirstChild<C.Scaling>() ?? new C.Scaling();
            if (scaling.Parent == null) {
                C.AxisId? axisId = axis.GetFirstChild<C.AxisId>();
                if (axisId != null) {
                    axis.InsertAfter(scaling, axisId);
                } else {
                    axis.InsertAt(scaling, 0);
                }
            }
            return scaling;
        }

        private static void ValidateTrendline(C.TrendlineValues type, int? order, int? period,
            double? forward, double? backward, string? lineColor, double? lineWidthPoints) {
            if (order != null && (order <= 0 || order > byte.MaxValue)) {
                throw new ArgumentOutOfRangeException(nameof(order));
            }
            if (period != null && period <= 0) {
                throw new ArgumentOutOfRangeException(nameof(period));
            }
            if (forward != null && forward < 0) {
                throw new ArgumentOutOfRangeException(nameof(forward));
            }
            if (backward != null && backward < 0) {
                throw new ArgumentOutOfRangeException(nameof(backward));
            }
            if (lineColor != null && string.IsNullOrWhiteSpace(lineColor)) {
                throw new ArgumentException("Trendline color cannot be empty.", nameof(lineColor));
            }
            if (lineWidthPoints != null && lineWidthPoints <= 0) {
                throw new ArgumentOutOfRangeException(nameof(lineWidthPoints));
            }

            bool isPolynomial = type.Equals(C.TrendlineValues.Polynomial);
            bool isMovingAverage = type.Equals(C.TrendlineValues.MovingAverage);
            if (isPolynomial && order == null) {
                throw new ArgumentException("Polynomial trendlines require an order.", nameof(order));
            }
            if (!isPolynomial && order != null) {
                throw new ArgumentException("Order is only valid for polynomial trendlines.", nameof(order));
            }
            if (isMovingAverage && period == null) {
                throw new ArgumentException("Moving average trendlines require a period.", nameof(period));
            }
            if (!isMovingAverage && period != null) {
                throw new ArgumentException("Period is only valid for moving average trendlines.", nameof(period));
            }
        }

        private static C.DataLabels EnsureDataLabels(OpenXmlCompositeElement chartElement) {
            C.DataLabels labels = chartElement.GetFirstChild<C.DataLabels>() ?? new C.DataLabels();
            if (labels.Parent == null) {
                chartElement.Append(labels);
            }
            return labels;
        }

        private static void ApplyDataLabelTextStyle(OpenXmlCompositeElement labels, double? fontSizePoints, bool? bold,
            bool? italic, string? color, string? fontName) {
            ApplyTextStyle(EnsureTextPropertiesRunProperties(labels), fontSizePoints, bold, italic, color, fontName);
        }

        private static void ApplyDataLabelShapeStyle(OpenXmlCompositeElement labels, string? fillColor, string? lineColor,
            double? lineWidthPoints, bool noFill, bool noLine) {
            C.ChartShapeProperties props = EnsureDataLabelShapeProperties(labels);
            if (noFill) {
                ApplyNoFill(props);
            } else if (fillColor != null) {
                ApplySolidFill(props, NormalizeHexColor(fillColor));
            }

            if (noLine) {
                ApplyNoLine(props);
            } else if (lineColor != null || lineWidthPoints != null) {
                string? normalizedLine = lineColor != null ? NormalizeHexColor(lineColor) : null;
                ApplyOptionalLine(props, normalizedLine, lineWidthPoints);
            }
        }

        private static void ApplyAreaStyle(OpenXmlCompositeElement props, string? fillColor, string? lineColor,
            double? lineWidthPoints, bool noFill, bool noLine) {
            if (noFill) {
                ApplyNoFill(props);
            } else if (fillColor != null) {
                ApplySolidFill(props, NormalizeHexColor(fillColor));
            }

            if (noLine) {
                ApplyNoLine(props);
            } else if (lineColor != null || lineWidthPoints != null) {
                string? normalizedLine = lineColor != null ? NormalizeHexColor(lineColor) : null;
                ApplyOptionalLine(props, normalizedLine, lineWidthPoints);
            }
        }

        private static void ApplyDataLabelLeaderLines(C.DataLabels labels, bool showLeaderLines, string? lineColor,
            double? lineWidthPoints) {
            ReplaceChild(labels, new C.ShowLeaderLines { Val = showLeaderLines });

            if (lineColor != null || lineWidthPoints != null) {
                C.LeaderLines leaderLines = labels.GetFirstChild<C.LeaderLines>() ?? new C.LeaderLines();
                C.ChartShapeProperties props = leaderLines.GetFirstChild<C.ChartShapeProperties>() ?? new C.ChartShapeProperties();
                string? normalizedLine = lineColor != null ? NormalizeHexColor(lineColor) : null;
                ApplyOptionalLine(props, normalizedLine, lineWidthPoints);
                if (props.Parent == null) {
                    leaderLines.Append(props);
                }
                if (leaderLines.Parent == null) {
                    labels.Append(leaderLines);
                }
            }
        }

        private static void ApplyDataLabelOverrides(OpenXmlCompositeElement label, bool? showLegendKey, bool? showValue,
            bool? showCategoryName, bool? showSeriesName, bool? showPercent, bool? showBubbleSize,
            C.DataLabelPositionValues? position, string? numberFormat, bool sourceLinked) {
            if (showLegendKey != null) {
                ReplaceChild(label, new C.ShowLegendKey { Val = showLegendKey.Value });
            }
            if (showValue != null) {
                ReplaceChild(label, new C.ShowValue { Val = showValue.Value });
            }
            if (showCategoryName != null) {
                ReplaceChild(label, new C.ShowCategoryName { Val = showCategoryName.Value });
            }
            if (showSeriesName != null) {
                ReplaceChild(label, new C.ShowSeriesName { Val = showSeriesName.Value });
            }
            if (showPercent != null) {
                ReplaceChild(label, new C.ShowPercent { Val = showPercent.Value });
            }
            if (showBubbleSize != null) {
                ReplaceChild(label, new C.ShowBubbleSize { Val = showBubbleSize.Value });
            }
            if (position != null) {
                ReplaceChild(label, new C.DataLabelPosition { Val = position.Value });
            }
            if (numberFormat != null) {
                ReplaceChild(label, new C.NumberingFormat {
                    FormatCode = numberFormat,
                    SourceLinked = sourceLinked
                });
            }
        }

        private static void ApplyDataLabelSeparator(OpenXmlCompositeElement label, string? separator) {
            C.Separator? existing = label.GetFirstChild<C.Separator>();
            if (separator == null) {
                existing?.Remove();
                return;
            }

            existing?.Remove();
            label.Append(new C.Separator { Text = separator });
        }

        private static void ApplyDataLabelTemplate(OpenXmlCompositeElement series, ExcelChartDataLabelTemplate template) {
            if (template.NumberFormat != null && string.IsNullOrWhiteSpace(template.NumberFormat)) {
                throw new ArgumentException("Number format cannot be empty.", nameof(template.NumberFormat));
            }

            bool applyTextStyle = template.FontSizePoints != null
                || template.Bold != null
                || template.Italic != null
                || template.TextColor != null
                || template.FontName != null;
            bool applyShapeStyle = template.NoFill
                || template.NoLine
                || template.FillColor != null
                || template.LineColor != null
                || template.LineWidthPoints != null;
            bool applyLeaderLines = template.ShowLeaderLines != null
                || template.LeaderLineColor != null
                || template.LeaderLineWidthPoints != null;

            if (applyTextStyle) {
                ValidateDataLabelTextStyle(template.FontSizePoints, template.TextColor, template.FontName);
            }
            if (applyShapeStyle) {
                ValidateDataLabelShapeStyle(template.FillColor, template.LineColor, template.LineWidthPoints,
                    template.NoFill, template.NoLine);
            }
            if (applyLeaderLines) {
                ValidateDataLabelLeaderLines(template.LeaderLineColor, template.LeaderLineWidthPoints);
            }

            C.DataLabels labels = EnsureDataLabels(series);
            ApplyDataLabelOverrides(labels, template.ShowLegendKey, template.ShowValue, template.ShowCategoryName,
                template.ShowSeriesName, template.ShowPercent, template.ShowBubbleSize,
                template.Position, template.NumberFormat, template.SourceLinked);
            if (template.Separator != null) {
                ApplyDataLabelSeparator(labels, template.Separator);
            }

            if (applyTextStyle) {
                ApplyDataLabelTextStyle(labels, template.FontSizePoints, template.Bold, template.Italic,
                    template.TextColor, template.FontName);
            }
            if (applyShapeStyle) {
                ApplyDataLabelShapeStyle(labels, template.FillColor, template.LineColor, template.LineWidthPoints,
                    template.NoFill, template.NoLine);
            }
            if (applyLeaderLines) {
                bool showLeaderLines = template.ShowLeaderLines ?? true;
                ApplyDataLabelLeaderLines(labels, showLeaderLines, template.LeaderLineColor, template.LeaderLineWidthPoints);
            }
        }

        private static C.DataLabel EnsureDataLabel(OpenXmlCompositeElement series, int pointIndex) {
            C.DataLabels labels = EnsureDataLabels(series);
            uint idx = (uint)pointIndex;
            C.DataLabel? label = labels.Elements<C.DataLabel>()
                .FirstOrDefault(item => item.GetFirstChild<C.Index>()?.Val?.Value == idx);
            if (label == null) {
                label = new C.DataLabel();
                label.Append(new C.Index { Val = idx });
                labels.Append(label);
            }
            return label;
        }

        private static void ApplyGridlines(OpenXmlCompositeElement axis, bool showMajor, bool showMinor,
            string? lineColor, double? lineWidthPoints) {
            if (showMajor) {
                C.MajorGridlines major = axis.GetFirstChild<C.MajorGridlines>() ?? new C.MajorGridlines();
                ApplyGridlineStyle(major, lineColor, lineWidthPoints);
                if (major.Parent == null) {
                    axis.Append(major);
                }
            } else {
                axis.GetFirstChild<C.MajorGridlines>()?.Remove();
            }

            if (showMinor) {
                C.MinorGridlines minor = axis.GetFirstChild<C.MinorGridlines>() ?? new C.MinorGridlines();
                ApplyGridlineStyle(minor, lineColor, lineWidthPoints);
                if (minor.Parent == null) {
                    axis.Append(minor);
                }
            } else {
                axis.GetFirstChild<C.MinorGridlines>()?.Remove();
            }
        }

        private static void ApplyTrendline(OpenXmlCompositeElement series, C.TrendlineValues type, int? order, int? period,
            double? forward, double? backward, double? intercept, bool displayEquation, bool displayRSquared,
            string? lineColor, double? lineWidthPoints) {
            if (!IsTrendlineSupportedSeries(series)) {
                throw new InvalidOperationException("Trendlines are only supported for line, bar/column, area, scatter, and bubble series.");
            }

            series.RemoveAllChildren<C.Trendline>();
            C.Trendline trendline = new C.Trendline();
            trendline.Append(new C.TrendlineType { Val = type });

            if (type.Equals(C.TrendlineValues.Polynomial) && order != null) {
                trendline.Append(new C.PolynomialOrder { Val = (byte)order.Value });
            }
            if (type.Equals(C.TrendlineValues.MovingAverage) && period != null) {
                trendline.Append(new C.Period { Val = (uint)period.Value });
            }
            if (forward != null) {
                trendline.Append(new C.Forward { Val = forward.Value });
            }
            if (backward != null) {
                trendline.Append(new C.Backward { Val = backward.Value });
            }
            if (intercept != null) {
                trendline.Append(new C.Intercept { Val = intercept.Value });
            }
            if (displayEquation) {
                trendline.Append(new C.DisplayEquation { Val = displayEquation });
            }
            if (displayRSquared) {
                trendline.Append(new C.DisplayRSquaredValue { Val = displayRSquared });
            }

            if (lineColor != null || lineWidthPoints != null) {
                C.ChartShapeProperties props = new C.ChartShapeProperties();
                string? normalizedLine = lineColor != null ? NormalizeHexColor(lineColor) : null;
                ApplyOptionalLine(props, normalizedLine, lineWidthPoints);
                trendline.Append(props);
            }

            series.Append(trendline);
        }

        private static bool IsTrendlineSupportedSeries(OpenXmlCompositeElement series) {
            return series is C.LineChartSeries
                || series is C.BarChartSeries
                || series is C.AreaChartSeries
                || series is C.ScatterChartSeries
                || series is C.BubbleChartSeries;
        }

        private static void ApplyGridlineStyle(OpenXmlCompositeElement gridlines, string? lineColor, double? lineWidthPoints) {
            if (lineColor == null && lineWidthPoints == null) {
                return;
            }
            C.ChartShapeProperties props = gridlines.GetFirstChild<C.ChartShapeProperties>() ?? new C.ChartShapeProperties();
            string? normalizedLine = lineColor != null ? NormalizeHexColor(lineColor) : null;
            ApplyOptionalLine(props, normalizedLine, lineWidthPoints);
            if (props.Parent == null) {
                gridlines.Append(props);
            }
        }

        private static void ApplyAxisCrossing(OpenXmlCompositeElement axis, C.CrossesValues crosses, double? crossesAt) {
            if (crossesAt != null) {
                ReplaceChild(axis, new C.CrossesAt { Val = crossesAt.Value });
                axis.GetFirstChild<C.Crosses>()?.Remove();
            } else {
                ReplaceChild(axis, new C.Crosses { Val = crosses });
                axis.GetFirstChild<C.CrossesAt>()?.Remove();
            }
        }

        private static void ApplyDisplayUnitsLabel(C.DisplayUnits displayUnits, bool showLabel, string? labelText = null) {
            if (!showLabel) {
                displayUnits.GetFirstChild<C.DisplayUnitsLabel>()?.Remove();
                return;
            }

            C.DisplayUnitsLabel label = displayUnits.GetFirstChild<C.DisplayUnitsLabel>() ?? new C.DisplayUnitsLabel();
            if (label.GetFirstChild<C.Layout>() == null) {
                label.Append(new C.Layout());
            }
            if (labelText != null) {
                label.RemoveAllChildren<C.ChartText>();
                label.Append(CreateChartText(labelText));
            }
            if (label.Parent == null) {
                displayUnits.Append(label);
            }
        }

        private static C.ChartShapeProperties EnsureDataLabelShapeProperties(OpenXmlCompositeElement labels) {
            C.ChartShapeProperties props = labels.GetFirstChild<C.ChartShapeProperties>() ?? new C.ChartShapeProperties();
            if (props.Parent == null) {
                labels.Append(props);
            }
            return props;
        }

        private static void ApplyNoFill(OpenXmlCompositeElement props) {
            props.RemoveAllChildren<A.SolidFill>();
            props.RemoveAllChildren<A.GradientFill>();
            props.RemoveAllChildren<A.PatternFill>();
            props.RemoveAllChildren<A.NoFill>();
            props.Append(new A.NoFill());
        }

        private static void ApplyNoLine(OpenXmlCompositeElement props) {
            A.Outline outline = props.GetFirstChild<A.Outline>() ?? new A.Outline();
            outline.RemoveAllChildren<A.SolidFill>();
            outline.RemoveAllChildren<A.NoFill>();
            outline.Append(new A.NoFill());
            if (outline.Parent == null) {
                props.Append(outline);
            }
        }

        private static A.DefaultRunProperties EnsureTextPropertiesRunProperties(OpenXmlCompositeElement parent) {
            C.TextProperties textProps = parent.GetFirstChild<C.TextProperties>() ?? new C.TextProperties();
            if (textProps.GetFirstChild<A.BodyProperties>() == null) {
                textProps.Append(new A.BodyProperties());
            }
            if (textProps.GetFirstChild<A.ListStyle>() == null) {
                textProps.Append(new A.ListStyle());
            }

            A.Paragraph paragraph = textProps.GetFirstChild<A.Paragraph>() ?? new A.Paragraph();
            if (paragraph.Parent == null) {
                textProps.Append(paragraph);
            }

            A.ParagraphProperties paragraphProps = paragraph.GetFirstChild<A.ParagraphProperties>() ?? new A.ParagraphProperties();
            if (paragraphProps.Parent == null) {
                paragraph.Append(paragraphProps);
            }

            A.DefaultRunProperties runProps = paragraphProps.GetFirstChild<A.DefaultRunProperties>() ?? new A.DefaultRunProperties();
            if (runProps.Parent == null) {
                paragraphProps.Append(runProps);
            }

            if (textProps.Parent == null) {
                parent.Append(textProps);
            }

            return runProps;
        }

        private static A.RunProperties EnsureChartTextRunProperties(C.ChartText chartText) {
            C.RichText richText = chartText.GetFirstChild<C.RichText>() ?? new C.RichText();
            if (richText.GetFirstChild<A.BodyProperties>() == null) {
                richText.Append(new A.BodyProperties());
            }
            if (richText.GetFirstChild<A.ListStyle>() == null) {
                richText.Append(new A.ListStyle());
            }

            A.Paragraph paragraph = richText.GetFirstChild<A.Paragraph>() ?? new A.Paragraph();
            if (paragraph.Parent == null) {
                richText.Append(paragraph);
            }

            A.Run run = paragraph.GetFirstChild<A.Run>() ?? new A.Run();
            if (run.Parent == null) {
                paragraph.Append(run);
            }

            A.RunProperties runProps = run.GetFirstChild<A.RunProperties>() ?? new A.RunProperties();
            if (runProps.Parent == null) {
                run.Append(runProps);
            }

            if (richText.Parent == null) {
                chartText.Append(richText);
            }

            return runProps;
        }

        private static C.ChartShapeProperties EnsureChartShapeProperties(OpenXmlCompositeElement series) {
            C.ChartShapeProperties props = series.GetFirstChild<C.ChartShapeProperties>() ?? new C.ChartShapeProperties();
            if (props.Parent == null) {
                series.Append(props);
            }
            return props;
        }

        private static void ApplySolidFill(OpenXmlCompositeElement props, string color) {
            props.RemoveAllChildren<A.SolidFill>();
            props.RemoveAllChildren<A.NoFill>();
            props.RemoveAllChildren<A.GradientFill>();
            props.RemoveAllChildren<A.PatternFill>();
            props.Append(new A.SolidFill(new A.RgbColorModelHex { Val = color }));
        }

        private static void ApplyLine(OpenXmlCompositeElement props, string color, double? widthPoints) {
            A.Outline outline = props.GetFirstChild<A.Outline>() ?? new A.Outline();
            outline.RemoveAllChildren<A.SolidFill>();
            outline.Append(new A.SolidFill(new A.RgbColorModelHex { Val = color }));
            if (widthPoints != null) {
                outline.Width = (int)Math.Round(widthPoints.Value * 12700d);
            }

            if (outline.Parent == null) {
                props.Append(outline);
            }
        }

        private static void ApplyOptionalLine(OpenXmlCompositeElement props, string? color, double? widthPoints) {
            A.Outline outline = props.GetFirstChild<A.Outline>() ?? new A.Outline();
            if (color != null) {
                outline.RemoveAllChildren<A.SolidFill>();
                outline.Append(new A.SolidFill(new A.RgbColorModelHex { Val = color }));
            }
            if (widthPoints != null) {
                outline.Width = (int)Math.Round(widthPoints.Value * 12700d);
            }
            if (outline.Parent == null) {
                props.Append(outline);
            }
        }

        private static void ApplyMarker(C.Marker marker, C.MarkerStyleValues style, int? size, string? fillColor, string? lineColor, double? lineWidthPoints) {
            marker.Symbol = new C.Symbol { Val = style };
            if (size != null) {
                marker.Size = new C.Size { Val = (byte)size.Value };
            }

            if (fillColor != null || lineColor != null || lineWidthPoints != null) {
                C.ChartShapeProperties props = marker.ChartShapeProperties ?? new C.ChartShapeProperties();
                if (fillColor != null) {
                    ApplySolidFill(props, NormalizeHexColor(fillColor));
                }
                if (lineColor != null || lineWidthPoints != null) {
                    string? normalizedLine = lineColor != null ? NormalizeHexColor(lineColor) : null;
                    ApplyOptionalLine(props, normalizedLine, lineWidthPoints);
                }
                if (props.Parent == null) {
                    marker.Append(props);
                }
            }
        }

        private bool ApplySeriesByIndex(int seriesIndex, Action<OpenXmlCompositeElement> apply) {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return false;
            }

            if (ApplySeriesByIndex(plotArea.Elements<C.BarChart>(), seriesIndex, apply)) return true;
            if (ApplySeriesByIndex(plotArea.Elements<C.LineChart>(), seriesIndex, apply)) return true;
            if (ApplySeriesByIndex(plotArea.Elements<C.AreaChart>(), seriesIndex, apply)) return true;
            if (ApplySeriesByIndex(plotArea.Elements<C.PieChart>(), seriesIndex, apply)) return true;
            if (ApplySeriesByIndex(plotArea.Elements<C.DoughnutChart>(), seriesIndex, apply)) return true;
            if (ApplySeriesByIndex(plotArea.Elements<C.ScatterChart>(), seriesIndex, apply)) return true;
            if (ApplySeriesByIndex(plotArea.Elements<C.BubbleChart>(), seriesIndex, apply)) return true;

            return false;
        }

        private bool ApplySeriesByName(string seriesName, bool ignoreCase, Action<OpenXmlCompositeElement> apply) {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return false;
            }

            if (ApplySeriesByName(plotArea.Elements<C.BarChart>(), seriesName, ignoreCase, apply)) return true;
            if (ApplySeriesByName(plotArea.Elements<C.LineChart>(), seriesName, ignoreCase, apply)) return true;
            if (ApplySeriesByName(plotArea.Elements<C.AreaChart>(), seriesName, ignoreCase, apply)) return true;
            if (ApplySeriesByName(plotArea.Elements<C.PieChart>(), seriesName, ignoreCase, apply)) return true;
            if (ApplySeriesByName(plotArea.Elements<C.DoughnutChart>(), seriesName, ignoreCase, apply)) return true;
            if (ApplySeriesByName(plotArea.Elements<C.ScatterChart>(), seriesName, ignoreCase, apply)) return true;
            if (ApplySeriesByName(plotArea.Elements<C.BubbleChart>(), seriesName, ignoreCase, apply)) return true;

            return false;
        }

        private bool ApplySeriesMarkerByIndex(int seriesIndex, Action<C.Marker> apply) {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return false;
            }

            if (ApplySeriesMarkerByIndex(plotArea.Elements<C.LineChart>(), seriesIndex, apply)) return true;
            if (ApplySeriesMarkerByIndex(plotArea.Elements<C.ScatterChart>(), seriesIndex, apply)) return true;

            return false;
        }

        private bool ApplySeriesMarkerByName(string seriesName, bool ignoreCase, Action<C.Marker> apply) {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return false;
            }

            if (ApplySeriesMarkerByName(plotArea.Elements<C.LineChart>(), seriesName, ignoreCase, apply)) return true;
            if (ApplySeriesMarkerByName(plotArea.Elements<C.ScatterChart>(), seriesName, ignoreCase, apply)) return true;

            return false;
        }

        private static bool ApplySeriesByIndex<TChart>(IEnumerable<TChart> charts, int seriesIndex,
            Action<OpenXmlCompositeElement> apply) where TChart : OpenXmlCompositeElement {
            foreach (TChart chart in charts) {
                List<OpenXmlCompositeElement> series = chart.ChildElements
                    .OfType<OpenXmlCompositeElement>()
                    .Where(IsSeriesElement)
                    .ToList();

                OpenXmlCompositeElement? match = series.FirstOrDefault(s => GetSeriesIndex(s) == seriesIndex);
                if (match != null) {
                    apply(match);
                    return true;
                }

                if (seriesIndex < 0 || seriesIndex >= series.Count) {
                    continue;
                }

                apply(series[seriesIndex]);
                return true;
            }

            return false;
        }

        private static bool ApplySeriesByName<TChart>(IEnumerable<TChart> charts, string seriesName, bool ignoreCase,
            Action<OpenXmlCompositeElement> apply) where TChart : OpenXmlCompositeElement {
            StringComparison comparison = ignoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
            foreach (TChart chart in charts) {
                foreach (OpenXmlCompositeElement series in chart.ChildElements.OfType<OpenXmlCompositeElement>().Where(IsSeriesElement)) {
                    string? name = GetSeriesName(series);
                    if (name != null && string.Equals(name, seriesName, comparison)) {
                        apply(series);
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool ApplySeriesMarkerByIndex<TChart>(IEnumerable<TChart> charts, int seriesIndex, Action<C.Marker> apply) where TChart : OpenXmlCompositeElement {
            foreach (TChart chart in charts) {
                List<OpenXmlCompositeElement> series = chart.ChildElements
                    .OfType<OpenXmlCompositeElement>()
                    .Where(IsSeriesElement)
                    .ToList();

                OpenXmlCompositeElement? seriesElement = series.FirstOrDefault(s => GetSeriesIndex(s) == seriesIndex);
                if (seriesElement == null) {
                    if (seriesIndex < 0 || seriesIndex >= series.Count) {
                        continue;
                    }
                    seriesElement = series[seriesIndex];
                }

                C.Marker marker = seriesElement.GetFirstChild<C.Marker>() ?? new C.Marker();
                apply(marker);
                if (marker.Parent == null) {
                    seriesElement.Append(marker);
                }
                return true;
            }

            return false;
        }

        private static bool ApplySeriesMarkerByName<TChart>(IEnumerable<TChart> charts, string seriesName, bool ignoreCase, Action<C.Marker> apply) where TChart : OpenXmlCompositeElement {
            StringComparison comparison = ignoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
            foreach (TChart chart in charts) {
                foreach (OpenXmlCompositeElement series in chart.ChildElements.OfType<OpenXmlCompositeElement>().Where(IsSeriesElement)) {
                    string? name = GetSeriesName(series);
                    if (name != null && string.Equals(name, seriesName, comparison)) {
                        C.Marker marker = series.GetFirstChild<C.Marker>() ?? new C.Marker();
                        apply(marker);
                        if (marker.Parent == null) {
                            series.Append(marker);
                        }
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool IsSeriesElement(OpenXmlCompositeElement element) {
            return element is C.BarChartSeries ||
                   element is C.LineChartSeries ||
                   element is C.AreaChartSeries ||
                   element is C.PieChartSeries ||
                   element is C.ScatterChartSeries ||
                   element is C.BubbleChartSeries;
        }

        private static int GetSeriesIndex(OpenXmlCompositeElement series) {
            return (int)(series.GetFirstChild<C.Index>()?.Val?.Value ?? 0U);
        }

        private static string? GetSeriesName(OpenXmlCompositeElement series) {
            C.SeriesText? seriesText = series.GetFirstChild<C.SeriesText>();
            if (seriesText == null) {
                return null;
            }

            C.StringReference? reference = seriesText.GetFirstChild<C.StringReference>();
            C.StringCache? cache = reference?.GetFirstChild<C.StringCache>();
            string? cachedText = cache?.Elements<C.StringPoint>()
                .FirstOrDefault()?
                .NumericValue?
                .Text;
            if (!string.IsNullOrWhiteSpace(cachedText)) {
                return cachedText;
            }

            C.StringLiteral? literal = seriesText.GetFirstChild<C.StringLiteral>();
            string? literalText = literal?.Elements<C.StringPoint>()
                .FirstOrDefault()?
                .NumericValue?
                .Text;
            if (!string.IsNullOrWhiteSpace(literalText)) {
                return literalText;
            }

            return string.IsNullOrWhiteSpace(seriesText.InnerText) ? null : seriesText.InnerText;
        }

        private static C.ChartText CreateChartText(string title) {
            return new C.ChartText(
                new C.RichText(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(
                        new A.Run(
                            new A.RunProperties { Language = "en-US" },
                            new A.Text { Text = title })
                    )));
        }

        private static C.Title CreateAxisTitle(string title) {
            return new C.Title(
                new C.ChartText(
                    new C.RichText(
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        new A.Paragraph(
                            new A.Run(
                                new A.RunProperties { Language = "en-US" },
                                new A.Text { Text = title })))
                ),
                new C.Layout(),
                new C.Overlay { Val = false }
            );
        }

        private static void ReplaceChild<T>(OpenXmlCompositeElement parent, T child) where T : OpenXmlElement {
            parent.GetFirstChild<T>()?.Remove();
            parent.Append(child);
        }

        private static string NormalizeHexColor(string hex) {
            hex = hex.Trim();
            if (hex.StartsWith("#", StringComparison.Ordinal)) {
                hex = hex.Substring(1);
            }
            if (hex.Length == 6) return hex.ToUpperInvariant();
            if (hex.Length == 8) return hex.Substring(2).ToUpperInvariant();
            return hex.ToUpperInvariant();
        }

        private C.Chart GetChart() {
            ChartPart chartPart = GetChartPart();
            C.Chart? chart = chartPart.ChartSpace?.GetFirstChild<C.Chart>();
            if (chart == null) {
                throw new InvalidOperationException("Chart element not found in chart part.");
            }
            return chart;
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
        }
    }
}
