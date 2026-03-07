using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.PowerPoint {
        /// <summary>
        ///     Represents a chart on a slide.
        /// </summary>
        public class PowerPointChart : PowerPointShape {
        private readonly SlidePart _slidePart;

        internal PowerPointChart(GraphicFrame frame, SlidePart slidePart) : base(frame) {
            _slidePart = slidePart ?? throw new ArgumentNullException(nameof(slidePart));
        }

        private GraphicFrame Frame => (GraphicFrame)Element;

        /// <summary>
        ///     Updates the chart data (series and categories).
        /// </summary>
        public PowerPointChart UpdateData(PowerPointChartData data) {
            if (data == null) {
                throw new ArgumentNullException(nameof(data));
            }

            ChartPart chartPart = GetChartPart();
            PowerPointUtils.UpdateChartData(chartPart, data);

            EmbeddedPackagePart? embedded = chartPart.GetPartsOfType<EmbeddedPackagePart>().FirstOrDefault();
            if (embedded != null) {
                byte[] workbookBytes = PowerPointUtils.BuildChartWorkbook(data);
                using var stream = new MemoryStream(workbookBytes);
                embedded.FeedData(stream);
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Updates scatter chart data (series X/Y values).
        /// </summary>
        public PowerPointChart UpdateData(PowerPointScatterChartData data) {
            if (data == null) {
                throw new ArgumentNullException(nameof(data));
            }

            ChartPart chartPart = GetChartPart();
            PowerPointUtils.UpdateChartData(chartPart, data);

            EmbeddedPackagePart? embedded = chartPart.GetPartsOfType<EmbeddedPackagePart>().FirstOrDefault();
            if (embedded != null) {
                byte[] workbookBytes = PowerPointUtils.BuildChartWorkbook(data);
                using var stream = new MemoryStream(workbookBytes);
                embedded.FeedData(stream);
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Updates the chart data using selectors.
        /// </summary>
        public PowerPointChart UpdateData<T>(IEnumerable<T> items, Func<T, string> categorySelector,
            params PowerPointChartSeriesDefinition<T>[] seriesDefinitions) {
            PowerPointChartData data = PowerPointChartData.From(items, categorySelector, seriesDefinitions);
            return UpdateData(data);
        }

        /// <summary>
        ///     Updates scatter chart data using selectors.
        /// </summary>
        public PowerPointChart UpdateData<T>(IEnumerable<T> items, Func<T, double> xSelector,
            params PowerPointScatterChartSeriesDefinition<T>[] seriesDefinitions) {
            PowerPointScatterChartData data = PowerPointScatterChartData.From(items, xSelector, seriesDefinitions);
            return UpdateData(data);
        }

        /// <summary>
        ///     Sets the chart title text.
        /// </summary>
        public PowerPointChart SetTitle(string title) {
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
        ///     Sets the chart title text style.
        /// </summary>
        public PowerPointChart SetTitleTextStyle(double? fontSizePoints = null, bool? bold = null, bool? italic = null,
            string? color = null, string? fontName = null) {
            ValidateTextStyle(fontSizePoints, color, fontName);

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
        ///     Removes the chart title.
        /// </summary>
        public PowerPointChart ClearTitle() {
            C.Chart chart = GetChart();
            chart.GetFirstChild<C.Title>()?.Remove();
            chart.AutoTitleDeleted = new C.AutoTitleDeleted { Val = true };
            Save();
            return this;
        }

        /// <summary>
        ///     Sets the legend position and visibility.
        /// </summary>
        public PowerPointChart SetLegend(C.LegendPositionValues position, bool overlay = false) {
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
        ///     Sets the legend text style.
        /// </summary>
        public PowerPointChart SetLegendTextStyle(double? fontSizePoints = null, bool? bold = null, bool? italic = null,
            string? color = null, string? fontName = null) {
            ValidateTextStyle(fontSizePoints, color, fontName);

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
        ///     Hides the chart legend.
        /// </summary>
        public PowerPointChart HideLegend() {
            C.Chart chart = GetChart();
            chart.GetFirstChild<C.Legend>()?.Remove();
            Save();
            return this;
        }

        /// <summary>
        ///     Configures data labels for all supported chart series.
        /// </summary>
        public PowerPointChart SetDataLabels(bool showValue = true, bool showCategoryName = false,
            bool showSeriesName = false, bool showLegendKey = false, bool showPercent = false) {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            foreach (C.BarChart barChart in plotArea.Elements<C.BarChart>()) {
                ApplyDataLabels(barChart, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent);
            }

            foreach (C.LineChart lineChart in plotArea.Elements<C.LineChart>()) {
                ApplyDataLabels(lineChart, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent);
            }

            foreach (C.AreaChart areaChart in plotArea.Elements<C.AreaChart>()) {
                ApplyDataLabels(areaChart, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent);
            }

            foreach (C.PieChart pieChart in plotArea.Elements<C.PieChart>()) {
                ApplyDataLabels(pieChart, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent);
            }

            foreach (C.DoughnutChart doughnutChart in plotArea.Elements<C.DoughnutChart>()) {
                ApplyDataLabels(doughnutChart, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent);
            }

            foreach (C.ScatterChart scatterChart in plotArea.Elements<C.ScatterChart>()) {
                ApplyDataLabels(scatterChart, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent);
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Sets the shared data label position for all supported chart types.
        /// </summary>
        public PowerPointChart SetDataLabelPosition(C.DataLabelPositionValues position) {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            foreach (C.BarChart barChart in plotArea.Elements<C.BarChart>()) {
                ReplaceChild(EnsureDataLabels(barChart), new C.DataLabelPosition { Val = position });
            }

            foreach (C.LineChart lineChart in plotArea.Elements<C.LineChart>()) {
                ReplaceChild(EnsureDataLabels(lineChart), new C.DataLabelPosition { Val = position });
            }

            foreach (C.AreaChart areaChart in plotArea.Elements<C.AreaChart>()) {
                ReplaceChild(EnsureDataLabels(areaChart), new C.DataLabelPosition { Val = position });
            }

            foreach (C.PieChart pieChart in plotArea.Elements<C.PieChart>()) {
                ReplaceChild(EnsureDataLabels(pieChart), new C.DataLabelPosition { Val = position });
            }

            foreach (C.DoughnutChart doughnutChart in plotArea.Elements<C.DoughnutChart>()) {
                ReplaceChild(EnsureDataLabels(doughnutChart), new C.DataLabelPosition { Val = position });
            }

            foreach (C.ScatterChart scatterChart in plotArea.Elements<C.ScatterChart>()) {
                ReplaceChild(EnsureDataLabels(scatterChart), new C.DataLabelPosition { Val = position });
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Sets the shared data label number format for all supported chart types.
        /// </summary>
        public PowerPointChart SetDataLabelNumberFormat(string formatCode, bool sourceLinked = false) {
            if (string.IsNullOrWhiteSpace(formatCode)) {
                throw new ArgumentException("Format code cannot be null or empty.", nameof(formatCode));
            }

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            foreach (C.BarChart barChart in plotArea.Elements<C.BarChart>()) {
                ReplaceChild(EnsureDataLabels(barChart), new C.NumberingFormat {
                    FormatCode = formatCode,
                    SourceLinked = sourceLinked
                });
            }

            foreach (C.LineChart lineChart in plotArea.Elements<C.LineChart>()) {
                ReplaceChild(EnsureDataLabels(lineChart), new C.NumberingFormat {
                    FormatCode = formatCode,
                    SourceLinked = sourceLinked
                });
            }

            foreach (C.AreaChart areaChart in plotArea.Elements<C.AreaChart>()) {
                ReplaceChild(EnsureDataLabels(areaChart), new C.NumberingFormat {
                    FormatCode = formatCode,
                    SourceLinked = sourceLinked
                });
            }

            foreach (C.PieChart pieChart in plotArea.Elements<C.PieChart>()) {
                ReplaceChild(EnsureDataLabels(pieChart), new C.NumberingFormat {
                    FormatCode = formatCode,
                    SourceLinked = sourceLinked
                });
            }

            foreach (C.DoughnutChart doughnutChart in plotArea.Elements<C.DoughnutChart>()) {
                ReplaceChild(EnsureDataLabels(doughnutChart), new C.NumberingFormat {
                    FormatCode = formatCode,
                    SourceLinked = sourceLinked
                });
            }

            foreach (C.ScatterChart scatterChart in plotArea.Elements<C.ScatterChart>()) {
                ReplaceChild(EnsureDataLabels(scatterChart), new C.NumberingFormat {
                    FormatCode = formatCode,
                    SourceLinked = sourceLinked
                });
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Sets the category axis title.
        /// </summary>
        public PowerPointChart SetCategoryAxisTitle(string title) {
            return SetAxisTitle<C.CategoryAxis>(title);
        }

        /// <summary>
        ///     Sets the value axis title.
        /// </summary>
        public PowerPointChart SetValueAxisTitle(string title) {
            return SetAxisTitle<C.ValueAxis>(title);
        }

        /// <summary>
        ///     Sets the category axis title text style.
        /// </summary>
        public PowerPointChart SetCategoryAxisTitleTextStyle(double? fontSizePoints = null, bool? bold = null,
            bool? italic = null, string? color = null, string? fontName = null) {
            return SetAxisTitleTextStyle<C.CategoryAxis>(fontSizePoints, bold, italic, color, fontName);
        }

        /// <summary>
        ///     Sets the value axis title text style.
        /// </summary>
        public PowerPointChart SetValueAxisTitleTextStyle(double? fontSizePoints = null, bool? bold = null,
            bool? italic = null, string? color = null, string? fontName = null) {
            return SetAxisTitleTextStyle<C.ValueAxis>(fontSizePoints, bold, italic, color, fontName);
        }

        /// <summary>
        ///     Sets the category axis label text style.
        /// </summary>
        public PowerPointChart SetCategoryAxisLabelTextStyle(double? fontSizePoints = null, bool? bold = null,
            bool? italic = null, string? color = null, string? fontName = null) {
            return SetAxisLabelTextStyle<C.CategoryAxis>(fontSizePoints, bold, italic, color, fontName);
        }

        /// <summary>
        ///     Sets the value axis label text style.
        /// </summary>
        public PowerPointChart SetValueAxisLabelTextStyle(double? fontSizePoints = null, bool? bold = null,
            bool? italic = null, string? color = null, string? fontName = null) {
            return SetAxisLabelTextStyle<C.ValueAxis>(fontSizePoints, bold, italic, color, fontName);
        }

        /// <summary>
        ///     Sets the category axis label rotation in degrees (-90..90).
        /// </summary>
        public PowerPointChart SetCategoryAxisLabelRotation(double rotationDegrees) {
            return SetAxisLabelRotation<C.CategoryAxis>(rotationDegrees);
        }

        /// <summary>
        ///     Sets the value axis label rotation in degrees (-90..90).
        /// </summary>
        public PowerPointChart SetValueAxisLabelRotation(double rotationDegrees) {
            return SetAxisLabelRotation<C.ValueAxis>(rotationDegrees);
        }

        /// <summary>
        ///     Sets the category axis tick label position.
        /// </summary>
        public PowerPointChart SetCategoryAxisTickLabelPosition(C.TickLabelPositionValues position) {
            return SetAxisTickLabelPosition<C.CategoryAxis>(position);
        }

        /// <summary>
        ///     Sets the value axis tick label position.
        /// </summary>
        public PowerPointChart SetValueAxisTickLabelPosition(C.TickLabelPositionValues position) {
            return SetAxisTickLabelPosition<C.ValueAxis>(position);
        }

        /// <summary>
        ///     Sets category axis gridlines visibility and optional styling.
        /// </summary>
        public PowerPointChart SetCategoryAxisGridlines(bool showMajor = true, bool showMinor = false,
            string? lineColor = null, double? lineWidthPoints = null) {
            return SetAxisGridlines<C.CategoryAxis>(showMajor, showMinor, lineColor, lineWidthPoints);
        }

        /// <summary>
        ///     Sets value axis gridlines visibility and optional styling.
        /// </summary>
        public PowerPointChart SetValueAxisGridlines(bool showMajor = true, bool showMinor = false,
            string? lineColor = null, double? lineWidthPoints = null) {
            return SetAxisGridlines<C.ValueAxis>(showMajor, showMinor, lineColor, lineWidthPoints);
        }

        /// <summary>
        ///     Sets scatter chart X-axis gridlines visibility and optional styling.
        /// </summary>
        public PowerPointChart SetScatterXAxisGridlines(bool showMajor = true, bool showMinor = false,
            string? lineColor = null, double? lineWidthPoints = null) {
            if (!CanResolveScatterAxis(ResolveScatterXAxis)) {
                return this;
            }

            return SetAxisGridlines<C.ValueAxis>(showMajor, showMinor, lineColor, lineWidthPoints,
                axis => HasAxisPosition(axis, C.AxisPositionValues.Bottom));
        }

        /// <summary>
        ///     Sets scatter chart Y-axis gridlines visibility and optional styling.
        /// </summary>
        public PowerPointChart SetScatterYAxisGridlines(bool showMajor = true, bool showMinor = false,
            string? lineColor = null, double? lineWidthPoints = null) {
            if (!CanResolveScatterAxis(ResolveScatterYAxis)) {
                return this;
            }

            return SetAxisGridlines<C.ValueAxis>(showMajor, showMinor, lineColor, lineWidthPoints,
                axis => HasAxisPosition(axis, C.AxisPositionValues.Left));
        }

        /// <summary>
        ///     Sets the scatter chart X-axis title.
        /// </summary>
        public PowerPointChart SetScatterXAxisTitle(string title) {
            if (!CanResolveScatterAxis(ResolveScatterXAxis)) {
                return this;
            }

            return SetAxisTitle<C.ValueAxis>(title, axis => HasAxisPosition(axis, C.AxisPositionValues.Bottom));
        }

        /// <summary>
        ///     Sets the scatter chart Y-axis title.
        /// </summary>
        public PowerPointChart SetScatterYAxisTitle(string title) {
            if (!CanResolveScatterAxis(ResolveScatterYAxis)) {
                return this;
            }

            return SetAxisTitle<C.ValueAxis>(title, axis => HasAxisPosition(axis, C.AxisPositionValues.Left));
        }

        /// <summary>
        ///     Sets the scatter chart X-axis label text style.
        /// </summary>
        public PowerPointChart SetScatterXAxisLabelTextStyle(double? fontSizePoints = null, bool? bold = null,
            bool? italic = null, string? color = null, string? fontName = null) {
            if (!CanResolveScatterAxis(ResolveScatterXAxis)) {
                return this;
            }

            return SetAxisLabelTextStyle<C.ValueAxis>(fontSizePoints, bold, italic, color, fontName,
                axis => HasAxisPosition(axis, C.AxisPositionValues.Bottom));
        }

        /// <summary>
        ///     Sets the scatter chart Y-axis label text style.
        /// </summary>
        public PowerPointChart SetScatterYAxisLabelTextStyle(double? fontSizePoints = null, bool? bold = null,
            bool? italic = null, string? color = null, string? fontName = null) {
            if (!CanResolveScatterAxis(ResolveScatterYAxis)) {
                return this;
            }

            return SetAxisLabelTextStyle<C.ValueAxis>(fontSizePoints, bold, italic, color, fontName,
                axis => HasAxisPosition(axis, C.AxisPositionValues.Left));
        }

        /// <summary>
        ///     Sets the scatter chart X-axis label rotation in degrees (-90..90).
        /// </summary>
        public PowerPointChart SetScatterXAxisLabelRotation(double rotationDegrees) {
            if (!CanResolveScatterAxis(ResolveScatterXAxis)) {
                return this;
            }

            return SetAxisLabelRotation<C.ValueAxis>(rotationDegrees, axis => HasAxisPosition(axis, C.AxisPositionValues.Bottom));
        }

        /// <summary>
        ///     Sets the scatter chart Y-axis label rotation in degrees (-90..90).
        /// </summary>
        public PowerPointChart SetScatterYAxisLabelRotation(double rotationDegrees) {
            if (!CanResolveScatterAxis(ResolveScatterYAxis)) {
                return this;
            }

            return SetAxisLabelRotation<C.ValueAxis>(rotationDegrees, axis => HasAxisPosition(axis, C.AxisPositionValues.Left));
        }

        /// <summary>
        ///     Sets the scatter chart X-axis tick label position.
        /// </summary>
        public PowerPointChart SetScatterXAxisTickLabelPosition(C.TickLabelPositionValues position) {
            if (!CanResolveScatterAxis(ResolveScatterXAxis)) {
                return this;
            }

            return SetAxisTickLabelPosition<C.ValueAxis>(position, axis => HasAxisPosition(axis, C.AxisPositionValues.Bottom));
        }

        /// <summary>
        ///     Sets the scatter chart Y-axis tick label position.
        /// </summary>
        public PowerPointChart SetScatterYAxisTickLabelPosition(C.TickLabelPositionValues position) {
            if (!CanResolveScatterAxis(ResolveScatterYAxis)) {
                return this;
            }

            return SetAxisTickLabelPosition<C.ValueAxis>(position, axis => HasAxisPosition(axis, C.AxisPositionValues.Left));
        }

        /// <summary>
        ///     Sets how the value axis crosses between categories.
        /// </summary>
        public PowerPointChart SetValueAxisCrossBetween(C.CrossBetweenValues between) {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            C.ValueAxis? axis = plotArea.Elements<C.ValueAxis>().FirstOrDefault();
            if (axis == null) {
                return this;
            }

            ReplaceValueAxisCrossBetween(axis, new C.CrossBetween { Val = between });
            Save();
            return this;
        }

        /// <summary>
        ///     Sets the category axis number format.
        /// </summary>
        public PowerPointChart SetCategoryAxisNumberFormat(string formatCode, bool sourceLinked = false) {
            return SetAxisNumberFormat<C.CategoryAxis>(formatCode, sourceLinked);
        }

        /// <summary>
        ///     Sets the value axis number format.
        /// </summary>
        public PowerPointChart SetValueAxisNumberFormat(string formatCode, bool sourceLinked = false) {
            return SetAxisNumberFormat<C.ValueAxis>(formatCode, sourceLinked);
        }

        /// <summary>
        ///     Sets the scatter chart X-axis number format.
        /// </summary>
        public PowerPointChart SetScatterXAxisNumberFormat(string formatCode, bool sourceLinked = false) {
            if (!CanResolveScatterAxis(ResolveScatterXAxis)) {
                return this;
            }

            return SetAxisNumberFormat<C.ValueAxis>(formatCode, sourceLinked,
                axis => HasAxisPosition(axis, C.AxisPositionValues.Bottom));
        }

        /// <summary>
        ///     Sets the scatter chart Y-axis number format.
        /// </summary>
        public PowerPointChart SetScatterYAxisNumberFormat(string formatCode, bool sourceLinked = false) {
            if (!CanResolveScatterAxis(ResolveScatterYAxis)) {
                return this;
            }

            return SetAxisNumberFormat<C.ValueAxis>(formatCode, sourceLinked,
                axis => HasAxisPosition(axis, C.AxisPositionValues.Left));
        }

        /// <summary>
        ///     Sets display units for the scatter chart X-axis.
        /// </summary>
        public PowerPointChart SetScatterXAxisDisplayUnits(C.BuiltInUnitValues unit, bool showLabel = true) {
            if (!CanResolveScatterAxis(ResolveScatterXAxis)) {
                return this;
            }

            return SetValueAxisDisplayUnitsCore(displayUnits => {
                displayUnits.RemoveAllChildren<C.CustomDisplayUnit>();
                displayUnits.RemoveAllChildren<C.BuiltInUnit>();
                displayUnits.Append(new C.BuiltInUnit { Val = unit });
            }, showLabel, null, axis => HasAxisPosition(axis, C.AxisPositionValues.Bottom));
        }

        /// <summary>
        ///     Sets display units for the scatter chart X-axis with custom label text.
        /// </summary>
        public PowerPointChart SetScatterXAxisDisplayUnits(C.BuiltInUnitValues unit, string labelText, bool showLabel = true) {
            if (string.IsNullOrWhiteSpace(labelText)) {
                throw new ArgumentException("Label text cannot be empty.", nameof(labelText));
            }
            if (!CanResolveScatterAxis(ResolveScatterXAxis)) {
                return this;
            }

            return SetValueAxisDisplayUnitsCore(displayUnits => {
                displayUnits.RemoveAllChildren<C.CustomDisplayUnit>();
                displayUnits.RemoveAllChildren<C.BuiltInUnit>();
                displayUnits.Append(new C.BuiltInUnit { Val = unit });
            }, showLabel, labelText, axis => HasAxisPosition(axis, C.AxisPositionValues.Bottom));
        }

        /// <summary>
        ///     Sets custom display units for the scatter chart X-axis.
        /// </summary>
        public PowerPointChart SetScatterXAxisDisplayUnits(double customUnit, bool showLabel = true) {
            if (!IsFinite(customUnit) || customUnit <= 0) {
                throw new ArgumentOutOfRangeException(nameof(customUnit));
            }
            if (!CanResolveScatterAxis(ResolveScatterXAxis)) {
                return this;
            }

            return SetValueAxisDisplayUnitsCore(displayUnits => {
                displayUnits.RemoveAllChildren<C.BuiltInUnit>();
                displayUnits.RemoveAllChildren<C.CustomDisplayUnit>();
                displayUnits.Append(new C.CustomDisplayUnit { Val = customUnit });
            }, showLabel, null, axis => HasAxisPosition(axis, C.AxisPositionValues.Bottom));
        }

        /// <summary>
        ///     Sets custom display units for the scatter chart X-axis with custom label text.
        /// </summary>
        public PowerPointChart SetScatterXAxisDisplayUnits(double customUnit, string labelText, bool showLabel = true) {
            if (!IsFinite(customUnit) || customUnit <= 0) {
                throw new ArgumentOutOfRangeException(nameof(customUnit));
            }
            if (string.IsNullOrWhiteSpace(labelText)) {
                throw new ArgumentException("Label text cannot be empty.", nameof(labelText));
            }
            if (!CanResolveScatterAxis(ResolveScatterXAxis)) {
                return this;
            }

            return SetValueAxisDisplayUnitsCore(displayUnits => {
                displayUnits.RemoveAllChildren<C.BuiltInUnit>();
                displayUnits.RemoveAllChildren<C.CustomDisplayUnit>();
                displayUnits.Append(new C.CustomDisplayUnit { Val = customUnit });
            }, showLabel, labelText, axis => HasAxisPosition(axis, C.AxisPositionValues.Bottom));
        }

        /// <summary>
        ///     Clears display units from the scatter chart X-axis.
        /// </summary>
        public PowerPointChart ClearScatterXAxisDisplayUnits() {
            if (!CanResolveScatterAxis(ResolveScatterXAxis)) {
                return this;
            }

            return ClearValueAxisDisplayUnits(axis => HasAxisPosition(axis, C.AxisPositionValues.Bottom));
        }

        /// <summary>
        ///     Sets display units for the scatter chart Y-axis.
        /// </summary>
        public PowerPointChart SetScatterYAxisDisplayUnits(C.BuiltInUnitValues unit, bool showLabel = true) {
            if (!CanResolveScatterAxis(ResolveScatterYAxis)) {
                return this;
            }

            return SetValueAxisDisplayUnitsCore(displayUnits => {
                displayUnits.RemoveAllChildren<C.CustomDisplayUnit>();
                displayUnits.RemoveAllChildren<C.BuiltInUnit>();
                displayUnits.Append(new C.BuiltInUnit { Val = unit });
            }, showLabel, null, axis => HasAxisPosition(axis, C.AxisPositionValues.Left));
        }

        /// <summary>
        ///     Sets display units for the scatter chart Y-axis with custom label text.
        /// </summary>
        public PowerPointChart SetScatterYAxisDisplayUnits(C.BuiltInUnitValues unit, string labelText, bool showLabel = true) {
            if (string.IsNullOrWhiteSpace(labelText)) {
                throw new ArgumentException("Label text cannot be empty.", nameof(labelText));
            }
            if (!CanResolveScatterAxis(ResolveScatterYAxis)) {
                return this;
            }

            return SetValueAxisDisplayUnitsCore(displayUnits => {
                displayUnits.RemoveAllChildren<C.CustomDisplayUnit>();
                displayUnits.RemoveAllChildren<C.BuiltInUnit>();
                displayUnits.Append(new C.BuiltInUnit { Val = unit });
            }, showLabel, labelText, axis => HasAxisPosition(axis, C.AxisPositionValues.Left));
        }

        /// <summary>
        ///     Sets custom display units for the scatter chart Y-axis.
        /// </summary>
        public PowerPointChart SetScatterYAxisDisplayUnits(double customUnit, bool showLabel = true) {
            if (!IsFinite(customUnit) || customUnit <= 0) {
                throw new ArgumentOutOfRangeException(nameof(customUnit));
            }
            if (!CanResolveScatterAxis(ResolveScatterYAxis)) {
                return this;
            }

            return SetValueAxisDisplayUnitsCore(displayUnits => {
                displayUnits.RemoveAllChildren<C.BuiltInUnit>();
                displayUnits.RemoveAllChildren<C.CustomDisplayUnit>();
                displayUnits.Append(new C.CustomDisplayUnit { Val = customUnit });
            }, showLabel, null, axis => HasAxisPosition(axis, C.AxisPositionValues.Left));
        }

        /// <summary>
        ///     Sets custom display units for the scatter chart Y-axis with custom label text.
        /// </summary>
        public PowerPointChart SetScatterYAxisDisplayUnits(double customUnit, string labelText, bool showLabel = true) {
            if (!IsFinite(customUnit) || customUnit <= 0) {
                throw new ArgumentOutOfRangeException(nameof(customUnit));
            }
            if (string.IsNullOrWhiteSpace(labelText)) {
                throw new ArgumentException("Label text cannot be empty.", nameof(labelText));
            }
            if (!CanResolveScatterAxis(ResolveScatterYAxis)) {
                return this;
            }

            return SetValueAxisDisplayUnitsCore(displayUnits => {
                displayUnits.RemoveAllChildren<C.BuiltInUnit>();
                displayUnits.RemoveAllChildren<C.CustomDisplayUnit>();
                displayUnits.Append(new C.CustomDisplayUnit { Val = customUnit });
            }, showLabel, labelText, axis => HasAxisPosition(axis, C.AxisPositionValues.Left));
        }

        /// <summary>
        ///     Clears display units from the scatter chart Y-axis.
        /// </summary>
        public PowerPointChart ClearScatterYAxisDisplayUnits() {
            if (!CanResolveScatterAxis(ResolveScatterYAxis)) {
                return this;
            }

            return ClearValueAxisDisplayUnits(axis => HasAxisPosition(axis, C.AxisPositionValues.Left));
        }

        /// <summary>
        ///     Sets display units for the value axis.
        /// </summary>
        public PowerPointChart SetValueAxisDisplayUnits(C.BuiltInUnitValues unit, bool showLabel = true) {
            return SetValueAxisDisplayUnitsCore(displayUnits => {
                displayUnits.RemoveAllChildren<C.CustomDisplayUnit>();
                displayUnits.RemoveAllChildren<C.BuiltInUnit>();
                displayUnits.Append(new C.BuiltInUnit { Val = unit });
            }, showLabel);
        }

        /// <summary>
        ///     Sets display units for the value axis with custom label text.
        /// </summary>
        public PowerPointChart SetValueAxisDisplayUnits(C.BuiltInUnitValues unit, string labelText, bool showLabel = true) {
            if (string.IsNullOrWhiteSpace(labelText)) {
                throw new ArgumentException("Label text cannot be empty.", nameof(labelText));
            }

            return SetValueAxisDisplayUnitsCore(displayUnits => {
                displayUnits.RemoveAllChildren<C.CustomDisplayUnit>();
                displayUnits.RemoveAllChildren<C.BuiltInUnit>();
                displayUnits.Append(new C.BuiltInUnit { Val = unit });
            }, showLabel, labelText);
        }

        /// <summary>
        ///     Sets custom display units for the value axis.
        /// </summary>
        public PowerPointChart SetValueAxisDisplayUnits(double customUnit, bool showLabel = true) {
            if (!IsFinite(customUnit) || customUnit <= 0) {
                throw new ArgumentOutOfRangeException(nameof(customUnit));
            }

            return SetValueAxisDisplayUnitsCore(displayUnits => {
                displayUnits.RemoveAllChildren<C.BuiltInUnit>();
                displayUnits.RemoveAllChildren<C.CustomDisplayUnit>();
                displayUnits.Append(new C.CustomDisplayUnit { Val = customUnit });
            }, showLabel);
        }

        /// <summary>
        ///     Sets custom display units for the value axis with custom label text.
        /// </summary>
        public PowerPointChart SetValueAxisDisplayUnits(double customUnit, string labelText, bool showLabel = true) {
            if (!IsFinite(customUnit) || customUnit <= 0) {
                throw new ArgumentOutOfRangeException(nameof(customUnit));
            }
            if (string.IsNullOrWhiteSpace(labelText)) {
                throw new ArgumentException("Label text cannot be empty.", nameof(labelText));
            }

            return SetValueAxisDisplayUnitsCore(displayUnits => {
                displayUnits.RemoveAllChildren<C.BuiltInUnit>();
                displayUnits.RemoveAllChildren<C.CustomDisplayUnit>();
                displayUnits.Append(new C.CustomDisplayUnit { Val = customUnit });
            }, showLabel, labelText);
        }

        /// <summary>
        ///     Clears display units from the value axis.
        /// </summary>
        public PowerPointChart ClearValueAxisDisplayUnits() {
            return ClearValueAxisDisplayUnits(null);
        }

        /// <summary>
        ///     Sets chart area fill/line styling.
        /// </summary>
        public PowerPointChart SetChartAreaStyle(string? fillColor = null, string? lineColor = null,
            double? lineWidthPoints = null, bool noFill = false, bool noLine = false) {
            ValidateAreaStyle(fillColor, lineColor, lineWidthPoints, noFill, noLine);

            ChartPart chartPart = GetChartPart();
            C.ChartSpace? chartSpace = chartPart.ChartSpace;
            if (chartSpace == null) {
                return this;
            }

            C.ShapeProperties props = chartSpace.GetFirstChild<C.ShapeProperties>() ?? new C.ShapeProperties();
            ApplyAreaStyle(props, fillColor, lineColor, lineWidthPoints, noFill, noLine);
            if (props.Parent == null) {
                InsertChartSpaceShapeProperties(chartSpace, props);
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Sets plot area fill/line styling.
        /// </summary>
        public PowerPointChart SetPlotAreaStyle(string? fillColor = null, string? lineColor = null,
            double? lineWidthPoints = null, bool noFill = false, bool noLine = false) {
            ValidateAreaStyle(fillColor, lineColor, lineWidthPoints, noFill, noLine);

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            C.ShapeProperties props = plotArea.GetFirstChild<C.ShapeProperties>() ?? new C.ShapeProperties();
            ApplyAreaStyle(props, fillColor, lineColor, lineWidthPoints, noFill, noLine);
            if (props.Parent == null) {
                InsertPlotAreaShapeProperties(plotArea, props);
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Adds or replaces a trendline for a chart series by index.
        /// </summary>
        public PowerPointChart SetSeriesTrendline(int seriesIndex, C.TrendlineValues type,
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
        ///     Adds or replaces a trendline for a chart series by name.
        /// </summary>
        public PowerPointChart SetSeriesTrendline(string seriesName, C.TrendlineValues type,
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
        ///     Removes trendlines from a chart series by index.
        /// </summary>
        public PowerPointChart ClearSeriesTrendline(int seriesIndex) {
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
        ///     Removes trendlines from a chart series by name.
        /// </summary>
        public PowerPointChart ClearSeriesTrendline(string seriesName, bool ignoreCase = true) {
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

        private PowerPointChart ClearValueAxisDisplayUnits(Func<C.ValueAxis, bool>? predicate) {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            C.ValueAxis? axis = predicate == null
                ? plotArea.Elements<C.ValueAxis>().FirstOrDefault()
                : plotArea.Elements<C.ValueAxis>().FirstOrDefault(predicate);
            if (axis == null) {
                return this;
            }

            axis.GetFirstChild<C.DisplayUnits>()?.Remove();
            Save();
            return this;
        }

        /// <summary>
        ///     Sets the category axis orientation (normal or reversed order).
        /// </summary>
        public PowerPointChart SetCategoryAxisReverseOrder(bool reverseOrder = true) {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            C.CategoryAxis? axis = plotArea.Elements<C.CategoryAxis>().FirstOrDefault();
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
        ///     Sets value axis scale parameters.
        /// </summary>
        public PowerPointChart SetValueAxisScale(double? minimum = null, double? maximum = null,
            double? majorUnit = null, double? minorUnit = null, double? logBase = null,
            bool? reverseOrder = null, bool? logScale = null) {
            ValidateAxisScale(minimum, maximum, majorUnit, minorUnit, logScale, logBase);

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            C.ValueAxis? axis = plotArea.Elements<C.ValueAxis>().FirstOrDefault();
            if (axis == null) {
                return this;
            }

            ApplyAxisScale(axis, minimum, maximum, majorUnit, minorUnit, reverseOrder, logScale, logBase);
            Save();
            return this;
        }

        /// <summary>
        ///     Sets where the category axis crosses the value axis.
        /// </summary>
        public PowerPointChart SetCategoryAxisCrossing(C.CrossesValues crosses, double? crossesAt = null) {
            if (crossesAt != null && (double.IsNaN(crossesAt.Value) || double.IsInfinity(crossesAt.Value))) {
                throw new ArgumentOutOfRangeException(nameof(crossesAt));
            }

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            C.CategoryAxis? axis = plotArea.Elements<C.CategoryAxis>().FirstOrDefault();
            if (axis == null) {
                return this;
            }

            ApplyAxisCrossing(axis, crosses, crossesAt);
            Save();
            return this;
        }

        /// <summary>
        ///     Sets where the value axis crosses the category axis.
        /// </summary>
        public PowerPointChart SetValueAxisCrossing(C.CrossesValues crosses, double? crossesAt = null) {
            if (crossesAt != null && (double.IsNaN(crossesAt.Value) || double.IsInfinity(crossesAt.Value))) {
                throw new ArgumentOutOfRangeException(nameof(crossesAt));
            }

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            C.ValueAxis? axis = plotArea.Elements<C.ValueAxis>().FirstOrDefault();
            if (axis == null) {
                return this;
            }

            ValidateCrossesAtForAxis(axis, crossesAt);
            ApplyAxisCrossing(axis, crosses, crossesAt);
            Save();
            return this;
        }

        /// <summary>
        ///     Sets scatter chart X-axis scale (bottom value axis).
        /// </summary>
        public PowerPointChart SetScatterXAxisScale(double? minimum = null, double? maximum = null,
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
        ///     Sets scatter chart Y-axis scale (left value axis).
        /// </summary>
        public PowerPointChart SetScatterYAxisScale(double? minimum = null, double? maximum = null,
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
        ///     Sets where the scatter X-axis crosses the Y-axis.
        /// </summary>
        public PowerPointChart SetScatterXAxisCrossing(C.CrossesValues? crosses = null, double? crossesAt = null) {
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
            ApplyAxisCrossing(axis, crosses ?? C.CrossesValues.AutoZero, crossesAt);
            Save();
            return this;
        }

        /// <summary>
        ///     Sets where the scatter Y-axis crosses the X-axis.
        /// </summary>
        public PowerPointChart SetScatterYAxisCrossing(C.CrossesValues? crosses = null, double? crossesAt = null) {
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
            ApplyAxisCrossing(axis, crosses ?? C.CrossesValues.AutoZero, crossesAt);
            Save();
            return this;
        }

        /// <summary>
        ///     Sets the fill color for a chart series by index.
        /// </summary>
        public PowerPointChart SetSeriesFillColor(int seriesIndex, string color) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            if (string.IsNullOrWhiteSpace(color)) {
                throw new ArgumentException("Series color cannot be null or empty.", nameof(color));
            }

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                C.ChartShapeProperties props = EnsureChartShapeProperties(series);
                ApplySolidFill(props, color);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Sets the fill color for a chart series by name.
        /// </summary>
        public PowerPointChart SetSeriesFillColor(string seriesName, string color, bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            if (string.IsNullOrWhiteSpace(color)) {
                throw new ArgumentException("Series color cannot be null or empty.", nameof(color));
            }

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                C.ChartShapeProperties props = EnsureChartShapeProperties(series);
                ApplySolidFill(props, color);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Sets the line color for a chart series by index.
        /// </summary>
        public PowerPointChart SetSeriesLineColor(int seriesIndex, string color, double? widthPoints = null) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            if (string.IsNullOrWhiteSpace(color)) {
                throw new ArgumentException("Series color cannot be null or empty.", nameof(color));
            }

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                C.ChartShapeProperties props = EnsureChartShapeProperties(series);
                ApplyLine(props, color, widthPoints);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Sets the line color for a chart series by name.
        /// </summary>
        public PowerPointChart SetSeriesLineColor(string seriesName, string color, double? widthPoints = null, bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            if (string.IsNullOrWhiteSpace(color)) {
                throw new ArgumentException("Series color cannot be null or empty.", nameof(color));
            }

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                C.ChartShapeProperties props = EnsureChartShapeProperties(series);
                ApplyLine(props, color, widthPoints);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Sets the marker style for a chart series by index.
        /// </summary>
        public PowerPointChart SetSeriesMarker(int seriesIndex, C.MarkerStyleValues style, int? size = null, string? fillColor = null, string? lineColor = null, double? lineWidthPoints = null) {
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
        ///     Sets the marker style for a chart series by name.
        /// </summary>
        public PowerPointChart SetSeriesMarker(string seriesName, C.MarkerStyleValues style, int? size = null, string? fillColor = null, string? lineColor = null, double? lineWidthPoints = null, bool ignoreCase = true) {
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

        private PowerPointChart SetAxisTitle<TAxis>(string title, Func<TAxis, bool>? predicate = null) where TAxis : OpenXmlCompositeElement {
            if (title == null) {
                throw new ArgumentNullException(nameof(title));
            }

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            TAxis? axis = predicate == null
                ? plotArea.Elements<TAxis>().FirstOrDefault()
                : plotArea.Elements<TAxis>().FirstOrDefault(predicate);
            if (axis == null) {
                return this;
            }

            axis.RemoveAllChildren<C.Title>();
            axis.Append(CreateAxisTitle(title));
            Save();
            return this;
        }

        private PowerPointChart SetAxisTitleTextStyle<TAxis>(double? fontSizePoints, bool? bold, bool? italic,
            string? color, string? fontName) where TAxis : OpenXmlCompositeElement {
            ValidateTextStyle(fontSizePoints, color, fontName);

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            TAxis? axis = plotArea.Elements<TAxis>().FirstOrDefault();
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

        private PowerPointChart SetAxisLabelTextStyle<TAxis>(double? fontSizePoints, bool? bold, bool? italic,
            string? color, string? fontName, Func<TAxis, bool>? predicate = null) where TAxis : OpenXmlCompositeElement {
            ValidateTextStyle(fontSizePoints, color, fontName);

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            TAxis? axis = predicate == null
                ? plotArea.Elements<TAxis>().FirstOrDefault()
                : plotArea.Elements<TAxis>().FirstOrDefault(predicate);
            if (axis == null) {
                return this;
            }

            ApplyTextStyle(EnsureTextPropertiesRunProperties(axis), fontSizePoints, bold, italic, color, fontName);
            Save();
            return this;
        }

        private PowerPointChart SetAxisLabelRotation<TAxis>(double rotationDegrees, Func<TAxis, bool>? predicate = null)
            where TAxis : OpenXmlCompositeElement {
            ValidateAxisLabelRotation(rotationDegrees);

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            TAxis? axis = predicate == null
                ? plotArea.Elements<TAxis>().FirstOrDefault()
                : plotArea.Elements<TAxis>().FirstOrDefault(predicate);
            if (axis == null) {
                return this;
            }

            EnsureTextPropertiesRunProperties(axis);
            C.TextProperties? textProps = axis.GetFirstChild<C.TextProperties>();
            if (textProps != null) {
                A.BodyProperties body = textProps.GetFirstChild<A.BodyProperties>() ?? new A.BodyProperties();
                body.Rotation = (int)Math.Round(rotationDegrees * 60000d);
                if (body.Parent == null) {
                    textProps.Append(body);
                }
            }

            Save();
            return this;
        }

        private PowerPointChart SetAxisTickLabelPosition<TAxis>(C.TickLabelPositionValues position, Func<TAxis, bool>? predicate = null)
            where TAxis : OpenXmlCompositeElement {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            TAxis? axis = predicate == null
                ? plotArea.Elements<TAxis>().FirstOrDefault()
                : plotArea.Elements<TAxis>().FirstOrDefault(predicate);
            if (axis == null) {
                return this;
            }

            ReplaceAxisChild(axis, new C.TickLabelPosition { Val = position });
            Save();
            return this;
        }

        private PowerPointChart SetAxisGridlines<TAxis>(bool showMajor, bool showMinor, string? lineColor,
            double? lineWidthPoints, Func<TAxis, bool>? predicate = null) where TAxis : OpenXmlCompositeElement {
            ValidateAxisGridlinesStyle(lineColor, lineWidthPoints);

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            TAxis? axis = predicate == null
                ? plotArea.Elements<TAxis>().FirstOrDefault()
                : plotArea.Elements<TAxis>().FirstOrDefault(predicate);
            if (axis == null) {
                return this;
            }

            ApplyGridlines(axis, showMajor, showMinor, lineColor, lineWidthPoints);
            Save();
            return this;
        }

        private PowerPointChart SetValueAxisDisplayUnitsCore(Action<C.DisplayUnits> configureUnits, bool showLabel,
            string? labelText = null, Func<C.ValueAxis, bool>? predicate = null) {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            C.ValueAxis? axis = predicate == null
                ? plotArea.Elements<C.ValueAxis>().FirstOrDefault()
                : plotArea.Elements<C.ValueAxis>().FirstOrDefault(predicate);
            if (axis == null) {
                return this;
            }

            C.DisplayUnits displayUnits = axis.GetFirstChild<C.DisplayUnits>() ?? new C.DisplayUnits();
            configureUnits(displayUnits);
            ApplyDisplayUnitsLabel(displayUnits, showLabel, labelText);
            if (displayUnits.Parent == null) {
                OpenXmlElement? insertBefore = axis.GetFirstChild<C.ExtensionList>();
                if (insertBefore != null) {
                    axis.InsertBefore(displayUnits, insertBefore);
                } else {
                    axis.Append(displayUnits);
                }
            }

            Save();
            return this;
        }

        private PowerPointChart SetAxisNumberFormat<TAxis>(string formatCode, bool sourceLinked, Func<TAxis, bool>? predicate = null)
            where TAxis : OpenXmlCompositeElement {
            if (string.IsNullOrWhiteSpace(formatCode)) {
                throw new ArgumentException("Format code cannot be null or empty.", nameof(formatCode));
            }

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            TAxis? axis = predicate == null
                ? plotArea.Elements<TAxis>().FirstOrDefault()
                : plotArea.Elements<TAxis>().FirstOrDefault(predicate);
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

        private static bool HasAxisPosition(C.ValueAxis axis, C.AxisPositionValues position) {
            return axis.GetFirstChild<C.AxisPosition>()?.Val?.Value == position;
        }

        private bool CanResolveScatterAxis(Func<C.PlotArea, C.ValueAxis?> resolver) {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return false;
            }

            return resolver(plotArea) != null;
        }

        private static C.ValueAxis? ResolveScatterXAxis(C.PlotArea plotArea) {
            if (plotArea.Elements<C.CategoryAxis>().Any()) {
                return null;
            }

            return plotArea.Elements<C.ValueAxis>()
                .FirstOrDefault(axis => HasAxisPosition(axis, C.AxisPositionValues.Bottom));
        }

        private static C.ValueAxis? ResolveScatterYAxis(C.PlotArea plotArea) {
            if (plotArea.Elements<C.CategoryAxis>().Any()) {
                return null;
            }

            return plotArea.Elements<C.ValueAxis>()
                .FirstOrDefault(axis => HasAxisPosition(axis, C.AxisPositionValues.Left));
        }

        private static void ValidateAxisScale(double? minimum, double? maximum, double? majorUnit, double? minorUnit,
            bool? logScale, double? logBase) {
            if (minimum != null && !IsFinite(minimum.Value)) {
                throw new ArgumentOutOfRangeException(nameof(minimum));
            }
            if (maximum != null && !IsFinite(maximum.Value)) {
                throw new ArgumentOutOfRangeException(nameof(maximum));
            }
            if (minimum != null && maximum != null && minimum.Value >= maximum.Value) {
                throw new ArgumentException("Minimum must be less than maximum.");
            }
            if (majorUnit != null && (!IsFinite(majorUnit.Value) || majorUnit.Value <= 0)) {
                throw new ArgumentOutOfRangeException(nameof(majorUnit));
            }
            if (minorUnit != null && (!IsFinite(minorUnit.Value) || minorUnit.Value <= 0)) {
                throw new ArgumentOutOfRangeException(nameof(minorUnit));
            }
            if (logScale == false && logBase != null) {
                throw new ArgumentException("Log base requires logScale to be enabled.", nameof(logBase));
            }

            bool effectiveLog = logScale == true || logBase != null;
            if (effectiveLog) {
                double baseValue = logBase ?? 10d;
                if (!IsFinite(baseValue) || baseValue <= 1d) {
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

        private static bool IsFinite(double value) {
            return !double.IsNaN(value) && !double.IsInfinity(value);
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
                ValidateEffectiveAxisScale(scaling, minimum, maximum, logScale, logBase);
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

                NormalizeScalingOrder(scaling);
            }

            if (majorUnit != null) {
                ReplaceChild(axis, new C.MajorUnit { Val = majorUnit.Value });
            }
            if (minorUnit != null) {
                ReplaceChild(axis, new C.MinorUnit { Val = minorUnit.Value });
            }
        }

        private static void ValidateEffectiveAxisScale(C.Scaling scaling, double? minimum, double? maximum, bool? logScale, double? logBase) {
            double? effectiveMinimum = minimum ?? scaling.GetFirstChild<C.MinAxisValue>()?.Val?.Value;
            double? effectiveMaximum = maximum ?? scaling.GetFirstChild<C.MaxAxisValue>()?.Val?.Value;
            if (effectiveMinimum != null && effectiveMaximum != null && effectiveMinimum.Value >= effectiveMaximum.Value) {
                throw new ArgumentException("Minimum must be less than maximum.");
            }

            bool effectiveLog = logScale == true || logBase != null;
            if (!effectiveLog && logScale != false) {
                effectiveLog = scaling.GetFirstChild<C.LogBase>() != null;
            }

            if (!effectiveLog) {
                return;
            }

            if (effectiveMinimum != null && effectiveMinimum.Value <= 0) {
                throw new ArgumentException("Minimum must be greater than 0 for log scale.", nameof(minimum));
            }
            if (effectiveMaximum != null && effectiveMaximum.Value <= 0) {
                throw new ArgumentException("Maximum must be greater than 0 for log scale.", nameof(maximum));
            }
        }

        private static C.Scaling EnsureScaling(OpenXmlCompositeElement axis) {
            C.Scaling scaling = axis.GetFirstChild<C.Scaling>() ?? new C.Scaling();
            if (scaling.Parent == null) {
                C.AxisId? axisId = axis.GetFirstChild<C.AxisId>();
                if (axisId != null) {
                    axis.InsertAfter(scaling, axisId);
                } else {
                    axis.PrependChild(scaling);
                }
            }

            if (scaling.GetFirstChild<C.Orientation>() == null) {
                scaling.PrependChild(new C.Orientation { Val = C.OrientationValues.MinMax });
            }

            return scaling;
        }

        private static void NormalizeScalingOrder(C.Scaling scaling) {
            C.Orientation? orientation = scaling.GetFirstChild<C.Orientation>();
            C.MaxAxisValue? maxAxisValue = scaling.GetFirstChild<C.MaxAxisValue>();
            C.MinAxisValue? minAxisValue = scaling.GetFirstChild<C.MinAxisValue>();
            C.LogBase? logBase = scaling.GetFirstChild<C.LogBase>();

            orientation?.Remove();
            maxAxisValue?.Remove();
            minAxisValue?.Remove();
            logBase?.Remove();

            if (logBase != null) {
                scaling.Append(logBase);
            }
            if (orientation != null) {
                scaling.Append(orientation);
            }
            if (maxAxisValue != null) {
                scaling.Append(maxAxisValue);
            }
            if (minAxisValue != null) {
                scaling.Append(minAxisValue);
            }
        }

        private static void ApplyAxisCrossing(OpenXmlCompositeElement axis, C.CrossesValues crosses, double? crossesAt) {
            axis.GetFirstChild<C.Crosses>()?.Remove();
            axis.GetFirstChild<C.CrossesAt>()?.Remove();

            OpenXmlElement crossing = crossesAt != null
                ? new C.CrossesAt { Val = crossesAt.Value }
                : new C.Crosses { Val = crosses };

            C.CrossingAxis? crossAxis = axis.GetFirstChild<C.CrossingAxis>();
            if (crossAxis != null) {
                axis.InsertAfter(crossing, crossAxis);
            } else {
                axis.Append(crossing);
            }
        }

        private static void ApplyDataLabels(OpenXmlCompositeElement chartElement, bool showLegendKey, bool showValue,
            bool showCategoryName, bool showSeriesName, bool showPercent) {
            C.DataLabels labels = EnsureDataLabels(chartElement);
            ReplaceChild(labels, new C.ShowLegendKey { Val = showLegendKey });
            ReplaceChild(labels, new C.ShowValue { Val = showValue });
            ReplaceChild(labels, new C.ShowCategoryName { Val = showCategoryName });
            ReplaceChild(labels, new C.ShowSeriesName { Val = showSeriesName });
            ReplaceChild(labels, new C.ShowPercent { Val = showPercent });
            ReplaceChild(labels, new C.ShowBubbleSize { Val = false });
            NormalizeDataLabelsOrder(labels);
        }

        private static C.DataLabels EnsureDataLabels(OpenXmlCompositeElement chartElement) {
            C.DataLabels labels = chartElement.GetFirstChild<C.DataLabels>() ?? new C.DataLabels();
            if (labels.Parent == null) {
                chartElement.Append(labels);
            }

            return labels;
        }

        private static void NormalizeDataLabelsOrder(C.DataLabels labels) {
            List<C.DataLabel> overrides = labels.Elements<C.DataLabel>().ToList();
            C.Delete? delete = labels.GetFirstChild<C.Delete>();
            C.NumberingFormat? numFmt = labels.GetFirstChild<C.NumberingFormat>();
            C.ChartShapeProperties? shapeProps = labels.GetFirstChild<C.ChartShapeProperties>();
            C.TextProperties? textProps = labels.GetFirstChild<C.TextProperties>();
            C.DataLabelPosition? position = labels.GetFirstChild<C.DataLabelPosition>();
            C.ShowLegendKey? showLegendKey = labels.GetFirstChild<C.ShowLegendKey>();
            C.ShowValue? showValue = labels.GetFirstChild<C.ShowValue>();
            C.ShowCategoryName? showCategoryName = labels.GetFirstChild<C.ShowCategoryName>();
            C.ShowSeriesName? showSeriesName = labels.GetFirstChild<C.ShowSeriesName>();
            C.ShowPercent? showPercent = labels.GetFirstChild<C.ShowPercent>();
            C.ShowBubbleSize? showBubbleSize = labels.GetFirstChild<C.ShowBubbleSize>();
            C.Separator? separator = labels.GetFirstChild<C.Separator>();
            C.ShowLeaderLines? showLeaderLines = labels.GetFirstChild<C.ShowLeaderLines>();
            C.LeaderLines? leaderLines = labels.GetFirstChild<C.LeaderLines>();
            C.ExtensionList? extLst = labels.GetFirstChild<C.ExtensionList>();

            List<OpenXmlElement> otherChildren = labels.ChildElements
                .Where(child => child is not C.DataLabel
                                && child is not C.Delete
                                && child is not C.NumberingFormat
                                && child is not C.ChartShapeProperties
                                && child is not C.TextProperties
                                && child is not C.DataLabelPosition
                                && child is not C.ShowLegendKey
                                && child is not C.ShowValue
                                && child is not C.ShowCategoryName
                                && child is not C.ShowSeriesName
                                && child is not C.ShowPercent
                                && child is not C.ShowBubbleSize
                                && child is not C.Separator
                                && child is not C.ShowLeaderLines
                                && child is not C.LeaderLines
                                && child is not C.ExtensionList)
                .ToList();

            labels.RemoveAllChildren();

            if (delete != null) {
                labels.Append(delete);
            }
            if (numFmt != null) {
                labels.Append(numFmt);
            }
            if (shapeProps != null) {
                labels.Append(shapeProps);
            }
            if (textProps != null) {
                labels.Append(textProps);
            }
            if (position != null) {
                labels.Append(position);
            }
            if (showLegendKey != null) {
                labels.Append(showLegendKey);
            }
            if (showValue != null) {
                labels.Append(showValue);
            }
            if (showCategoryName != null) {
                labels.Append(showCategoryName);
            }
            if (showSeriesName != null) {
                labels.Append(showSeriesName);
            }
            if (showPercent != null) {
                labels.Append(showPercent);
            }
            if (showBubbleSize != null) {
                labels.Append(showBubbleSize);
            }
            if (separator != null) {
                labels.Append(separator);
            }
            if (showLeaderLines != null) {
                labels.Append(showLeaderLines);
            }
            if (leaderLines != null) {
                labels.Append(leaderLines);
            }

            foreach (C.DataLabel dataLabel in overrides) {
                labels.Append(dataLabel);
            }

            foreach (OpenXmlElement otherChild in otherChildren) {
                labels.Append(otherChild);
            }

            if (extLst != null) {
                labels.Append(extLst);
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
            A.LatinFont? existingLatinFont = runProps.GetFirstChild<A.LatinFont>()?.CloneNode(true) as A.LatinFont;
            if (color != null) {
                runProps.RemoveAllChildren<A.LatinFont>();
                ApplySolidFill(runProps, color);
            }
            if (fontName != null) {
                runProps.RemoveAllChildren<A.LatinFont>();
                runProps.Append(new A.LatinFont { Typeface = fontName });
            } else if (color != null && existingLatinFont != null) {
                runProps.Append(existingLatinFont);
            }
        }

        private static void ValidateTextStyle(double? fontSizePoints, string? color, string? fontName) {
            if (fontSizePoints != null && fontSizePoints <= 0) {
                throw new ArgumentOutOfRangeException(nameof(fontSizePoints));
            }
            if (color != null && string.IsNullOrWhiteSpace(color)) {
                throw new ArgumentException("Color cannot be empty.", nameof(color));
            }
            if (fontName != null && string.IsNullOrWhiteSpace(fontName)) {
                throw new ArgumentException("Font name cannot be empty.", nameof(fontName));
            }
        }

        private static void ValidateAxisLabelRotation(double rotationDegrees) {
            if (double.IsNaN(rotationDegrees) || double.IsInfinity(rotationDegrees)) {
                throw new ArgumentOutOfRangeException(nameof(rotationDegrees));
            }
            if (rotationDegrees < -90d || rotationDegrees > 90d) {
                throw new ArgumentOutOfRangeException(nameof(rotationDegrees), "Rotation must be between -90 and 90 degrees.");
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

        private static void ValidateAreaStyle(string? fillColor, string? lineColor, double? lineWidthPoints,
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
            if (color == null && widthPoints == null) {
                return;
            }

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

        private static void ApplyGridlines(OpenXmlCompositeElement axis, bool showMajor, bool showMinor,
            string? lineColor, double? lineWidthPoints) {
            ApplyGridline<C.MajorGridlines>(axis, showMajor, lineColor, lineWidthPoints);
            ApplyGridline<C.MinorGridlines>(axis, showMinor, lineColor, lineWidthPoints);
        }

        private static void ApplyAreaStyle(OpenXmlCompositeElement props, string? fillColor, string? lineColor,
            double? lineWidthPoints, bool noFill, bool noLine) {
            if (noFill) {
                ApplyNoFill(props);
            } else if (fillColor != null) {
                ApplySolidFill(props, fillColor);
            }

            if (noLine) {
                ApplyNoLine(props);
            } else if (lineColor != null || lineWidthPoints != null) {
                ApplyOptionalLine(props, lineColor, lineWidthPoints);
            }
        }

        private static void ApplyGridline<TGridlines>(OpenXmlCompositeElement axis, bool show,
            string? lineColor, double? lineWidthPoints) where TGridlines : OpenXmlCompositeElement, new() {
            TGridlines? gridlines = axis.GetFirstChild<TGridlines>();
            if (!show) {
                gridlines?.Remove();
                return;
            }

            gridlines ??= new TGridlines();
            if (lineColor != null || lineWidthPoints != null) {
                C.ChartShapeProperties props = gridlines.GetFirstChild<C.ChartShapeProperties>() ?? new C.ChartShapeProperties();
                ApplyOptionalLine(props, lineColor, lineWidthPoints);
                if (props.Parent == null) {
                    gridlines.Append(props);
                }
            }

            if (gridlines.Parent == null) {
                InsertAxisGridlines(axis, gridlines);
            }
        }

        private static void ApplyTrendline(OpenXmlCompositeElement series, C.TrendlineValues type, int? order, int? period,
            double? forward, double? backward, double? intercept, bool displayEquation, bool displayRSquared,
            string? lineColor, double? lineWidthPoints) {
            if (!IsTrendlineSupportedSeries(series)) {
                throw new InvalidOperationException("Trendlines are only supported for line, bar/column, area, and scatter series.");
            }

            series.RemoveAllChildren<C.Trendline>();
            C.Trendline trendline = new C.Trendline();

            if (lineColor != null || lineWidthPoints != null) {
                C.ChartShapeProperties props = new C.ChartShapeProperties();
                ApplyOptionalLine(props, lineColor, lineWidthPoints);
                trendline.Append(props);
            }

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
            if (displayRSquared) {
                trendline.Append(new C.DisplayRSquaredValue { Val = displayRSquared });
            }
            if (displayEquation) {
                trendline.Append(new C.DisplayEquation { Val = displayEquation });
            }

            InsertTrendline(series, trendline);
        }

        private static void InsertChartSpaceShapeProperties(C.ChartSpace chartSpace, C.ShapeProperties props) {
            OpenXmlElement? insertBefore = chartSpace.GetFirstChild<C.TextProperties>();
            insertBefore ??= chartSpace.GetFirstChild<C.ExternalData>();
            insertBefore ??= chartSpace.GetFirstChild<C.PrintSettings>();
            insertBefore ??= chartSpace.GetFirstChild<C.UserShapesReference>();
            insertBefore ??= chartSpace.GetFirstChild<C.ExtensionList>();

            if (insertBefore != null) {
                chartSpace.InsertBefore(props, insertBefore);
            } else {
                chartSpace.Append(props);
            }
        }

        private static void InsertPlotAreaShapeProperties(C.PlotArea plotArea, C.ShapeProperties props) {
            OpenXmlElement? insertBefore = plotArea.GetFirstChild<C.ExtensionList>();
            if (insertBefore != null) {
                plotArea.InsertBefore(props, insertBefore);
            } else {
                plotArea.Append(props);
            }
        }

        private static bool IsTrendlineSupportedSeries(OpenXmlCompositeElement series) {
            return series is C.LineChartSeries
                   || series is C.BarChartSeries
                   || series is C.AreaChartSeries
                   || series is C.ScatterChartSeries;
        }

        private static void InsertTrendline(OpenXmlCompositeElement series, C.Trendline trendline) {
            OpenXmlElement? insertBefore = series.GetFirstChild<C.ErrorBars>();
            insertBefore ??= series.GetFirstChild<C.CategoryAxisData>();
            insertBefore ??= series.GetFirstChild<C.Values>();
            insertBefore ??= series.GetFirstChild<C.XValues>();
            insertBefore ??= series.GetFirstChild<C.YValues>();
            insertBefore ??= series.GetFirstChild<C.BubbleSize>();
            insertBefore ??= series.GetFirstChild<C.Smooth>();
            insertBefore ??= series.GetFirstChild<C.ExtensionList>();

            if (insertBefore != null) {
                series.InsertBefore(trendline, insertBefore);
            } else {
                series.Append(trendline);
            }
        }

        private static void InsertSeriesMarker(OpenXmlCompositeElement series, C.Marker marker) {
            OpenXmlElement? insertBefore = series.GetFirstChild<C.DataPoint>();
            insertBefore ??= series.GetFirstChild<C.DataLabels>();
            insertBefore ??= series.GetFirstChild<C.Trendline>();
            insertBefore ??= series.GetFirstChild<C.ErrorBars>();
            insertBefore ??= series.GetFirstChild<C.CategoryAxisData>();
            insertBefore ??= series.GetFirstChild<C.Values>();
            insertBefore ??= series.GetFirstChild<C.XValues>();
            insertBefore ??= series.GetFirstChild<C.YValues>();
            insertBefore ??= series.GetFirstChild<C.BubbleSize>();
            insertBefore ??= series.GetFirstChild<C.Smooth>();
            insertBefore ??= series.GetFirstChild<C.ExtensionList>();

            if (insertBefore != null) {
                series.InsertBefore(marker, insertBefore);
            } else {
                series.Append(marker);
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
                if (parent is C.CategoryAxis || parent is C.ValueAxis) {
                    OpenXmlElement? insertBefore = parent.GetFirstChild<C.CrossingAxis>();
                    insertBefore ??= parent.GetFirstChild<C.Crosses>();
                    insertBefore ??= parent.GetFirstChild<C.CrossesAt>();
                    insertBefore ??= parent.GetFirstChild<C.AutoLabeled>();
                    insertBefore ??= parent.GetFirstChild<C.LabelAlignment>();
                    insertBefore ??= parent.GetFirstChild<C.LabelOffset>();
                    insertBefore ??= parent.GetFirstChild<C.NoMultiLevelLabels>();
                    insertBefore ??= parent.GetFirstChild<C.CrossBetween>();
                    insertBefore ??= parent.GetFirstChild<C.MajorUnit>();
                    insertBefore ??= parent.GetFirstChild<C.MinorUnit>();
                    insertBefore ??= parent.GetFirstChild<C.DisplayUnits>();
                    insertBefore ??= parent.GetFirstChild<C.ExtensionList>();
                    if (insertBefore != null) {
                        parent.InsertBefore(textProps, insertBefore);
                    } else {
                        parent.Append(textProps);
                    }
                } else {
                    parent.Append(textProps);
                }
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
                run.InsertAt(runProps, 0);
            } else if (runProps != run.FirstChild) {
                runProps.Remove();
                run.InsertAt(runProps, 0);
            }

            if (richText.Parent == null) {
                chartText.Append(richText);
            }

            return runProps;
        }

        private static void ApplyMarker(C.Marker marker, C.MarkerStyleValues style, int? size, string? fillColor, string? lineColor, double? lineWidthPoints) {
            marker.Symbol = new C.Symbol { Val = style };
            if (size != null) {
                marker.Size = new C.Size { Val = (byte)size.Value };
            }

            if (fillColor != null || lineColor != null || lineWidthPoints != null) {
                C.ChartShapeProperties props = marker.ChartShapeProperties ?? new C.ChartShapeProperties();
                if (fillColor != null) {
                    ApplySolidFill(props, fillColor);
                }
                if (lineColor != null || lineWidthPoints != null) {
                    ApplyOptionalLine(props, lineColor, lineWidthPoints);
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

        private static bool ApplySeriesMarkerByIndex<TChart>(IEnumerable<TChart> charts, int seriesIndex, Action<C.Marker> apply) where TChart : OpenXmlCompositeElement {
            foreach (TChart chart in charts) {
                List<OpenXmlCompositeElement> series = chart.ChildElements
                    .OfType<OpenXmlCompositeElement>()
                    .Where(IsSeriesElement)
                    .OrderBy(GetSeriesIndex)
                    .ToList();

                if (seriesIndex < 0 || seriesIndex >= series.Count) {
                    continue;
                }

                OpenXmlCompositeElement seriesElement = series[seriesIndex];
                C.Marker marker = seriesElement.GetFirstChild<C.Marker>() ?? new C.Marker();
                apply(marker);
                if (marker.Parent == null) {
                    InsertSeriesMarker(seriesElement, marker);
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
                            InsertSeriesMarker(series, marker);
                        }
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool ApplySeriesByIndex<TChart>(IEnumerable<TChart> charts, int seriesIndex,
            Action<OpenXmlCompositeElement> apply) where TChart : OpenXmlCompositeElement {
            foreach (TChart chart in charts) {
                List<OpenXmlCompositeElement> series = chart.ChildElements
                    .OfType<OpenXmlCompositeElement>()
                    .Where(IsSeriesElement)
                    .OrderBy(GetSeriesIndex)
                    .ToList();

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

        private static bool IsSeriesElement(OpenXmlCompositeElement element) {
            return element is C.BarChartSeries ||
                   element is C.LineChartSeries ||
                   element is C.AreaChartSeries ||
                   element is C.PieChartSeries ||
                   element is C.ScatterChartSeries;
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

        private static void ReplaceAxisChild<T>(OpenXmlCompositeElement axis, T child) where T : OpenXmlElement {
            axis.GetFirstChild<T>()?.Remove();

            OpenXmlElement? insertBefore = axis.GetFirstChild<C.ShapeProperties>();
            insertBefore ??= axis.GetFirstChild<C.TextProperties>();
            insertBefore ??= axis.GetFirstChild<C.CrossingAxis>();
            insertBefore ??= axis.GetFirstChild<C.Crosses>();
            insertBefore ??= axis.GetFirstChild<C.CrossesAt>();
            insertBefore ??= axis.GetFirstChild<C.AutoLabeled>();
            insertBefore ??= axis.GetFirstChild<C.LabelAlignment>();
            insertBefore ??= axis.GetFirstChild<C.LabelOffset>();
            insertBefore ??= axis.GetFirstChild<C.NoMultiLevelLabels>();
            insertBefore ??= axis.GetFirstChild<C.CrossBetween>();
            insertBefore ??= axis.GetFirstChild<C.MajorUnit>();
            insertBefore ??= axis.GetFirstChild<C.MinorUnit>();
            insertBefore ??= axis.GetFirstChild<C.DisplayUnits>();
            insertBefore ??= axis.GetFirstChild<C.ExtensionList>();

            if (insertBefore != null) {
                axis.InsertBefore(child, insertBefore);
            } else {
                axis.Append(child);
            }
        }

        private static void InsertAxisGridlines<TGridlines>(OpenXmlCompositeElement axis, TGridlines gridlines)
            where TGridlines : OpenXmlCompositeElement {
            OpenXmlElement? insertBefore = typeof(TGridlines) == typeof(C.MajorGridlines)
                ? axis.GetFirstChild<C.MinorGridlines>()
                : null;
            insertBefore ??= axis.GetFirstChild<C.Title>();
            insertBefore ??= axis.GetFirstChild<C.NumberingFormat>();
            insertBefore ??= axis.GetFirstChild<C.MajorTickMark>();
            insertBefore ??= axis.GetFirstChild<C.MinorTickMark>();
            insertBefore ??= axis.GetFirstChild<C.TickLabelPosition>();
            insertBefore ??= axis.GetFirstChild<C.ShapeProperties>();
            insertBefore ??= axis.GetFirstChild<C.TextProperties>();
            insertBefore ??= axis.GetFirstChild<C.CrossingAxis>();
            insertBefore ??= axis.GetFirstChild<C.Crosses>();
            insertBefore ??= axis.GetFirstChild<C.CrossesAt>();
            insertBefore ??= axis.GetFirstChild<C.AutoLabeled>();
            insertBefore ??= axis.GetFirstChild<C.LabelAlignment>();
            insertBefore ??= axis.GetFirstChild<C.LabelOffset>();
            insertBefore ??= axis.GetFirstChild<C.NoMultiLevelLabels>();
            insertBefore ??= axis.GetFirstChild<C.CrossBetween>();
            insertBefore ??= axis.GetFirstChild<C.MajorUnit>();
            insertBefore ??= axis.GetFirstChild<C.MinorUnit>();
            insertBefore ??= axis.GetFirstChild<C.DisplayUnits>();
            insertBefore ??= axis.GetFirstChild<C.ExtensionList>();

            if (insertBefore != null) {
                axis.InsertBefore(gridlines, insertBefore);
            } else {
                axis.Append(gridlines);
            }
        }

        private static void ReplaceValueAxisCrossBetween(C.ValueAxis axis, C.CrossBetween crossBetween) {
            axis.GetFirstChild<C.CrossBetween>()?.Remove();

            OpenXmlElement? insertBefore = axis.GetFirstChild<C.MajorUnit>();
            insertBefore ??= axis.GetFirstChild<C.MinorUnit>();
            insertBefore ??= axis.GetFirstChild<C.DisplayUnits>();
            insertBefore ??= axis.GetFirstChild<C.ExtensionList>();

            if (insertBefore != null) {
                axis.InsertBefore(crossBetween, insertBefore);
            } else {
                axis.Append(crossBetween);
            }
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
            C.ChartReference? chartReference = Frame.Graphic?.GraphicData?.GetFirstChild<C.ChartReference>();
            StringValue? relationshipId = chartReference?.Id;
            if (relationshipId == null) {
                throw new InvalidOperationException("Chart reference not found for the shape.");
            }

            string relId = relationshipId.Value ?? throw new InvalidOperationException("Chart relationship id is empty.");
            return (ChartPart)_slidePart.GetPartById(relId);
        }

        private void Save() {
            ChartPart chartPart = GetChartPart();
            chartPart.ChartSpace?.Save();
        }
    }
}
