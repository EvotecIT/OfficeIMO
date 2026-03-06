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

        private PowerPointChart SetAxisTitle<TAxis>(string title) where TAxis : OpenXmlCompositeElement {
            if (title == null) {
                throw new ArgumentNullException(nameof(title));
            }

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            TAxis? axis = plotArea.Elements<TAxis>().FirstOrDefault();
            if (axis == null) {
                return this;
            }

            axis.RemoveAllChildren<C.Title>();
            axis.Append(CreateAxisTitle(title));
            Save();
            return this;
        }

        private PowerPointChart SetAxisNumberFormat<TAxis>(string formatCode, bool sourceLinked)
            where TAxis : OpenXmlCompositeElement {
            if (string.IsNullOrWhiteSpace(formatCode)) {
                throw new ArgumentException("Format code cannot be null or empty.", nameof(formatCode));
            }

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            TAxis? axis = plotArea.Elements<TAxis>().FirstOrDefault();
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

                NormalizeScalingOrder(scaling);
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
            if (fontName != null) {
                runProps.RemoveAllChildren<A.LatinFont>();
                runProps.Append(new A.LatinFont { Typeface = fontName });
            }
            if (color != null) {
                ApplySolidFill(runProps, color);
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
