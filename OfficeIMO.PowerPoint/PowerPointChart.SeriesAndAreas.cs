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
        public partial class PowerPointChart : PowerPointShape {
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

        private PowerPointChart ApplyToAllDataLabels(Action<C.DataLabels> apply) {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            foreach (C.BarChart barChart in plotArea.Elements<C.BarChart>()) {
                apply(EnsureDataLabels(barChart));
            }

            foreach (C.LineChart lineChart in plotArea.Elements<C.LineChart>()) {
                apply(EnsureDataLabels(lineChart));
            }

            foreach (C.AreaChart areaChart in plotArea.Elements<C.AreaChart>()) {
                apply(EnsureDataLabels(areaChart));
            }

            foreach (C.PieChart pieChart in plotArea.Elements<C.PieChart>()) {
                apply(EnsureDataLabels(pieChart));
            }

            foreach (C.DoughnutChart doughnutChart in plotArea.Elements<C.DoughnutChart>()) {
                apply(EnsureDataLabels(doughnutChart));
            }

            foreach (C.ScatterChart scatterChart in plotArea.Elements<C.ScatterChart>()) {
                apply(EnsureDataLabels(scatterChart));
            }

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

    }
}
