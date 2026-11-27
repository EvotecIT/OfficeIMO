using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
        //     Represents a chart on a slide.
        public partial class PowerPointChart : PowerPointShape {
        /// <summary>
        ///     Sets chart area fill/line styling.
        /// </summary>
        public PowerPointChart SetChartAreaStyle(OfficeColor? fillColor = null, OfficeColor? lineColor = null,
            OfficeLength? lineWidth = null, bool noFill = false, bool noLine = false) {
            ValidateAreaStyle(fillColor, lineColor, lineWidth, noFill, noLine);

            OfficeXmlNode chartSpace = ChartSpaceNode;
            if (chartSpace.IsNone) {
                return this;
            }

            OfficeXmlNode props = chartSpace.Child(CNs, "spPr");
            if (props.IsNone) {
                props = InsertChartSpaceShapeProperties(chartSpace, chartSpace.CreateElement(CNs, "spPr"));
            }
            ApplyAreaStyle(props, fillColor, lineColor, lineWidth, noFill, noLine);

            MarkChartEdited();
            return this;
        }

        /// <summary>
        ///     Sets plot area fill/line styling.
        /// </summary>
        public PowerPointChart SetPlotAreaStyle(OfficeColor? fillColor = null, OfficeColor? lineColor = null,
            OfficeLength? lineWidth = null, bool noFill = false, bool noLine = false) {
            ValidateAreaStyle(fillColor, lineColor, lineWidth, noFill, noLine);

            OfficeXmlNode plotArea = PlotAreaNode;
            if (plotArea.IsNone) {
                return this;
            }

            OfficeXmlNode props = plotArea.Child(CNs, "spPr");
            if (props.IsNone) {
                props = InsertPlotAreaShapeProperties(plotArea, plotArea.CreateElement(CNs, "spPr"));
            }
            ApplyAreaStyle(props, fillColor, lineColor, lineWidth, noFill, noLine);

            MarkChartEdited();
            return this;
        }

        /// <summary>
        ///     Adds or replaces a trendline for a chart series by index.
        /// </summary>
#pragma warning disable RS0026 // Existing overload family; only the color type or font-family parameter names changed.
        public PowerPointChart SetSeriesTrendline(int seriesIndex, OfficeChartTrendlineType type,
            int? order = null, int? period = null, double? forward = null, double? backward = null, double? intercept = null,
            bool displayEquation = false, bool displayRSquared = false, OfficeColor? lineColor = null, OfficeLength? lineWidth = null) {
#pragma warning restore RS0026
            ArgumentOutOfRangeException.ThrowIfNegative(seriesIndex);
            ValidateTrendline(type, order, period, forward, backward, lineColor, lineWidth);

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                ApplyTrendline(series, type, order, period, forward, backward, intercept, displayEquation, displayRSquared, lineColor, lineWidth);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            MarkChartEdited();
            return this;
        }

        /// <summary>
        ///     Adds or replaces a trendline for a chart series by name.
        /// </summary>
#pragma warning disable RS0026 // Existing overload family; only the color type or font-family parameter names changed.
        public PowerPointChart SetSeriesTrendline(string seriesName, OfficeChartTrendlineType type,
            int? order = null, int? period = null, double? forward = null, double? backward = null, double? intercept = null,
            bool displayEquation = false, bool displayRSquared = false, OfficeColor? lineColor = null, OfficeLength? lineWidth = null,
            bool ignoreCase = true) {
#pragma warning restore RS0026
            ArgumentNullException.ThrowIfNull(seriesName);
            ValidateTrendline(type, order, period, forward, backward, lineColor, lineWidth);

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                ApplyTrendline(series, type, order, period, forward, backward, intercept, displayEquation, displayRSquared, lineColor, lineWidth);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            MarkChartEdited();
            return this;
        }

        /// <summary>
        ///     Removes trendlines from a chart series by index.
        /// </summary>
        public PowerPointChart ClearSeriesTrendline(int seriesIndex) {
            ArgumentOutOfRangeException.ThrowIfNegative(seriesIndex);

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                series.RemoveChildren(CNs, "trendline");
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            MarkChartEdited();
            return this;
        }

        /// <summary>
        ///     Removes trendlines from a chart series by name.
        /// </summary>
        public PowerPointChart ClearSeriesTrendline(string seriesName, bool ignoreCase = true) {
            ArgumentNullException.ThrowIfNull(seriesName);

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                series.RemoveChildren(CNs, "trendline");
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            MarkChartEdited();
            return this;
        }

        private PowerPointChart ClearValueAxisDisplayUnits(Func<OfficeXmlNode, bool>? predicate) {
            OfficeXmlNode axis = FindAxis(ValueAxisName, predicate);
            if (axis.IsNone) {
                return this;
            }

            OfficeXmlNode displayUnits = axis.Child(CNs, "dispUnits");
            if (displayUnits.Exists) displayUnits.Remove();
            MarkChartEdited();
            return this;
        }

        private PowerPointChart ApplyToAllDataLabels(Action<OfficeXmlNode> apply) {
            OfficeXmlNode plotArea = PlotAreaNode;
            if (plotArea.IsNone) {
                return this;
            }

            foreach (OfficeXmlNode chartGroup in DataLabelGroups(plotArea)) {
                apply(EnsureDataLabels(chartGroup));
            }

            MarkChartEdited();
            return this;
        }

        /// <summary>
        ///     Sets the category axis orientation (normal or reversed order).
        /// </summary>
        public PowerPointChart SetCategoryAxisReverseOrder(bool reverseOrder = true) {
            OfficeXmlNode axis = FindAxis(CategoryAxisName, null);
            if (axis.IsNone) {
                return this;
            }

            OfficeXmlNode scaling = EnsureScaling(axis);
            ReplaceValChild(scaling, "orientation", reverseOrder ? "maxMin" : "minMax");
            MarkChartEdited();
            return this;
        }

        /// <summary>
        ///     Sets value axis scale parameters.
        /// </summary>
        public PowerPointChart SetValueAxisScale(double? minimum = null, double? maximum = null,
            double? majorUnit = null, double? minorUnit = null, double? logBase = null,
            bool? reverseOrder = null, bool? logScale = null) {
            ValidateAxisScale(minimum, maximum, majorUnit, minorUnit, logScale, logBase);

            OfficeXmlNode axis = FindAxis(ValueAxisName, null);
            if (axis.IsNone) {
                return this;
            }

            ApplyAxisScale(axis, minimum, maximum, majorUnit, minorUnit, reverseOrder, logScale, logBase);
            MarkChartEdited();
            return this;
        }

        /// <summary>
        ///     Sets where the category axis crosses the value axis.
        /// </summary>
        public PowerPointChart SetCategoryAxisCrossing(OfficeChartAxisCrossingPosition crosses, double? crossesAt = null) {
            if (crossesAt != null && (double.IsNaN(crossesAt.Value) || double.IsInfinity(crossesAt.Value))) {
                throw new ArgumentOutOfRangeException(nameof(crossesAt));
            }

            OfficeXmlNode axis = FindAxis(CategoryAxisName, null);
            if (axis.IsNone) {
                return this;
            }

            ApplyAxisCrossing(axis, crosses, crossesAt);
            MarkChartEdited();
            return this;
        }

        /// <summary>
        ///     Sets where the value axis crosses the category axis.
        /// </summary>
        public PowerPointChart SetValueAxisCrossing(OfficeChartAxisCrossingPosition crosses, double? crossesAt = null) {
            if (crossesAt != null && (double.IsNaN(crossesAt.Value) || double.IsInfinity(crossesAt.Value))) {
                throw new ArgumentOutOfRangeException(nameof(crossesAt));
            }

            OfficeXmlNode axis = FindAxis(ValueAxisName, null);
            if (axis.IsNone) {
                return this;
            }

            ValidateCrossesAtForAxis(axis, crossesAt);
            ApplyAxisCrossing(axis, crosses, crossesAt);
            MarkChartEdited();
            return this;
        }

        /// <summary>
        ///     Sets scatter chart X-axis scale (bottom value axis).
        /// </summary>
        public PowerPointChart SetScatterXAxisScale(double? minimum = null, double? maximum = null,
            double? majorUnit = null, double? minorUnit = null, bool? reverseOrder = null,
            bool? logScale = null, double? logBase = null) {
            ValidateAxisScale(minimum, maximum, majorUnit, minorUnit, logScale, logBase);

            OfficeXmlNode plotArea = PlotAreaNode;
            if (plotArea.IsNone) {
                return this;
            }

            OfficeXmlNode axis = ResolveScatterXAxis(plotArea);
            if (axis.IsNone) {
                return this;
            }

            ApplyAxisScale(axis, minimum, maximum, majorUnit, minorUnit, reverseOrder, logScale, logBase);
            MarkChartEdited();
            return this;
        }

        /// <summary>
        ///     Sets scatter chart Y-axis scale (left value axis).
        /// </summary>
        public PowerPointChart SetScatterYAxisScale(double? minimum = null, double? maximum = null,
            double? majorUnit = null, double? minorUnit = null, bool? reverseOrder = null,
            bool? logScale = null, double? logBase = null) {
            ValidateAxisScale(minimum, maximum, majorUnit, minorUnit, logScale, logBase);

            OfficeXmlNode plotArea = PlotAreaNode;
            if (plotArea.IsNone) {
                return this;
            }

            OfficeXmlNode axis = ResolveScatterYAxis(plotArea);
            if (axis.IsNone) {
                return this;
            }

            ApplyAxisScale(axis, minimum, maximum, majorUnit, minorUnit, reverseOrder, logScale, logBase);
            MarkChartEdited();
            return this;
        }

        /// <summary>
        ///     Sets where the scatter X-axis crosses the Y-axis.
        /// </summary>
        public PowerPointChart SetScatterXAxisCrossing(OfficeChartAxisCrossingPosition? crosses = null, double? crossesAt = null) {
            if (crossesAt != null && (double.IsNaN(crossesAt.Value) || double.IsInfinity(crossesAt.Value))) {
                throw new ArgumentOutOfRangeException(nameof(crossesAt));
            }

            OfficeXmlNode plotArea = PlotAreaNode;
            if (plotArea.IsNone) {
                return this;
            }

            OfficeXmlNode axis = ResolveScatterXAxis(plotArea);
            if (axis.IsNone) {
                return this;
            }

            ValidateCrossesAtForAxis(axis, crossesAt);
            ApplyAxisCrossing(axis, crosses ?? OfficeChartAxisCrossingPosition.AutoZero, crossesAt);
            MarkChartEdited();
            return this;
        }

        /// <summary>
        ///     Sets where the scatter Y-axis crosses the X-axis.
        /// </summary>
        public PowerPointChart SetScatterYAxisCrossing(OfficeChartAxisCrossingPosition? crosses = null, double? crossesAt = null) {
            if (crossesAt != null && (double.IsNaN(crossesAt.Value) || double.IsInfinity(crossesAt.Value))) {
                throw new ArgumentOutOfRangeException(nameof(crossesAt));
            }

            OfficeXmlNode plotArea = PlotAreaNode;
            if (plotArea.IsNone) {
                return this;
            }

            OfficeXmlNode axis = ResolveScatterYAxis(plotArea);
            if (axis.IsNone) {
                return this;
            }

            ValidateCrossesAtForAxis(axis, crossesAt);
            ApplyAxisCrossing(axis, crosses ?? OfficeChartAxisCrossingPosition.AutoZero, crossesAt);
            MarkChartEdited();
            return this;
        }

        // A pie or doughnut series fill colours every slice in PowerPoint, while other renderers
        // keeps varying the slices when c:varyColors stays on. Turning it off, as PowerPoint
        // does when a series fill is chosen, makes every renderer draw the same chart.
        private static void UseSeriesColorForEverySlice(OfficeXmlNode series) {
            if (SeriesKind(series) != ChartSeriesKind.Pie) return;
            OfficeXmlNode varyColors = series.Parent.Child(CNs, "varyColors");
            if (varyColors.Exists) varyColors.SetBooleanAttribute("val", false);
        }

        /// <summary>
        ///     Sets the fill color for a chart series by index.
        /// </summary>
        public PowerPointChart SetSeriesFillColor(int seriesIndex, OfficeColor color) {
            ArgumentOutOfRangeException.ThrowIfNegative(seriesIndex);

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                OfficeXmlNode props = EnsureChartShapeProperties(series);
                ApplySolidFill(props, color);
                OfficeOpenXmlChartAutomaticColors.ClearAutomaticSeriesColor(series);
                UseSeriesColorForEverySlice(series);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            MarkChartEdited();
            return this;
        }

        /// <summary>
        ///     Sets the fill color for a chart series by name.
        /// </summary>
        public PowerPointChart SetSeriesFillColor(string seriesName, OfficeColor color, bool ignoreCase = true) {
            ArgumentNullException.ThrowIfNull(seriesName);

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                OfficeXmlNode props = EnsureChartShapeProperties(series);
                ApplySolidFill(props, color);
                OfficeOpenXmlChartAutomaticColors.ClearAutomaticSeriesColor(series);
                UseSeriesColorForEverySlice(series);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            MarkChartEdited();
            return this;
        }

        /// <summary>
        ///     Sets the line color for a chart series by index.
        /// </summary>
#pragma warning disable RS0026 // Existing overload family; only the color type or font-family parameter names changed.
        public PowerPointChart SetSeriesLineColor(int seriesIndex, OfficeColor color, OfficeLength? width = null) {
#pragma warning restore RS0026
            ArgumentOutOfRangeException.ThrowIfNegative(seriesIndex);

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                OfficeXmlNode props = EnsureChartShapeProperties(series);
                ApplyLine(props, color, width);
                OfficeOpenXmlChartAutomaticColors.ClearAutomaticSeriesColor(series);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            MarkChartEdited();
            return this;
        }

        /// <summary>
        ///     Sets the line color for a chart series by name.
        /// </summary>
#pragma warning disable RS0026 // Existing overload family; only the color type or font-family parameter names changed.
        public PowerPointChart SetSeriesLineColor(string seriesName, OfficeColor color, OfficeLength? width = null, bool ignoreCase = true) {
#pragma warning restore RS0026
            ArgumentNullException.ThrowIfNull(seriesName);

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                OfficeXmlNode props = EnsureChartShapeProperties(series);
                ApplyLine(props, color, width);
                OfficeOpenXmlChartAutomaticColors.ClearAutomaticSeriesColor(series);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            MarkChartEdited();
            return this;
        }

        /// <summary>
        ///     Sets the marker style for a chart series by index.
        /// </summary>
#pragma warning disable RS0026 // Existing overload family; only the color type or font-family parameter names changed.
        public PowerPointChart SetSeriesMarker(int seriesIndex, OfficeChartMarkerShape style, OfficeLength? size = null, OfficeColor? fillColor = null, OfficeColor? lineColor = null, OfficeLength? lineWidth = null) {
#pragma warning restore RS0026
            ArgumentOutOfRangeException.ThrowIfNegative(seriesIndex);
            if (size.HasValue && (size.Value < OfficeLength.Points(1) || size.Value > OfficeLength.Points(72))) {
                throw new ArgumentOutOfRangeException(nameof(size), "Marker size must be between 1 and 72.");
            }

            bool applied = ApplySeriesMarkerByIndex(seriesIndex, marker => {
                ApplyMarker(marker, style, size, fillColor, lineColor, lineWidth);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            MarkChartEdited();
            return this;
        }

        /// <summary>
        ///     Sets the marker style for a chart series by name.
        /// </summary>
#pragma warning disable RS0026 // Existing overload family; only the color type or font-family parameter names changed.
        public PowerPointChart SetSeriesMarker(string seriesName, OfficeChartMarkerShape style, OfficeLength? size = null, OfficeColor? fillColor = null, OfficeColor? lineColor = null, OfficeLength? lineWidth = null, bool ignoreCase = true) {
#pragma warning restore RS0026
            ArgumentNullException.ThrowIfNull(seriesName);
            if (size.HasValue && (size.Value < OfficeLength.Points(1) || size.Value > OfficeLength.Points(72))) {
                throw new ArgumentOutOfRangeException(nameof(size), "Marker size must be between 1 and 72.");
            }

            bool applied = ApplySeriesMarkerByName(seriesName, ignoreCase, marker => {
                ApplyMarker(marker, style, size, fillColor, lineColor, lineWidth);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            MarkChartEdited();
            return this;
        }

    }
}
