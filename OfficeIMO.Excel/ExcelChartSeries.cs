using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Represents a chart series for Excel charts.
    /// </summary>
    public sealed class ExcelChartSeries {
        /// <summary>
        /// Creates a chart series with the specified name and values.
        /// </summary>
        public ExcelChartSeries(string name, IEnumerable<double> values, OfficeChartType? chartType = null, OfficeChartAxisGroup axisGroup = OfficeChartAxisGroup.Primary, string? seriesColorArgb = null)
            : this(name, (values ?? Array.Empty<double>()).ToList(), xValues: null, chartType, axisGroup, seriesColorArgb, seriesLineWidth: null, seriesLineDashStyle: null, pointColorArgb: null, showMarkers: true, connectLine: true, markerSize: null, markerShape: null, markerOutlineColorArgb: null, markerOutlineWidth: null, ownsValues: true) {
        }

        /// <summary>
        /// Creates a chart series with explicit per-point X values for scatter-style charts.
        /// </summary>
        public ExcelChartSeries(string name, IEnumerable<double> values, IEnumerable<double> xValues, OfficeChartType? chartType = null, OfficeChartAxisGroup axisGroup = OfficeChartAxisGroup.Primary, string? seriesColorArgb = null)
            : this(name, (values ?? Array.Empty<double>()).ToList(), (xValues ?? Array.Empty<double>()).ToList(), chartType, axisGroup, seriesColorArgb, seriesLineWidth: null, seriesLineDashStyle: null, pointColorArgb: null, showMarkers: true, connectLine: true, markerSize: null, markerShape: null, markerOutlineColorArgb: null, markerOutlineWidth: null, ownsValues: true) {
        }

        private ExcelChartSeries(string name, IReadOnlyList<double> values, IReadOnlyList<double>? xValues, OfficeChartType? chartType, OfficeChartAxisGroup axisGroup, string? seriesColorArgb, double? seriesLineWidth, OfficeStrokeDashStyle? seriesLineDashStyle, IReadOnlyList<string?>? pointColorArgb, bool showMarkers, bool connectLine, int? markerSize, OfficeChartMarkerShape? markerShape, string? markerOutlineColorArgb, double? markerOutlineWidth, bool ownsValues, IReadOnlyList<double>? bubbleSizes = null) {
            OfficeChartStyleBounds.ValidateLineWidth(seriesLineWidth, nameof(seriesLineWidth), allowZero: true);
            OfficeChartStyleBounds.ValidateLineWidth(markerOutlineWidth, nameof(markerOutlineWidth), allowZero: true);
            Name = name ?? string.Empty;
            Values = values ?? Array.Empty<double>();
            XValues = xValues;
            BubbleSizes = bubbleSizes;
            ChartType = chartType;
            AxisGroup = axisGroup;
            SeriesColorArgb = NormalizeColor(seriesColorArgb);
            SeriesLineWidth = seriesLineWidth;
            SeriesLineDashStyle = seriesLineDashStyle;
            PointColorArgb = NormalizeColors(pointColorArgb);
            ShowMarkers = showMarkers;
            ConnectLine = connectLine;
            MarkerSize = markerSize;
            MarkerShape = markerShape;
            MarkerOutlineColorArgb = NormalizeColor(markerOutlineColorArgb);
            MarkerOutlineWidth = markerOutlineWidth;
        }

        internal static ExcelChartSeries CreateOwned(string name, IReadOnlyList<double> values, OfficeChartType? chartType = null, OfficeChartAxisGroup axisGroup = OfficeChartAxisGroup.Primary)
            => new(name, values, xValues: null, chartType, axisGroup, seriesColorArgb: null, seriesLineWidth: null, seriesLineDashStyle: null, pointColorArgb: null, showMarkers: true, connectLine: true, markerSize: null, markerShape: null, markerOutlineColorArgb: null, markerOutlineWidth: null, ownsValues: true);

        internal ExcelChartSeries WithXValues(IReadOnlyList<double>? xValues) =>
            new(Name, Values, xValues, ChartType, AxisGroup, SeriesColorArgb, SeriesLineWidth, SeriesLineDashStyle, PointColorArgb, ShowMarkers, ConnectLine, MarkerSize, MarkerShape, MarkerOutlineColorArgb, MarkerOutlineWidth, ownsValues: false, bubbleSizes: BubbleSizes) { SeriesStyle = SeriesStyle, PointStyles = PointStyles, PointExplosions = PointExplosions, HasAutomaticColor = HasAutomaticColor };

        internal ExcelChartSeries WithValuesAndXValues(IReadOnlyList<double> values, IReadOnlyList<double>? xValues) =>
            new(Name, values, xValues, ChartType, AxisGroup, SeriesColorArgb, SeriesLineWidth, SeriesLineDashStyle, PointColorArgb, ShowMarkers, ConnectLine, MarkerSize, MarkerShape, MarkerOutlineColorArgb, MarkerOutlineWidth, ownsValues: false, bubbleSizes: BubbleSizes) { SeriesStyle = SeriesStyle, PointStyles = PointStyles?.Count == values.Count ? PointStyles : null, PointExplosions = PointExplosions?.Count == values.Count ? PointExplosions : null, HasAutomaticColor = HasAutomaticColor };

        internal ExcelChartSeries WithImageExportStyle(string? seriesColorArgb, double? seriesLineWidth, OfficeStrokeDashStyle? seriesLineDashStyle, IReadOnlyList<string?>? pointColorArgb, bool showMarkers, bool? connectLine, int? markerSize, OfficeChartMarkerShape? markerShape, string? markerOutlineColorArgb, double? markerOutlineWidth, bool automaticColor = false) =>
            new(Name, Values, XValues, ChartType, AxisGroup, seriesColorArgb ?? SeriesColorArgb, seriesLineWidth ?? SeriesLineWidth, seriesLineDashStyle ?? SeriesLineDashStyle, pointColorArgb ?? PointColorArgb, showMarkers, connectLine ?? ConnectLine, markerSize ?? MarkerSize, markerShape ?? MarkerShape, markerOutlineColorArgb ?? MarkerOutlineColorArgb, markerOutlineWidth ?? MarkerOutlineWidth, ownsValues: false, bubbleSizes: BubbleSizes) {
                SeriesStyle = SeriesStyle,
                PointStyles = PointStyles,
                PointExplosions = PointExplosions,
                HasAutomaticColor = automaticColor && seriesColorArgb != null
            };

        /// <summary>Whether the series colour was resolved from the chart style and theme rather than stored.</summary>
        /// <remarks>A caller-supplied export palette replaces only automatic colours.</remarks>
        internal bool HasAutomaticColor { get; private set; }

        /// <summary>
        /// Gets the series name.
        /// </summary>
        public string Name { get; }

        internal OfficeChartPointStyle? SeriesStyle { get; private set; }

        internal ExcelChartSeries WithSeriesStyle(OfficeChartPointStyle? style) {
            var copy = WithPointStyles(PointStyles);
            copy.SeriesStyle = style;
            return copy;
        }

        internal IReadOnlyList<OfficeChartPointStyle?>? PointStyles { get; private set; }
        internal IReadOnlyList<int>? PointExplosions { get; private set; }

        internal ExcelChartSeries WithPointStyles(IReadOnlyList<OfficeChartPointStyle?>? styles) =>
            new ExcelChartSeries(Name, Values, XValues, ChartType, AxisGroup, SeriesColorArgb,
                SeriesLineWidth, SeriesLineDashStyle, PointColorArgb, ShowMarkers, ConnectLine,
                MarkerSize, MarkerShape, MarkerOutlineColorArgb, MarkerOutlineWidth, ownsValues: false, bubbleSizes: BubbleSizes) {
                SeriesStyle = SeriesStyle,
                PointStyles = styles,
                PointExplosions = PointExplosions,
                HasAutomaticColor = HasAutomaticColor
            };

        internal ExcelChartSeries WithPointExplosions(IReadOnlyList<int>? explosions) =>
            new ExcelChartSeries(Name, Values, XValues, ChartType, AxisGroup, SeriesColorArgb,
                SeriesLineWidth, SeriesLineDashStyle, PointColorArgb, ShowMarkers, ConnectLine,
                MarkerSize, MarkerShape, MarkerOutlineColorArgb, MarkerOutlineWidth, ownsValues: false, bubbleSizes: BubbleSizes) {
                SeriesStyle = SeriesStyle,
                PointStyles = PointStyles,
                PointExplosions = explosions,
                HasAutomaticColor = HasAutomaticColor
            };

        /// <summary>
        /// Gets the series values.
        /// </summary>
        public IReadOnlyList<double> Values { get; }

        /// <summary>
        /// Gets optional per-point X values for scatter-style chart rendering.
        /// </summary>
        public IReadOnlyList<double>? XValues { get; }

        /// <summary>Gets signed source bubble sizes aligned with X and Y values; render adapters diagnose unsupported negative-size painting.</summary>
        public IReadOnlyList<double>? BubbleSizes { get; }

        internal ExcelChartSeries WithBubbleSizes(IReadOnlyList<double> sizes) =>
            new ExcelChartSeries(Name, Values, XValues, ChartType, AxisGroup, SeriesColorArgb,
                SeriesLineWidth, SeriesLineDashStyle, PointColorArgb, ShowMarkers, ConnectLine,
                MarkerSize, MarkerShape, MarkerOutlineColorArgb, MarkerOutlineWidth, ownsValues: false, bubbleSizes: sizes) {
                SeriesStyle = SeriesStyle,
                PointStyles = PointStyles,
                PointExplosions = PointExplosions,
                HasAutomaticColor = HasAutomaticColor
            };

        /// <summary>
        /// Gets the optional chart type override for this series.
        /// </summary>
        public OfficeChartType? ChartType { get; }

        /// <summary>
        /// Gets the axis group for this series.
        /// </summary>
        public OfficeChartAxisGroup AxisGroup { get; }

        /// <summary>
        /// Gets the optional authored series color in RGB or ARGB hexadecimal form.
        /// </summary>
        public string? SeriesColorArgb { get; }

        /// <summary>
        /// Gets the optional authored series line width in drawing units.
        /// </summary>
        public double? SeriesLineWidth { get; }

        /// <summary>
        /// Gets the optional authored series line dash style.
        /// </summary>
        public OfficeStrokeDashStyle? SeriesLineDashStyle { get; }

        /// <summary>
        /// Gets optional authored point colors in RGB or ARGB hexadecimal form, aligned with <see cref="Values"/>.
        /// </summary>
        public IReadOnlyList<string?>? PointColorArgb { get; }

        /// <summary>
        /// Gets whether marker-capable chart renderers should render this series' markers.
        /// </summary>
        public bool ShowMarkers { get; }

        /// <summary>
        /// Gets whether line-capable chart renderers should render this series' connecting line segments.
        /// </summary>
        public bool ConnectLine { get; }

        /// <summary>
        /// Gets the optional authored marker diameter in drawing units.
        /// </summary>
        public int? MarkerSize { get; }

        /// <summary>
        /// Gets the optional authored marker shape.
        /// </summary>
        public OfficeChartMarkerShape? MarkerShape { get; }

        /// <summary>
        /// Gets the optional authored marker outline color in RGB or ARGB hexadecimal form.
        /// </summary>
        public string? MarkerOutlineColorArgb { get; }

        /// <summary>
        /// Gets the optional authored marker outline width in drawing units.
        /// </summary>
        public double? MarkerOutlineWidth { get; }

        private static string? NormalizeColor(string? color) {
            if (string.IsNullOrWhiteSpace(color)) {
                return null;
            }

            string value = color!.Trim();
            if (value.StartsWith("#", StringComparison.Ordinal)) {
                value = value.Substring(1);
            }

            return value.Length == 6 || value.Length == 8 ? value.ToUpperInvariant() : color;
        }

        private static IReadOnlyList<string?>? NormalizeColors(IReadOnlyList<string?>? colors) {
            if (colors == null) {
                return null;
            }

            var normalized = new string?[colors.Count];
            bool any = false;
            for (int i = 0; i < colors.Count; i++) {
                normalized[i] = NormalizeColor(colors[i]);
                any |= normalized[i] != null;
            }

            return any ? normalized : null;
        }
    }
}
