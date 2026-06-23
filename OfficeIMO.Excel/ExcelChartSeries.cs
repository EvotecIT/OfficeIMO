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
        public ExcelChartSeries(string name, IEnumerable<double> values, ExcelChartType? chartType = null, ExcelChartAxisGroup axisGroup = ExcelChartAxisGroup.Primary, string? seriesColorArgb = null)
            : this(name, (values ?? Array.Empty<double>()).ToList(), chartType, axisGroup, seriesColorArgb, seriesLineWidth: null, seriesLineDashStyle: null, pointColorArgb: null, showMarkers: true, markerSize: null, markerShape: null, markerOutlineColorArgb: null, markerOutlineWidth: null, ownsValues: true) {
        }

        private ExcelChartSeries(string name, IReadOnlyList<double> values, ExcelChartType? chartType, ExcelChartAxisGroup axisGroup, string? seriesColorArgb, double? seriesLineWidth, OfficeStrokeDashStyle? seriesLineDashStyle, IReadOnlyList<string?>? pointColorArgb, bool showMarkers, int? markerSize, OfficeChartMarkerShape? markerShape, string? markerOutlineColorArgb, double? markerOutlineWidth, bool ownsValues) {
            Name = name ?? string.Empty;
            Values = values ?? Array.Empty<double>();
            ChartType = chartType;
            AxisGroup = axisGroup;
            SeriesColorArgb = NormalizeColor(seriesColorArgb);
            SeriesLineWidth = seriesLineWidth;
            SeriesLineDashStyle = seriesLineDashStyle;
            PointColorArgb = NormalizeColors(pointColorArgb);
            ShowMarkers = showMarkers;
            MarkerSize = markerSize;
            MarkerShape = markerShape;
            MarkerOutlineColorArgb = NormalizeColor(markerOutlineColorArgb);
            MarkerOutlineWidth = markerOutlineWidth;
        }

        internal static ExcelChartSeries CreateOwned(string name, IReadOnlyList<double> values, ExcelChartType? chartType = null, ExcelChartAxisGroup axisGroup = ExcelChartAxisGroup.Primary)
            => new(name, values, chartType, axisGroup, seriesColorArgb: null, seriesLineWidth: null, seriesLineDashStyle: null, pointColorArgb: null, showMarkers: true, markerSize: null, markerShape: null, markerOutlineColorArgb: null, markerOutlineWidth: null, ownsValues: true);

        internal ExcelChartSeries WithImageExportStyle(string? seriesColorArgb, double? seriesLineWidth, OfficeStrokeDashStyle? seriesLineDashStyle, IReadOnlyList<string?>? pointColorArgb, bool showMarkers, int? markerSize, OfficeChartMarkerShape? markerShape, string? markerOutlineColorArgb, double? markerOutlineWidth) =>
            new(Name, Values, ChartType, AxisGroup, seriesColorArgb, seriesLineWidth, seriesLineDashStyle, pointColorArgb, showMarkers, markerSize, markerShape, markerOutlineColorArgb, markerOutlineWidth, ownsValues: false);

        /// <summary>
        /// Gets the series name.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Gets the series values.
        /// </summary>
        public IReadOnlyList<double> Values { get; }

        /// <summary>
        /// Gets the optional chart type override for this series.
        /// </summary>
        public ExcelChartType? ChartType { get; }

        /// <summary>
        /// Gets the axis group for this series.
        /// </summary>
        public ExcelChartAxisGroup AxisGroup { get; }

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
