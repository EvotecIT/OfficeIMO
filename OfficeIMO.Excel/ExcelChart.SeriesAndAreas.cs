using System;
using System.Collections.Generic;
using System.Globalization;
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
    public sealed partial class ExcelChart {
        /// <summary>
        /// Removes the fill from a chart series and optionally removes its outline.
        /// </summary>
        public ExcelChart SetSeriesNoFill(int seriesIndex, bool noLine = true) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                C.ChartShapeProperties props = EnsureChartShapeProperties(series);
                ApplyNoFill(props);
                if (noLine) {
                    ApplyNoLine(props);
                }
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
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
        /// Sets the fill color for one point in a chart series by index.
        /// </summary>
        public ExcelChart SetSeriesPointFillColor(int seriesIndex, int pointIndex, string color) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            if (pointIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(pointIndex));
            }
            if (string.IsNullOrWhiteSpace(color)) {
                throw new ArgumentException("Point color cannot be null or empty.", nameof(color));
            }

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                ApplyPointFill(series, pointIndex, NormalizeHexColor(color));
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets the fill color for one point in a chart series by name.
        /// </summary>
        public ExcelChart SetSeriesPointFillColor(string seriesName, int pointIndex, string color, bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            if (pointIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(pointIndex));
            }
            if (string.IsNullOrWhiteSpace(color)) {
                throw new ArgumentException("Point color cannot be null or empty.", nameof(color));
            }

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                ApplyPointFill(series, pointIndex, NormalizeHexColor(color));
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

        private static void ApplyPointFill(OpenXmlCompositeElement series, int pointIndex, string color) {
            C.DataPoint point = EnsureDataPoint(series, pointIndex);
            C.ChartShapeProperties props = point.GetFirstChild<C.ChartShapeProperties>() ?? new C.ChartShapeProperties();
            ApplySolidFill(props, color);
            if (props.Parent == null) {
                point.Append(props);
            }
        }

        private static C.DataPoint EnsureDataPoint(OpenXmlCompositeElement series, int pointIndex) {
            foreach (C.DataPoint existing in series.Elements<C.DataPoint>()) {
                uint? existingIndex = existing.GetFirstChild<C.Index>()?.Val?.Value;
                if (existingIndex.HasValue && existingIndex.Value == (uint)pointIndex) {
                    return existing;
                }
            }

            var point = new C.DataPoint(new C.Index { Val = (uint)pointIndex });
            OpenXmlElement? insertBefore = series.Elements<C.DataPoint>()
                .FirstOrDefault(existing => {
                    uint? existingIndex = existing.GetFirstChild<C.Index>()?.Val?.Value;
                    return existingIndex.HasValue && existingIndex.Value > (uint)pointIndex;
                });
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
            EnsureSeriesChildPosition(series, point, insertBefore);
            return point;
        }

    }
}
