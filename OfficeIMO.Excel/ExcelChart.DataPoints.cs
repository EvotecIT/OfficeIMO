using System;
using System.Linq;
using DocumentFormat.OpenXml;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Represents a chart on a worksheet.
    /// </summary>
    public sealed partial class ExcelChart {
        /// <summary>
        /// Sets the fill color for a single chart data point by series index and zero-based point index.
        /// </summary>
        public ExcelChart SetDataPointFillColor(int seriesIndex, uint pointIndex, string color) {
            if (string.IsNullOrWhiteSpace(color)) {
                throw new ArgumentException("Data point fill color cannot be null or empty.", nameof(color));
            }

            return SetDataPointColor(seriesIndex, pointIndex, fillColor: color);
        }

        /// <summary>
        /// Sets the fill color for a single chart data point by series name and zero-based point index.
        /// </summary>
        public ExcelChart SetDataPointFillColor(string seriesName, uint pointIndex, string color, bool ignoreCase = true) {
            if (string.IsNullOrWhiteSpace(color)) {
                throw new ArgumentException("Data point fill color cannot be null or empty.", nameof(color));
            }

            return SetDataPointColor(seriesName, pointIndex, fillColor: color, ignoreCase: ignoreCase);
        }

        /// <summary>
        /// Sets the outline color and optional outline width for a single chart data point by series index and zero-based point index.
        /// </summary>
        public ExcelChart SetDataPointLineColor(int seriesIndex, uint pointIndex, string color, double? widthPoints = null) {
            if (string.IsNullOrWhiteSpace(color)) {
                throw new ArgumentException("Data point line color cannot be null or empty.", nameof(color));
            }

            return SetDataPointColor(seriesIndex, pointIndex, lineColor: color, lineWidthPoints: widthPoints);
        }

        /// <summary>
        /// Sets the outline color and optional outline width for a single chart data point by series name and zero-based point index.
        /// </summary>
        public ExcelChart SetDataPointLineColor(string seriesName, uint pointIndex, string color, double? widthPoints = null, bool ignoreCase = true) {
            if (string.IsNullOrWhiteSpace(color)) {
                throw new ArgumentException("Data point line color cannot be null or empty.", nameof(color));
            }

            return SetDataPointColor(seriesName, pointIndex, lineColor: color, lineWidthPoints: widthPoints, ignoreCase: ignoreCase);
        }

        /// <summary>
        /// Sets fill and/or outline styling for a single chart data point by series index and zero-based point index.
        /// </summary>
        public ExcelChart SetDataPointColor(int seriesIndex, uint pointIndex, string? fillColor = null, string? lineColor = null, double? lineWidthPoints = null) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }

            ValidateDataPointStyle(fillColor, lineColor, lineWidthPoints);

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                C.DataPoint point = EnsureDataPoint(series, pointIndex);
                ApplyDataPointStyle(point, fillColor, lineColor, lineWidthPoints);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets fill and/or outline styling for a single chart data point by series name and zero-based point index.
        /// </summary>
        public ExcelChart SetDataPointColor(string seriesName, uint pointIndex, string? fillColor = null, string? lineColor = null, double? lineWidthPoints = null, bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }

            ValidateDataPointStyle(fillColor, lineColor, lineWidthPoints);

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                C.DataPoint point = EnsureDataPoint(series, pointIndex);
                ApplyDataPointStyle(point, fillColor, lineColor, lineWidthPoints);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        private static void ValidateDataPointStyle(string? fillColor, string? lineColor, double? lineWidthPoints) {
            if (fillColor != null && string.IsNullOrWhiteSpace(fillColor)) {
                throw new ArgumentException("Data point fill color cannot be empty.", nameof(fillColor));
            }
            if (lineColor != null && string.IsNullOrWhiteSpace(lineColor)) {
                throw new ArgumentException("Data point line color cannot be empty.", nameof(lineColor));
            }
            if (fillColor == null && lineColor == null && lineWidthPoints == null) {
                throw new ArgumentException("Specify a fill color, line color, or line width to style the data point.");
            }
        }

        private static C.DataPoint EnsureDataPoint(OpenXmlCompositeElement series, uint pointIndex) {
            C.DataPoint? point = series.Elements<C.DataPoint>()
                .FirstOrDefault(dataPoint => dataPoint.GetFirstChild<C.Index>()?.Val?.Value == pointIndex);

            if (point == null) {
                point = new C.DataPoint(new C.Index { Val = pointIndex });
                InsertDataPoint(series, point);
            }

            return point;
        }

        private static void ApplyDataPointStyle(C.DataPoint point, string? fillColor, string? lineColor, double? lineWidthPoints) {
            C.ChartShapeProperties props = point.GetFirstChild<C.ChartShapeProperties>() ?? new C.ChartShapeProperties();
            if (fillColor != null) {
                ApplySolidFill(props, NormalizeHexColor(fillColor));
            }
            if (lineColor != null) {
                ApplyLine(props, NormalizeHexColor(lineColor), lineWidthPoints);
            } else if (lineWidthPoints != null) {
                ApplyOptionalLine(props, null, lineWidthPoints);
            }

            if (props.Parent == null) {
                OpenXmlElement? insertBefore = point.GetFirstChild<C.PictureOptions>();
                insertBefore ??= point.GetFirstChild<C.ExtensionList>();
                if (insertBefore != null) {
                    point.InsertBefore(props, insertBefore);
                } else {
                    point.Append(props);
                }
            }
        }

        private static void InsertDataPoint(OpenXmlCompositeElement series, C.DataPoint point) {
            OpenXmlElement? insertBefore = series.GetFirstChild<C.DataLabels>();
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
                series.InsertBefore(point, insertBefore);
            } else {
                series.Append(point);
            }
        }
    }
}
