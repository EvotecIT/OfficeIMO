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
    // Represents a chart on a worksheet.
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
        public ExcelChart SetSeriesFillColor(int seriesIndex, OfficeColor color) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                C.ChartShapeProperties props = EnsureChartShapeProperties(series);
                ApplySolidFill(props, color);
                OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartAutomaticColors.ClearAutomaticSeriesColor(series);
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
        public ExcelChart SetSeriesFillColor(string seriesName, OfficeColor color, bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                C.ChartShapeProperties props = EnsureChartShapeProperties(series);
                ApplySolidFill(props, color);
                OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartAutomaticColors.ClearAutomaticSeriesColor(series);
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
        public ExcelChart SetSeriesLineColor(int seriesIndex, OfficeColor color, OfficeLength? width = null) {
            OfficeChartStyleBounds.ValidateLineWidth(width?.ToPoints(), nameof(width), allowZero: true);
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                C.ChartShapeProperties props = EnsureChartShapeProperties(series);
                ApplyLine(props, color, width);
                OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartAutomaticColors.ClearAutomaticSeriesColor(series);
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
        public ExcelChart SetSeriesLineColor(string seriesName, OfficeColor color, OfficeLength? width = null, bool ignoreCase = true) {
            OfficeChartStyleBounds.ValidateLineWidth(width?.ToPoints(), nameof(width), allowZero: true);
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                C.ChartShapeProperties props = EnsureChartShapeProperties(series);
                ApplyLine(props, color, width);
                OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartAutomaticColors.ClearAutomaticSeriesColor(series);
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
        public ExcelChart SetSeriesPointFillColor(int seriesIndex, int pointIndex, OfficeColor color) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            if (pointIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(pointIndex));
            }

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                ApplyPointFill(series, pointIndex, color);
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
        public ExcelChart SetSeriesPointFillColor(string seriesName, int pointIndex, OfficeColor color, bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            if (pointIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(pointIndex));
            }

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                ApplyPointFill(series, pointIndex, color);
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
        public ExcelChart SetSeriesMarker(int seriesIndex, OfficeChartMarkerShape style, OfficeLength? size = null, OfficeColor? fillColor = null, OfficeColor? lineColor = null, OfficeLength? lineWidth = null) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            if (size.HasValue && (size.Value < OfficeLength.Points(2) || size.Value > OfficeLength.Points(72))) {
                throw new ArgumentOutOfRangeException(nameof(size), "Marker size must be between 2 and 72.");
            }

            OfficeChartStyleBounds.ValidateLineWidth(lineWidth?.ToPoints(), nameof(lineWidth), allowZero: true);
            bool applied = ApplySeriesMarkerByIndex(seriesIndex, marker => {
                ApplyMarker(marker, style, size, fillColor, lineColor, lineWidth);
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
        public ExcelChart SetSeriesMarker(string seriesName, OfficeChartMarkerShape style, OfficeLength? size = null, OfficeColor? fillColor = null, OfficeColor? lineColor = null, OfficeLength? lineWidth = null, bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            if (size.HasValue && (size.Value < OfficeLength.Points(2) || size.Value > OfficeLength.Points(72))) {
                throw new ArgumentOutOfRangeException(nameof(size), "Marker size must be between 2 and 72.");
            }

            OfficeChartStyleBounds.ValidateLineWidth(lineWidth?.ToPoints(), nameof(lineWidth), allowZero: true);
            bool applied = ApplySeriesMarkerByName(seriesName, ignoreCase, marker => {
                ApplyMarker(marker, style, size, fillColor, lineColor, lineWidth);
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
        public ExcelChart SetChartAreaStyle(OfficeColor? fillColor = null, OfficeColor? lineColor = null,
            OfficeLength? lineWidth = null, bool noFill = false, bool noLine = false) {
            ValidateDataLabelShapeStyle(fillColor, lineColor, lineWidth, noFill, noLine);

            ChartPart chartPart = GetChartPart();
            C.ChartSpace? chartSpace = chartPart.ChartSpace;
            if (chartSpace == null) {
                return this;
            }

            C.ShapeProperties props = chartSpace.GetFirstChild<C.ShapeProperties>() ?? new C.ShapeProperties();
            ApplyAreaStyle(props, fillColor, lineColor, lineWidth, noFill, noLine);
            if (props.Parent == null) {
                chartSpace.Append(props);
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets plot area fill/line styling.
        /// </summary>
        public ExcelChart SetPlotAreaStyle(OfficeColor? fillColor = null, OfficeColor? lineColor = null,
            OfficeLength? lineWidth = null, bool noFill = false, bool noLine = false) {
            ValidateDataLabelShapeStyle(fillColor, lineColor, lineWidth, noFill, noLine);

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            C.ShapeProperties props = plotArea.GetFirstChild<C.ShapeProperties>() ?? new C.ShapeProperties();
            ApplyAreaStyle(props, fillColor, lineColor, lineWidth, noFill, noLine);
            if (props.Parent == null) {
                plotArea.Append(props);
            }

            Save();
            return this;
        }

        /// <summary>
        /// Adds or replaces a trendline for a chart series by index.
        /// </summary>
        public ExcelChart SetSeriesTrendline(int seriesIndex, OfficeChartTrendlineType type,
            int? order = null, int? period = null, double? forward = null, double? backward = null, double? intercept = null,
            bool displayEquation = false, bool displayRSquared = false, OfficeColor? lineColor = null, OfficeLength? lineWidth = null) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            ValidateTrendline(type, order, period, forward, backward, lineColor, lineWidth);

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                ApplyTrendline(series, type, order, period, forward, backward, intercept, displayEquation, displayRSquared, lineColor, lineWidth);
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
        public ExcelChart SetSeriesTrendline(string seriesName, OfficeChartTrendlineType type,
            int? order = null, int? period = null, double? forward = null, double? backward = null, double? intercept = null,
            bool displayEquation = false, bool displayRSquared = false, OfficeColor? lineColor = null, OfficeLength? lineWidth = null,
            bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            ValidateTrendline(type, order, period, forward, backward, lineColor, lineWidth);

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                ApplyTrendline(series, type, order, period, forward, backward, intercept, displayEquation, displayRSquared, lineColor, lineWidth);
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

        private static void ApplyPointFill(OpenXmlCompositeElement series, int pointIndex, OfficeColor color) {
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
