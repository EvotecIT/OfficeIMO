using OfficeIMO.OpenXml.Internal;
using static OfficeIMO.Excel.OpenXml.ExcelChartXml;

namespace OfficeIMO.Excel {
    // Represents a chart on a worksheet.
    public sealed partial class ExcelChart {
        /// <summary>
        /// Removes the fill from a chart series and optionally removes its outline.
        /// </summary>
        public ExcelChart SetSeriesNoFill(int seriesIndex, bool noLine = true) {
            ArgumentOutOfRangeException.ThrowIfNegative(seriesIndex);

            bool applied = ApplySeriesByIndex(seriesIndex, (OfficeXmlNode series) => {
                OfficeXmlNode props = EnsureChartShapeProperties(series);
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
            ArgumentOutOfRangeException.ThrowIfNegative(seriesIndex);

            bool applied = ApplySeriesByIndex(seriesIndex, (OfficeXmlNode series) => {
                OfficeXmlNode props = EnsureChartShapeProperties(series);
                ApplySolidFill(props, color);
                ClearAutomaticSeriesColor(series);
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
            ArgumentNullException.ThrowIfNull(seriesName);

            bool applied = ApplySeriesByName(seriesName, ignoreCase, (OfficeXmlNode series) => {
                OfficeXmlNode props = EnsureChartShapeProperties(series);
                ApplySolidFill(props, color);
                ClearAutomaticSeriesColor(series);
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
            ArgumentOutOfRangeException.ThrowIfNegative(seriesIndex);

            bool applied = ApplySeriesByIndex(seriesIndex, (OfficeXmlNode series) => {
                OfficeXmlNode props = EnsureChartShapeProperties(series);
                ApplyLine(props, color, width);
                ClearAutomaticSeriesColor(series);
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
            ArgumentNullException.ThrowIfNull(seriesName);

            bool applied = ApplySeriesByName(seriesName, ignoreCase, (OfficeXmlNode series) => {
                OfficeXmlNode props = EnsureChartShapeProperties(series);
                ApplyLine(props, color, width);
                ClearAutomaticSeriesColor(series);
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
            ArgumentOutOfRangeException.ThrowIfNegative(seriesIndex);
            ArgumentOutOfRangeException.ThrowIfNegative(pointIndex);

            bool applied = ApplySeriesByIndex(seriesIndex, (OfficeXmlNode series) => {
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
            ArgumentNullException.ThrowIfNull(seriesName);
            ArgumentOutOfRangeException.ThrowIfNegative(pointIndex);

            bool applied = ApplySeriesByName(seriesName, ignoreCase, (OfficeXmlNode series) => {
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
        public ExcelChart SetSeriesMarker(int seriesIndex, OfficeChartMarkerShape? style = null, OfficeLength? size = null, OfficeColor? fillColor = null, OfficeColor? lineColor = null, OfficeLength? lineWidth = null) {
            ArgumentOutOfRangeException.ThrowIfNegative(seriesIndex);
            if (size.HasValue && (size.Value < OfficeLength.Points(2) || size.Value > OfficeLength.Points(72))) {
                throw new ArgumentOutOfRangeException(nameof(size), "Marker size must be between 2 and 72.");
            }

            OfficeChartStyleBounds.ValidateLineWidth(lineWidth?.ToPoints(), nameof(lineWidth), allowZero: true);
            bool applied = ApplySeriesMarkerByIndex(seriesIndex, (OfficeXmlNode marker) => {
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
        public ExcelChart SetSeriesMarkerByName(string seriesName, OfficeChartMarkerShape? style = null, OfficeLength? size = null, OfficeColor? fillColor = null, OfficeColor? lineColor = null, OfficeLength? lineWidth = null, bool ignoreCase = true) {
            ArgumentNullException.ThrowIfNull(seriesName);
            if (size.HasValue && (size.Value < OfficeLength.Points(2) || size.Value > OfficeLength.Points(72))) {
                throw new ArgumentOutOfRangeException(nameof(size), "Marker size must be between 2 and 72.");
            }

            OfficeChartStyleBounds.ValidateLineWidth(lineWidth?.ToPoints(), nameof(lineWidth), allowZero: true);
            bool applied = ApplySeriesMarkerByName(seriesName, ignoreCase, (OfficeXmlNode marker) => {
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

            OfficeXmlNode chartSpace = ChartSpaceNode();
            if (chartSpace.IsNone) {
                return this;
            }

            OfficeXmlNode existing = chartSpace.Child(C, "spPr");
            OfficeXmlNode props = existing.Exists ? existing : chartSpace.Append(chartSpace.New(C, "spPr"));
            ApplyAreaStyle(props, fillColor, lineColor, lineWidth, noFill, noLine);

            Save();
            return this;
        }

        /// <summary>
        /// Sets plot area fill/line styling.
        /// </summary>
        public ExcelChart SetPlotAreaStyle(OfficeColor? fillColor = null, OfficeColor? lineColor = null,
            OfficeLength? lineWidth = null, bool noFill = false, bool noLine = false) {
            ValidateDataLabelShapeStyle(fillColor, lineColor, lineWidth, noFill, noLine);

            OfficeXmlNode plotArea = ChartNode().Child(C, "plotArea");
            if (plotArea.IsNone) {
                return this;
            }

            OfficeXmlNode existing = plotArea.Child(C, "spPr");
            OfficeXmlNode props = existing.Exists ? existing : plotArea.Append(plotArea.New(C, "spPr"));
            ApplyAreaStyle(props, fillColor, lineColor, lineWidth, noFill, noLine);

            Save();
            return this;
        }

        /// <summary>
        /// Adds or replaces a trendline for a chart series by index.
        /// </summary>
        public ExcelChart SetSeriesTrendline(int seriesIndex, OfficeChartTrendlineType type,
            int? order = null, int? period = null, double? forward = null, double? backward = null, double? intercept = null,
            bool displayEquation = false, bool displayRSquared = false, OfficeColor? lineColor = null, OfficeLength? lineWidth = null) {
            ArgumentOutOfRangeException.ThrowIfNegative(seriesIndex);
            ValidateTrendline(type, order, period, forward, backward, lineColor, lineWidth);

            bool applied = ApplySeriesByIndex(seriesIndex, (OfficeXmlNode series) => {
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
            ArgumentNullException.ThrowIfNull(seriesName);
            ValidateTrendline(type, order, period, forward, backward, lineColor, lineWidth);

            bool applied = ApplySeriesByName(seriesName, ignoreCase, (OfficeXmlNode series) => {
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
            ArgumentOutOfRangeException.ThrowIfNegative(seriesIndex);

            bool applied = ApplySeriesByIndex(seriesIndex, (OfficeXmlNode series) => {
                series.RemoveChildren(C, "trendline");
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
            ArgumentNullException.ThrowIfNull(seriesName);

            bool applied = ApplySeriesByName(seriesName, ignoreCase, (OfficeXmlNode series) => {
                series.RemoveChildren(C, "trendline");
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        private static void ApplyPointFill(OfficeXmlNode series, int pointIndex, OfficeColor color) {
            OfficeXmlNode point = EnsureDataPoint(series, pointIndex);
            OfficeXmlNode existing = point.Child(C, "spPr");
            OfficeXmlNode props = existing.Exists ? existing : point.Append(point.New(C, "spPr"));
            ApplySolidFill(props, color);
        }

        private static OfficeXmlNode EnsureDataPoint(OfficeXmlNode series, int pointIndex) {
            foreach (OfficeXmlNode existing in series.Children(C, "dPt")) {
                uint? existingIndex = existing.Child(C, "idx").UInt32Val();
                if (existingIndex.HasValue && existingIndex.Value == (uint)pointIndex) {
                    return existing;
                }
            }

            OfficeXmlNode point = NewDataPoint(series);
            point.Append(point.NewVal(C, "idx", (uint)pointIndex));
            OfficeXmlNode insertBefore = OfficeXmlNode.None;
            foreach (OfficeXmlNode existing in series.Children(C, "dPt")) {
                uint? existingIndex = existing.Child(C, "idx").UInt32Val();
                if (existingIndex.HasValue && existingIndex.Value > (uint)pointIndex) {
                    insertBefore = existing;
                    break;
                }
            }
            if (insertBefore.IsNone) insertBefore = FirstExistingChild(series, SeriesChildrenAfterDataPoint);
            return EnsureSeriesChildPosition(series, point, insertBefore);
        }

    }
}
