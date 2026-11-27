using OfficeIMO.OpenXml.Internal;
using static OfficeIMO.Excel.OpenXml.ExcelChartXml;

namespace OfficeIMO.Excel {
    // Represents a chart on a worksheet.
    public sealed partial class ExcelChart {
        /// <summary>
        /// Sets the fill color for a single chart data point by series index and zero-based point index.
        /// </summary>
        public ExcelChart SetDataPointFillColor(int seriesIndex, uint pointIndex, OfficeColor color) {

            return SetDataPointColor(seriesIndex, pointIndex, fillColor: color);
        }

        /// <summary>
        /// Sets the fill color for a single chart data point by series name and zero-based point index.
        /// </summary>
        public ExcelChart SetDataPointFillColor(string seriesName, uint pointIndex, OfficeColor color, bool ignoreCase = true) {

            return SetDataPointColor(seriesName, pointIndex, fillColor: color, ignoreCase: ignoreCase);
        }

        /// <summary>
        /// Sets the outline color and optional outline width for a single chart data point by series index and zero-based point index.
        /// </summary>
        public ExcelChart SetDataPointLineColor(int seriesIndex, uint pointIndex, OfficeColor color, OfficeLength? width = null) {

            return SetDataPointColor(seriesIndex, pointIndex, lineColor: color, lineWidth: width);
        }

        /// <summary>
        /// Sets the outline color and optional outline width for a single chart data point by series name and zero-based point index.
        /// </summary>
        public ExcelChart SetDataPointLineColor(string seriesName, uint pointIndex, OfficeColor color, OfficeLength? width = null, bool ignoreCase = true) {

            return SetDataPointColor(seriesName, pointIndex, lineColor: color, lineWidth: width, ignoreCase: ignoreCase);
        }

        /// <summary>
        /// Sets fill and/or outline styling for a single chart data point by series index and zero-based point index.
        /// </summary>
        public ExcelChart SetDataPointColor(int seriesIndex, uint pointIndex, OfficeColor? fillColor = null, OfficeColor? lineColor = null, OfficeLength? lineWidth = null) {
            ArgumentOutOfRangeException.ThrowIfNegative(seriesIndex);

            ValidateDataPointStyle(fillColor, lineColor, lineWidth);

            bool applied = ApplySeriesByIndex(seriesIndex, (OfficeXmlNode series) => {
                OfficeXmlNode point = EnsureDataPoint(series, pointIndex);
                ApplyDataPointStyle(point, fillColor, lineColor, lineWidth);
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
        public ExcelChart SetDataPointColor(string seriesName, uint pointIndex, OfficeColor? fillColor = null, OfficeColor? lineColor = null, OfficeLength? lineWidth = null, bool ignoreCase = true) {
            ArgumentNullException.ThrowIfNull(seriesName);

            ValidateDataPointStyle(fillColor, lineColor, lineWidth);

            bool applied = ApplySeriesByName(seriesName, ignoreCase, (OfficeXmlNode series) => {
                OfficeXmlNode point = EnsureDataPoint(series, pointIndex);
                ApplyDataPointStyle(point, fillColor, lineColor, lineWidth);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        private static void ValidateDataPointStyle(OfficeColor? fillColor, OfficeColor? lineColor, OfficeLength? lineWidth) {
            OfficeChartStyleBounds.ValidateLineWidth(lineWidth?.ToPoints(), nameof(lineWidth), allowZero: true);
            if (fillColor == null && lineColor == null && lineWidth == null) {
                throw new ArgumentException("Specify a fill color, line color, or line width to style the data point.");
            }
        }

        private static OfficeXmlNode EnsureDataPoint(OfficeXmlNode series, uint pointIndex) {
            foreach (OfficeXmlNode dataPoint in series.Children(C, "dPt")) {
                if (dataPoint.Child(C, "idx").UInt32Val() == pointIndex) {
                    return dataPoint;
                }
            }

            OfficeXmlNode point = NewDataPoint(series);
            point.Append(point.NewVal(C, "idx", pointIndex));
            return InsertDataPoint(series, point);
        }

        private static void ApplyDataPointStyle(OfficeXmlNode point, OfficeColor? fillColor, OfficeColor? lineColor, OfficeLength? lineWidth) {
            OfficeXmlNode props = point.Child(C, "spPr");
            if (props.IsNone) {
                OfficeXmlNode insertBefore = point.Child(C, "pictureOptions");
                if (insertBefore.IsNone) insertBefore = point.Child(C, "extLst");
                props = point.InsertBefore(point.New(C, "spPr"), insertBefore);
            }
            if (fillColor != null) {
                ApplySolidFill(props, fillColor);
            }
            if (lineColor != null) {
                ApplyLine(props, lineColor, lineWidth);
            } else if (lineWidth != null) {
                ApplyOptionalLine(props, null, lineWidth);
            }
        }

        private static OfficeXmlNode InsertDataPoint(OfficeXmlNode series, OfficeXmlNode point) {
            OfficeXmlNode insertBefore = FirstExistingChild(series, SeriesChildrenAfterDataPoint);
            return series.InsertBefore(point, insertBefore);
        }
    }
}
