using OfficeIMO.OpenXml.Internal;
using static OfficeIMO.Excel.OpenXml.ExcelChartXml;

namespace OfficeIMO.Excel {
    // Represents a chart on a worksheet.
    public sealed partial class ExcelChart {
        /// <summary>Updates supplied data-label settings for all supported chart families; omitted flags, position and number-format settings remain unchanged.</summary>
        public ExcelChart SetDataLabels(bool? showValue = null, bool? showCategoryName = null,
            bool? showSeriesName = null, bool? showLegendKey = null, bool? showPercent = null,
            OfficeChartDataLabelPosition? position = null, string? numberFormat = null, bool? sourceLinked = null) {
            if (numberFormat != null && string.IsNullOrWhiteSpace(numberFormat)) {
                throw new ArgumentException("Number format cannot be empty.", nameof(numberFormat));
            }
            if (!ForEachDataLabelChartGroup(chartGroup => {
                ApplyDataLabels(chartGroup, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent,
                    position, numberFormat, sourceLinked);
            })) {
                return this;
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets data label text style for all chart labels.
        /// </summary>
        public ExcelChart SetDataLabelTextStyle(OfficeLength? fontSize = null, bool? bold = null, bool? italic = null,
            OfficeColor? color = null, string? fontFamily = null) {
            ValidateDataLabelTextStyle(fontSize, color, fontFamily);

            if (!ForEachDataLabelChartGroup(chartGroup => {
                ApplyDataLabelTextStyle(EnsureDataLabels(chartGroup), fontSize, bold, italic, color, fontFamily);
            })) {
                return this;
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets data label shape styling for all chart labels.
        /// </summary>
        public ExcelChart SetDataLabelShapeStyle(OfficeColor? fillColor = null, OfficeColor? lineColor = null, OfficeLength? lineWidth = null,
            bool noFill = false, bool noLine = false) {
            ValidateDataLabelShapeStyle(fillColor, lineColor, lineWidth, noFill, noLine);

            if (!ForEachDataLabelChartGroup(chartGroup => {
                ApplyDataLabelShapeStyle(EnsureDataLabels(chartGroup), fillColor, lineColor, lineWidth, noFill, noLine);
            })) {
                return this;
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets data label shape styling for a series by index.
        /// </summary>
        public ExcelChart SetSeriesDataLabelShapeStyle(int seriesIndex, OfficeColor? fillColor = null, OfficeColor? lineColor = null,
            OfficeLength? lineWidth = null, bool noFill = false, bool noLine = false) {
            ArgumentOutOfRangeException.ThrowIfNegative(seriesIndex);
            ValidateDataLabelShapeStyle(fillColor, lineColor, lineWidth, noFill, noLine);

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                ApplyDataLabelShapeStyle(EnsureDataLabels(series), fillColor, lineColor, lineWidth, noFill, noLine);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets data label shape styling for a series by name.
        /// </summary>
        public ExcelChart SetSeriesDataLabelShapeStyle(string seriesName, OfficeColor? fillColor = null, OfficeColor? lineColor = null,
            OfficeLength? lineWidth = null, bool noFill = false, bool noLine = false, bool ignoreCase = true) {
            ArgumentNullException.ThrowIfNull(seriesName);
            ValidateDataLabelShapeStyle(fillColor, lineColor, lineWidth, noFill, noLine);

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                ApplyDataLabelShapeStyle(EnsureDataLabels(series), fillColor, lineColor, lineWidth, noFill, noLine);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Configures data label leader lines for all chart labels.
        /// </summary>
        public ExcelChart SetDataLabelLeaderLines(bool showLeaderLines = true, OfficeColor? lineColor = null, OfficeLength? lineWidth = null) {
            ValidateDataLabelLeaderLines(lineColor, lineWidth);

            if (!ForEachDataLabelChartGroup(chartGroup => {
                ApplyDataLabelLeaderLines(EnsureDataLabels(chartGroup), showLeaderLines, lineColor, lineWidth);
            })) {
                return this;
            }

            Save();
            return this;
        }

        /// <summary>
        /// Configures data label leader lines for a series by index.
        /// </summary>
        public ExcelChart SetSeriesDataLabelLeaderLines(int seriesIndex, bool showLeaderLines = true, OfficeColor? lineColor = null, OfficeLength? lineWidth = null) {
            ArgumentOutOfRangeException.ThrowIfNegative(seriesIndex);
            ValidateDataLabelLeaderLines(lineColor, lineWidth);

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                ApplyDataLabelLeaderLines(EnsureDataLabels(series), showLeaderLines, lineColor, lineWidth);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Configures data label leader lines for a series by name.
        /// </summary>
        public ExcelChart SetSeriesDataLabelLeaderLines(string seriesName, bool showLeaderLines = true, OfficeColor? lineColor = null,
            OfficeLength? lineWidth = null, bool ignoreCase = true) {
            ArgumentNullException.ThrowIfNull(seriesName);
            ValidateDataLabelLeaderLines(lineColor, lineWidth);

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                ApplyDataLabelLeaderLines(EnsureDataLabels(series), showLeaderLines, lineColor, lineWidth);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Enables callout-style labels by positioning labels outside with leader lines.
        /// </summary>
        public ExcelChart SetDataLabelCallouts(bool enabled = true, OfficeChartDataLabelPosition? position = null,
            OfficeColor? lineColor = null, OfficeLength? lineWidth = null) {
            var resolvedPosition = enabled ? position ?? OfficeChartDataLabelPosition.OutsideEnd : position;
            SetDataLabels(showValue: enabled, showCategoryName: false, showSeriesName: false, showLegendKey: false,
                showPercent: false, position: resolvedPosition, numberFormat: null, sourceLinked: false);
            return SetDataLabelLeaderLines(enabled, lineColor, lineWidth);
        }

        /// <summary>
        /// Enables callout-style labels for a series by index.
        /// </summary>
        public ExcelChart SetSeriesDataLabelCallouts(int seriesIndex, bool enabled = true, OfficeChartDataLabelPosition? position = null,
            OfficeColor? lineColor = null, OfficeLength? lineWidth = null) {
            ArgumentOutOfRangeException.ThrowIfNegative(seriesIndex);
            var resolvedPosition = enabled ? position ?? OfficeChartDataLabelPosition.OutsideEnd : position;
            SetSeriesDataLabels(seriesIndex, showValue: enabled, showCategoryName: false, showSeriesName: false,
                showLegendKey: false, showPercent: false, position: resolvedPosition, numberFormat: null, sourceLinked: false);
            return SetSeriesDataLabelLeaderLines(seriesIndex, enabled, lineColor, lineWidth);
        }

        /// <summary>
        /// Enables callout-style labels for a series by name.
        /// </summary>
        public ExcelChart SetSeriesDataLabelCallouts(string seriesName, bool enabled = true, OfficeChartDataLabelPosition? position = null,
            OfficeColor? lineColor = null, OfficeLength? lineWidth = null, bool ignoreCase = true) {
            ArgumentNullException.ThrowIfNull(seriesName);
            var resolvedPosition = enabled ? position ?? OfficeChartDataLabelPosition.OutsideEnd : position;
            SetSeriesDataLabels(seriesName, showValue: enabled, showCategoryName: false, showSeriesName: false,
                showLegendKey: false, showPercent: false, position: resolvedPosition, numberFormat: null, sourceLinked: false,
                ignoreCase: ignoreCase);
            return SetSeriesDataLabelLeaderLines(seriesName, enabled, lineColor, lineWidth, ignoreCase);
        }

        /// <summary>
        /// Sets data label text style for a series by index.
        /// </summary>
        public ExcelChart SetSeriesDataLabelTextStyle(int seriesIndex, OfficeLength? fontSize = null, bool? bold = null,
            bool? italic = null, OfficeColor? color = null, string? fontFamily = null) {
            ArgumentOutOfRangeException.ThrowIfNegative(seriesIndex);
            ValidateDataLabelTextStyle(fontSize, color, fontFamily);

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                ApplyDataLabelTextStyle(EnsureDataLabels(series), fontSize, bold, italic, color, fontFamily);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets data label text style for a series by name.
        /// </summary>
        public ExcelChart SetSeriesDataLabelTextStyleByName(string seriesName, OfficeLength? fontSize = null, bool? bold = null,
            bool? italic = null, OfficeColor? color = null, string? fontFamily = null, bool ignoreCase = true) {
            ArgumentNullException.ThrowIfNull(seriesName);
            ValidateDataLabelTextStyle(fontSize, color, fontFamily);

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                ApplyDataLabelTextStyle(EnsureDataLabels(series), fontSize, bold, italic, color, fontFamily);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Applies a reusable data label template to a series by index.
        /// </summary>
        public ExcelChart SetSeriesDataLabelTemplate(int seriesIndex, ExcelChartDataLabelTemplate template) {
            ArgumentOutOfRangeException.ThrowIfNegative(seriesIndex);
            ArgumentNullException.ThrowIfNull(template);

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                ApplyDataLabelTemplate(series, template);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Applies a reusable data label template to a series by name.
        /// </summary>
        public ExcelChart SetSeriesDataLabelTemplate(string seriesName, ExcelChartDataLabelTemplate template,
            bool ignoreCase = true) {
            ArgumentNullException.ThrowIfNull(seriesName);
            ArgumentNullException.ThrowIfNull(template);

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                ApplyDataLabelTemplate(series, template);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Configures a single data label point by series index and point index.
        /// </summary>
        public ExcelChart SetSeriesDataLabelForPoint(int seriesIndex, int pointIndex, bool? showValue = null,
            bool? showCategoryName = null, bool? showSeriesName = null, bool? showLegendKey = null,
            bool? showPercent = null, bool? showBubbleSize = null,
            OfficeChartDataLabelPosition? position = null, string? numberFormat = null, bool sourceLinked = false) {
            ArgumentOutOfRangeException.ThrowIfNegative(seriesIndex);
            ArgumentOutOfRangeException.ThrowIfNegative(pointIndex);
            if (numberFormat != null && string.IsNullOrWhiteSpace(numberFormat)) {
                throw new ArgumentException("Number format cannot be empty.", nameof(numberFormat));
            }

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                OfficeXmlNode label = EnsureDataLabel(series, pointIndex);
                ApplyDataLabelOverrides(label, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent,
                    showBubbleSize, position, numberFormat, numberFormat != null ? sourceLinked : (bool?)null);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        internal ExcelChart SetSeriesDataLabelsForPoints(int seriesIndex, IReadOnlyList<int> pointIndices,
            bool? showValue = null, bool? showCategoryName = null, bool? showSeriesName = null,
            bool? showLegendKey = null, bool? showPercent = null, bool? showBubbleSize = null,
            OfficeChartDataLabelPosition? position = null, string? numberFormat = null, bool sourceLinked = false) {
            ArgumentOutOfRangeException.ThrowIfNegative(seriesIndex);
            ArgumentNullException.ThrowIfNull(pointIndices);
            if (numberFormat != null && string.IsNullOrWhiteSpace(numberFormat)) {
                throw new ArgumentException("Number format cannot be empty.", nameof(numberFormat));
            }

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                OfficeXmlNode labels = EnsureDataLabels(series);
                var labelsByIndex = new Dictionary<uint, OfficeXmlNode>();
                foreach (OfficeXmlNode existing in labels.Children(C, "dLbl")) {
                    uint? existingIndex = existing.Child(C, "idx").UInt32Val();
                    if (existingIndex.HasValue && !labelsByIndex.ContainsKey(existingIndex.Value)) {
                        labelsByIndex.Add(existingIndex.Value, existing);
                    }
                }

                var requested = new HashSet<uint>();
                foreach (int pointIndex in pointIndices) {
                    if (pointIndex < 0) {
                        throw new ArgumentOutOfRangeException(nameof(pointIndices), "Point indices cannot be negative.");
                    }

                    uint index = (uint)pointIndex;
                    if (!requested.Add(index)) continue;
                    if (!labelsByIndex.TryGetValue(index, out OfficeXmlNode label)) {
                        label = InsertDataLabel(labels, index);
                        labelsByIndex.Add(index, label);
                    }

                    ApplyDataLabelOverrides(label, showLegendKey, showValue, showCategoryName, showSeriesName,
                        showPercent, showBubbleSize, position, numberFormat, numberFormat != null ? sourceLinked : (bool?)null);
                }
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Configures a single data label point by series name and point index.
        /// </summary>
        public ExcelChart SetSeriesDataLabelForPoint(string seriesName, int pointIndex, bool? showValue = null,
            bool? showCategoryName = null, bool? showSeriesName = null, bool? showLegendKey = null,
            bool? showPercent = null, bool? showBubbleSize = null,
            OfficeChartDataLabelPosition? position = null, string? numberFormat = null, bool sourceLinked = false,
            bool ignoreCase = true) {
            ArgumentNullException.ThrowIfNull(seriesName);
            ArgumentOutOfRangeException.ThrowIfNegative(pointIndex);
            if (numberFormat != null && string.IsNullOrWhiteSpace(numberFormat)) {
                throw new ArgumentException("Number format cannot be empty.", nameof(numberFormat));
            }

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                OfficeXmlNode label = EnsureDataLabel(series, pointIndex);
                ApplyDataLabelOverrides(label, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent,
                    showBubbleSize, position, numberFormat, numberFormat != null ? sourceLinked : (bool?)null);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets data label text style for a specific point by series index.
        /// </summary>
        public ExcelChart SetSeriesDataLabelTextStyleForPoint(int seriesIndex, int pointIndex,
            OfficeLength? fontSize = null, bool? bold = null, bool? italic = null,
            OfficeColor? color = null, string? fontFamily = null) {
            ArgumentOutOfRangeException.ThrowIfNegative(seriesIndex);
            ArgumentOutOfRangeException.ThrowIfNegative(pointIndex);
            ValidateDataLabelTextStyle(fontSize, color, fontFamily);

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                OfficeXmlNode label = EnsureDataLabel(series, pointIndex);
                ApplyDataLabelTextStyle(label, fontSize, bold, italic, color, fontFamily);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets data label text style for a specific point by series name.
        /// </summary>
        public ExcelChart SetSeriesDataLabelTextStyleForPointByName(string seriesName, int pointIndex,
            OfficeLength? fontSize = null, bool? bold = null, bool? italic = null,
            OfficeColor? color = null, string? fontFamily = null, bool ignoreCase = true) {
            ArgumentNullException.ThrowIfNull(seriesName);
            ArgumentOutOfRangeException.ThrowIfNegative(pointIndex);
            ValidateDataLabelTextStyle(fontSize, color, fontFamily);

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                OfficeXmlNode label = EnsureDataLabel(series, pointIndex);
                ApplyDataLabelTextStyle(label, fontSize, bold, italic, color, fontFamily);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets data label shape styling for a specific point by series index.
        /// </summary>
        public ExcelChart SetSeriesDataLabelShapeStyleForPoint(int seriesIndex, int pointIndex,
            OfficeColor? fillColor = null, OfficeColor? lineColor = null,
            OfficeLength? lineWidth = null, bool noFill = false, bool noLine = false) {
            ArgumentOutOfRangeException.ThrowIfNegative(seriesIndex);
            ArgumentOutOfRangeException.ThrowIfNegative(pointIndex);
            ValidateDataLabelShapeStyle(fillColor, lineColor, lineWidth, noFill, noLine);

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                OfficeXmlNode label = EnsureDataLabel(series, pointIndex);
                ApplyDataLabelShapeStyle(label, fillColor, lineColor, lineWidth, noFill, noLine);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets data label shape styling for a specific point by series name.
        /// </summary>
        public ExcelChart SetSeriesDataLabelShapeStyleForPoint(string seriesName, int pointIndex,
            OfficeColor? fillColor = null, OfficeColor? lineColor = null,
            OfficeLength? lineWidth = null, bool noFill = false, bool noLine = false,
            bool ignoreCase = true) {
            ArgumentNullException.ThrowIfNull(seriesName);
            ArgumentOutOfRangeException.ThrowIfNegative(pointIndex);
            ValidateDataLabelShapeStyle(fillColor, lineColor, lineWidth, noFill, noLine);

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                OfficeXmlNode label = EnsureDataLabel(series, pointIndex);
                ApplyDataLabelShapeStyle(label, fillColor, lineColor, lineWidth, noFill, noLine);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets the separator for all data labels.
        /// </summary>
        public ExcelChart SetDataLabelSeparator(string? separator) {
            if (!ForEachDataLabelChartGroup(chartGroup => {
                ApplyDataLabelSeparator(EnsureDataLabels(chartGroup), separator);
            })) {
                return this;
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets the separator for data labels in a series by index.
        /// </summary>
        public ExcelChart SetSeriesDataLabelSeparator(int seriesIndex, string? separator) {
            ArgumentOutOfRangeException.ThrowIfNegative(seriesIndex);

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                ApplyDataLabelSeparator(EnsureDataLabels(series), separator);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets the separator for data labels in a series by name.
        /// </summary>
        public ExcelChart SetSeriesDataLabelSeparator(string seriesName, string? separator, bool ignoreCase = true) {
            ArgumentNullException.ThrowIfNull(seriesName);

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                ApplyDataLabelSeparator(EnsureDataLabels(series), separator);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets the separator for a specific data label point by series index.
        /// </summary>
        public ExcelChart SetSeriesDataLabelSeparatorForPoint(int seriesIndex, int pointIndex, string? separator) {
            ArgumentOutOfRangeException.ThrowIfNegative(seriesIndex);
            ArgumentOutOfRangeException.ThrowIfNegative(pointIndex);

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                OfficeXmlNode label = EnsureDataLabel(series, pointIndex);
                ApplyDataLabelSeparator(label, separator);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets the separator for a specific data label point by series name.
        /// </summary>
        public ExcelChart SetSeriesDataLabelSeparatorForPoint(string seriesName, int pointIndex, string? separator,
            bool ignoreCase = true) {
            ArgumentNullException.ThrowIfNull(seriesName);
            ArgumentOutOfRangeException.ThrowIfNegative(pointIndex);

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                OfficeXmlNode label = EnsureDataLabel(series, pointIndex);
                ApplyDataLabelSeparator(label, separator);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Configures data labels for a single series by index.
        /// </summary>
        public ExcelChart SetSeriesDataLabels(int seriesIndex, bool showValue = true, bool showCategoryName = false,
            bool showSeriesName = false, bool showLegendKey = false, bool showPercent = false,
            OfficeChartDataLabelPosition? position = null, string? numberFormat = null, bool sourceLinked = false) {
            ArgumentOutOfRangeException.ThrowIfNegative(seriesIndex);
            if (numberFormat != null && string.IsNullOrWhiteSpace(numberFormat)) {
                throw new ArgumentException("Number format cannot be empty.", nameof(numberFormat));
            }

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                ApplyDataLabels(series, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent,
                    position, numberFormat, numberFormat != null ? sourceLinked : (bool?)null);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        /// Configures data labels for a single series by name.
        /// </summary>
        public ExcelChart SetSeriesDataLabels(string seriesName, bool showValue = true, bool showCategoryName = false,
            bool showSeriesName = false, bool showLegendKey = false, bool showPercent = false,
            OfficeChartDataLabelPosition? position = null, string? numberFormat = null, bool sourceLinked = false,
            bool ignoreCase = true) {
            ArgumentNullException.ThrowIfNull(seriesName);
            if (numberFormat != null && string.IsNullOrWhiteSpace(numberFormat)) {
                throw new ArgumentException("Number format cannot be empty.", nameof(numberFormat));
            }

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                ApplyDataLabels(series, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent,
                    position, numberFormat, numberFormat != null ? sourceLinked : (bool?)null);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        // Data labels are set on these chart groups, in this order.
        private static readonly string[] DataLabelChartGroupNames = {
            "barChart", "bar3DChart", "lineChart", "line3DChart", "areaChart", "area3DChart", "pieChart", "pie3DChart",
            "ofPieChart", "doughnutChart", "scatterChart", "bubbleChart", "radarChart", "stockChart"
        };

        // Runs apply on every chart group that takes data labels; false when the chart has no plot area.
        private bool ForEachDataLabelChartGroup(Action<OfficeXmlNode> apply) {
            OfficeXmlNode plotArea = ChartNode().Child(C, "plotArea");
            if (plotArea.IsNone) {
                return false;
            }

            foreach (string chartGroupName in DataLabelChartGroupNames) {
                foreach (OfficeXmlNode chartGroup in plotArea.Children(C, chartGroupName)) {
                    apply(chartGroup);
                }
            }
            return true;
        }
    }
}
