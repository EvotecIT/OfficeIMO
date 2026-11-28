using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
        //     Represents a chart on a slide.
        public partial class PowerPointChart : PowerPointShape {
        /// <summary>
        ///     Configures data labels for all supported chart series.
        /// </summary>
        public PowerPointChart SetDataLabels(bool showValue = true, bool showCategoryName = false,
            bool showSeriesName = false, bool showLegendKey = false, bool showPercent = false) {
            OfficeXmlNode plotArea = PlotAreaNode;
            if (plotArea.IsNone) {
                return this;
            }

            foreach (OfficeXmlNode chartGroup in DataLabelGroups(plotArea)) {
                ApplyDataLabels(chartGroup, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent);
            }

            MarkChartEdited();
            return this;
        }

        /// <summary>
        ///     Removes shared data label settings from all supported chart types in the current plot area.
        /// </summary>
        public PowerPointChart ClearDataLabels() {
            OfficeXmlNode plotArea = PlotAreaNode;
            if (plotArea.IsNone) {
                return this;
            }

            foreach (OfficeXmlNode chartGroup in DataLabelGroups(plotArea)) {
                RemoveDataLabels(chartGroup);
            }

            MarkChartEdited();
            return this;
        }

        /// <summary>
        ///     Sets the shared data label position for all supported chart types.
        /// </summary>
        public PowerPointChart SetDataLabelPosition(OfficeChartDataLabelPosition position) {
            OfficeXmlNode plotArea = PlotAreaNode;
            if (plotArea.IsNone) {
                return this;
            }

            foreach (OfficeXmlNode chartGroup in DataLabelGroups(plotArea)) {
                OfficeChartDataLabelPosition? groupPosition = chartGroup.LocalName is "pieChart" or "doughnutChart"
                    ? GetPowerPointCompatibleDataLabelPosition(chartGroup, position)
                    : position;
                SetDataLabelPosition(EnsureDataLabels(chartGroup), groupPosition);
            }

            MarkChartEdited();
            return this;
        }

        /// <summary>
        ///     Sets the shared data label number format for all supported chart types.
        /// </summary>
        public PowerPointChart SetDataLabelNumberFormat(string formatCode, bool sourceLinked = false) {
            ArgumentException.ThrowIfNullOrWhiteSpace(formatCode);

            OfficeXmlNode plotArea = PlotAreaNode;
            if (plotArea.IsNone) {
                return this;
            }

            foreach (OfficeXmlNode chartGroup in DataLabelGroups(plotArea)) {
                SetDataLabelNumberFormat(EnsureDataLabels(chartGroup), formatCode, sourceLinked);
            }

            MarkChartEdited();
            return this;
        }

        /// <summary>
        ///     Sets data label text style for all chart labels.
        /// </summary>
        public PowerPointChart SetDataLabelTextStyle(OfficeLength? fontSize = null, bool? bold = null, bool? italic = null,
            OfficeColor? color = null, string? fontFamily = null) {
            ValidateTextStyle(fontSize, color, fontFamily);

            return ApplyToAllDataLabels(labels => {
                ApplyDataLabelTextStyle(labels, fontSize, bold, italic, color, fontFamily);
            });
        }

        /// <summary>
        ///     Sets data label shape styling for all chart labels.
        /// </summary>
        public PowerPointChart SetDataLabelShapeStyle(OfficeColor? fillColor = null, OfficeColor? lineColor = null,
            OfficeLength? lineWidth = null, bool noFill = false, bool noLine = false) {
            ValidateAreaStyle(fillColor, lineColor, lineWidth, noFill, noLine);

            return ApplyToAllDataLabels(labels => {
                ApplyDataLabelShapeStyle(labels, fillColor, lineColor, lineWidth, noFill, noLine);
            });
        }

        /// <summary>
        ///     Configures data label leader lines for all chart labels.
        /// </summary>
        public PowerPointChart SetDataLabelLeaderLines(bool showLeaderLines = true, OfficeColor? lineColor = null,
            OfficeLength? lineWidth = null) {
            ValidateDataLabelLeaderLines(lineColor, lineWidth);

            return ApplyToAllDataLabels(labels => {
                ApplyDataLabelLeaderLines(labels, showLeaderLines, lineColor, lineWidth);
            });
        }

        /// <summary>
        ///     Sets the data label separator for all chart labels.
        /// </summary>
        public PowerPointChart SetDataLabelSeparator(string? separator) {
            if (separator != null && string.IsNullOrWhiteSpace(separator)) {
                throw new ArgumentException("Separator cannot be empty.", nameof(separator));
            }

            return ApplyToAllDataLabels(labels => {
                ApplyDataLabelSeparator(labels, separator);
            });
        }

        /// <summary>
        ///     Applies a reusable data label template to all supported chart labels.
        /// </summary>
        public PowerPointChart SetDataLabelTemplate(PowerPointChartDataLabelTemplate template) {
            ArgumentNullException.ThrowIfNull(template);

            return ApplyToAllDataLabels(labels => {
                ApplyDataLabelsTemplate(labels, template);
            });
        }

        /// <summary>
        ///     Enables callout-style labels by positioning labels outside with leader lines.
        /// </summary>
        public PowerPointChart SetDataLabelCallouts(bool enabled = true, OfficeChartDataLabelPosition? position = null,
            OfficeColor? lineColor = null, OfficeLength? lineWidth = null) {
            OfficeChartDataLabelPosition? resolvedPosition = enabled ? position ?? OfficeChartDataLabelPosition.OutsideEnd : position;
            SetDataLabels(showValue: enabled, showCategoryName: false, showSeriesName: false, showLegendKey: false,
                showPercent: false);
            if (resolvedPosition != null) {
                SetDataLabelPosition(resolvedPosition.Value);
            }

            return SetDataLabelLeaderLines(enabled, lineColor, lineWidth);
        }
    }
}
