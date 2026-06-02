using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.PowerPoint {
        /// <summary>
        ///     Represents a chart on a slide.
        /// </summary>
        public partial class PowerPointChart : PowerPointShape {
        /// <summary>
        ///     Configures data labels for all supported chart series.
        /// </summary>
        public PowerPointChart SetDataLabels(bool showValue = true, bool showCategoryName = false,
            bool showSeriesName = false, bool showLegendKey = false, bool showPercent = false) {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            foreach (C.BarChart barChart in plotArea.Elements<C.BarChart>()) {
                ApplyDataLabels(barChart, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent);
            }

            foreach (C.LineChart lineChart in plotArea.Elements<C.LineChart>()) {
                ApplyDataLabels(lineChart, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent);
            }

            foreach (C.AreaChart areaChart in plotArea.Elements<C.AreaChart>()) {
                ApplyDataLabels(areaChart, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent);
            }

            foreach (C.PieChart pieChart in plotArea.Elements<C.PieChart>()) {
                ApplyDataLabels(pieChart, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent);
            }

            foreach (C.DoughnutChart doughnutChart in plotArea.Elements<C.DoughnutChart>()) {
                ApplyDataLabels(doughnutChart, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent);
            }

            foreach (C.ScatterChart scatterChart in plotArea.Elements<C.ScatterChart>()) {
                ApplyDataLabels(scatterChart, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent);
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Removes shared data label settings from all supported chart types in the current plot area.
        /// </summary>
        public PowerPointChart ClearDataLabels() {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            foreach (C.BarChart barChart in plotArea.Elements<C.BarChart>()) {
                RemoveDataLabels(barChart);
            }

            foreach (C.LineChart lineChart in plotArea.Elements<C.LineChart>()) {
                RemoveDataLabels(lineChart);
            }

            foreach (C.AreaChart areaChart in plotArea.Elements<C.AreaChart>()) {
                RemoveDataLabels(areaChart);
            }

            foreach (C.PieChart pieChart in plotArea.Elements<C.PieChart>()) {
                RemoveDataLabels(pieChart);
            }

            foreach (C.DoughnutChart doughnutChart in plotArea.Elements<C.DoughnutChart>()) {
                RemoveDataLabels(doughnutChart);
            }

            foreach (C.ScatterChart scatterChart in plotArea.Elements<C.ScatterChart>()) {
                RemoveDataLabels(scatterChart);
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Sets the shared data label position for all supported chart types.
        /// </summary>
        public PowerPointChart SetDataLabelPosition(C.DataLabelPositionValues position) {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            foreach (C.BarChart barChart in plotArea.Elements<C.BarChart>()) {
                SetDataLabelPosition(EnsureDataLabels(barChart), position);
            }

            foreach (C.LineChart lineChart in plotArea.Elements<C.LineChart>()) {
                SetDataLabelPosition(EnsureDataLabels(lineChart), position);
            }

            foreach (C.AreaChart areaChart in plotArea.Elements<C.AreaChart>()) {
                SetDataLabelPosition(EnsureDataLabels(areaChart), position);
            }

            foreach (C.PieChart pieChart in plotArea.Elements<C.PieChart>()) {
                SetDataLabelPosition(EnsureDataLabels(pieChart), GetPowerPointCompatibleDataLabelPosition(pieChart, position));
            }

            foreach (C.DoughnutChart doughnutChart in plotArea.Elements<C.DoughnutChart>()) {
                SetDataLabelPosition(EnsureDataLabels(doughnutChart), GetPowerPointCompatibleDataLabelPosition(doughnutChart, position));
            }

            foreach (C.ScatterChart scatterChart in plotArea.Elements<C.ScatterChart>()) {
                SetDataLabelPosition(EnsureDataLabels(scatterChart), position);
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Sets the shared data label number format for all supported chart types.
        /// </summary>
        public PowerPointChart SetDataLabelNumberFormat(string formatCode, bool sourceLinked = false) {
            if (string.IsNullOrWhiteSpace(formatCode)) {
                throw new ArgumentException("Format code cannot be null or empty.", nameof(formatCode));
            }

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            foreach (C.BarChart barChart in plotArea.Elements<C.BarChart>()) {
                SetDataLabelNumberFormat(EnsureDataLabels(barChart), formatCode, sourceLinked);
            }

            foreach (C.LineChart lineChart in plotArea.Elements<C.LineChart>()) {
                SetDataLabelNumberFormat(EnsureDataLabels(lineChart), formatCode, sourceLinked);
            }

            foreach (C.AreaChart areaChart in plotArea.Elements<C.AreaChart>()) {
                SetDataLabelNumberFormat(EnsureDataLabels(areaChart), formatCode, sourceLinked);
            }

            foreach (C.PieChart pieChart in plotArea.Elements<C.PieChart>()) {
                SetDataLabelNumberFormat(EnsureDataLabels(pieChart), formatCode, sourceLinked);
            }

            foreach (C.DoughnutChart doughnutChart in plotArea.Elements<C.DoughnutChart>()) {
                SetDataLabelNumberFormat(EnsureDataLabels(doughnutChart), formatCode, sourceLinked);
            }

            foreach (C.ScatterChart scatterChart in plotArea.Elements<C.ScatterChart>()) {
                SetDataLabelNumberFormat(EnsureDataLabels(scatterChart), formatCode, sourceLinked);
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Sets data label text style for all chart labels.
        /// </summary>
        public PowerPointChart SetDataLabelTextStyle(double? fontSizePoints = null, bool? bold = null, bool? italic = null,
            string? color = null, string? fontName = null) {
            ValidateTextStyle(fontSizePoints, color, fontName);

            return ApplyToAllDataLabels(labels => {
                ApplyDataLabelTextStyle(labels, fontSizePoints, bold, italic, color, fontName);
            });
        }

        /// <summary>
        ///     Sets data label shape styling for all chart labels.
        /// </summary>
        public PowerPointChart SetDataLabelShapeStyle(string? fillColor = null, string? lineColor = null,
            double? lineWidthPoints = null, bool noFill = false, bool noLine = false) {
            ValidateAreaStyle(fillColor, lineColor, lineWidthPoints, noFill, noLine);

            return ApplyToAllDataLabels(labels => {
                ApplyDataLabelShapeStyle(labels, fillColor, lineColor, lineWidthPoints, noFill, noLine);
            });
        }

        /// <summary>
        ///     Configures data label leader lines for all chart labels.
        /// </summary>
        public PowerPointChart SetDataLabelLeaderLines(bool showLeaderLines = true, string? lineColor = null,
            double? lineWidthPoints = null) {
            ValidateDataLabelLeaderLines(lineColor, lineWidthPoints);

            return ApplyToAllDataLabels(labels => {
                ApplyDataLabelLeaderLines(labels, showLeaderLines, lineColor, lineWidthPoints);
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
            if (template == null) {
                throw new ArgumentNullException(nameof(template));
            }

            return ApplyToAllDataLabels(labels => {
                ApplyDataLabelTemplate(labels, template);
            });
        }

        /// <summary>
        ///     Enables callout-style labels by positioning labels outside with leader lines.
        /// </summary>
        public PowerPointChart SetDataLabelCallouts(bool enabled = true, C.DataLabelPositionValues? position = null,
            string? lineColor = null, double? lineWidthPoints = null) {
            C.DataLabelPositionValues? resolvedPosition = enabled ? position ?? C.DataLabelPositionValues.OutsideEnd : position;
            SetDataLabels(showValue: enabled, showCategoryName: false, showSeriesName: false, showLegendKey: false,
                showPercent: false);
            if (resolvedPosition != null) {
                SetDataLabelPosition(resolvedPosition.Value);
            }

            return SetDataLabelLeaderLines(enabled, lineColor, lineWidthPoints);
        }
    }
}
