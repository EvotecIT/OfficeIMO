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

        /// <summary>
        ///     Configures data labels for a single series by index.
        /// </summary>
        public PowerPointChart SetSeriesDataLabels(int seriesIndex, bool showValue = true, bool showCategoryName = false,
            bool showSeriesName = false, bool showLegendKey = false, bool showPercent = false,
            C.DataLabelPositionValues? position = null, string? numberFormat = null, bool sourceLinked = false) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            if (numberFormat != null && string.IsNullOrWhiteSpace(numberFormat)) {
                throw new ArgumentException("Number format cannot be empty.", nameof(numberFormat));
            }

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                ApplyDataLabelOverrides(EnsureDataLabels(series), showLegendKey, showValue, showCategoryName, showSeriesName,
                    showPercent, position, numberFormat, sourceLinked);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Configures data labels for a single series by name.
        /// </summary>
        public PowerPointChart SetSeriesDataLabels(string seriesName, bool showValue = true, bool showCategoryName = false,
            bool showSeriesName = false, bool showLegendKey = false, bool showPercent = false,
            C.DataLabelPositionValues? position = null, string? numberFormat = null, bool sourceLinked = false,
            bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            if (numberFormat != null && string.IsNullOrWhiteSpace(numberFormat)) {
                throw new ArgumentException("Number format cannot be empty.", nameof(numberFormat));
            }

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                ApplyDataLabelOverrides(EnsureDataLabels(series), showLegendKey, showValue, showCategoryName, showSeriesName,
                    showPercent, position, numberFormat, sourceLinked);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Enables callout-style labels for a series by index.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelCallouts(int seriesIndex, bool enabled = true,
            C.DataLabelPositionValues? position = null, string? lineColor = null, double? lineWidthPoints = null) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }

            C.DataLabelPositionValues? resolvedPosition = enabled ? position ?? C.DataLabelPositionValues.OutsideEnd : position;
            SetSeriesDataLabels(seriesIndex, showValue: enabled, showCategoryName: false, showSeriesName: false,
                showLegendKey: false, showPercent: false, position: resolvedPosition, numberFormat: null, sourceLinked: false);
            return SetSeriesDataLabelLeaderLines(seriesIndex, enabled, lineColor, lineWidthPoints);
        }

        /// <summary>
        ///     Enables callout-style labels for a series by name.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelCallouts(string seriesName, bool enabled = true,
            C.DataLabelPositionValues? position = null, string? lineColor = null, double? lineWidthPoints = null,
            bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }

            C.DataLabelPositionValues? resolvedPosition = enabled ? position ?? C.DataLabelPositionValues.OutsideEnd : position;
            SetSeriesDataLabels(seriesName, showValue: enabled, showCategoryName: false, showSeriesName: false,
                showLegendKey: false, showPercent: false, position: resolvedPosition, numberFormat: null, sourceLinked: false,
                ignoreCase: ignoreCase);
            return SetSeriesDataLabelLeaderLines(seriesName, enabled, lineColor, lineWidthPoints, ignoreCase);
        }

        /// <summary>
        ///     Sets data label text style for a series by index.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelTextStyle(int seriesIndex, double? fontSizePoints = null, bool? bold = null,
            bool? italic = null, string? color = null, string? fontName = null) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            ValidateTextStyle(fontSizePoints, color, fontName);

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                ApplyDataLabelTextStyle(EnsureDataLabels(series), fontSizePoints, bold, italic, color, fontName);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Sets data label text style for a series by name.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelTextStyle(string seriesName, double? fontSizePoints = null, bool? bold = null,
            bool? italic = null, string? color = null, string? fontName = null, bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            ValidateTextStyle(fontSizePoints, color, fontName);

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                ApplyDataLabelTextStyle(EnsureDataLabels(series), fontSizePoints, bold, italic, color, fontName);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Sets data label shape styling for a series by index.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelShapeStyle(int seriesIndex, string? fillColor = null, string? lineColor = null,
            double? lineWidthPoints = null, bool noFill = false, bool noLine = false) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            ValidateAreaStyle(fillColor, lineColor, lineWidthPoints, noFill, noLine);

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                ApplyDataLabelShapeStyle(EnsureDataLabels(series), fillColor, lineColor, lineWidthPoints, noFill, noLine);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Sets data label shape styling for a series by name.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelShapeStyle(string seriesName, string? fillColor = null, string? lineColor = null,
            double? lineWidthPoints = null, bool noFill = false, bool noLine = false, bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            ValidateAreaStyle(fillColor, lineColor, lineWidthPoints, noFill, noLine);

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                ApplyDataLabelShapeStyle(EnsureDataLabels(series), fillColor, lineColor, lineWidthPoints, noFill, noLine);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Configures data label leader lines for a series by index.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelLeaderLines(int seriesIndex, bool showLeaderLines = true, string? lineColor = null,
            double? lineWidthPoints = null) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            ValidateDataLabelLeaderLines(lineColor, lineWidthPoints);

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                ApplyDataLabelLeaderLines(EnsureDataLabels(series), showLeaderLines, lineColor, lineWidthPoints);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Configures data label leader lines for a series by name.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelLeaderLines(string seriesName, bool showLeaderLines = true, string? lineColor = null,
            double? lineWidthPoints = null, bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            ValidateDataLabelLeaderLines(lineColor, lineWidthPoints);

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                ApplyDataLabelLeaderLines(EnsureDataLabels(series), showLeaderLines, lineColor, lineWidthPoints);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Sets the data label separator for a series by index.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelSeparator(int seriesIndex, string? separator) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            if (separator != null && string.IsNullOrWhiteSpace(separator)) {
                throw new ArgumentException("Separator cannot be empty.", nameof(separator));
            }

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
        ///     Sets the data label separator for a series by name.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelSeparator(string seriesName, string? separator, bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            if (separator != null && string.IsNullOrWhiteSpace(separator)) {
                throw new ArgumentException("Separator cannot be empty.", nameof(separator));
            }

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
        ///     Applies a reusable data label template to a series by index.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelTemplate(int seriesIndex, PowerPointChartDataLabelTemplate template) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            if (template == null) {
                throw new ArgumentNullException(nameof(template));
            }

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
        ///     Applies a reusable data label template to a series by name.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelTemplate(string seriesName, PowerPointChartDataLabelTemplate template,
            bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            if (template == null) {
                throw new ArgumentNullException(nameof(template));
            }

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
        ///     Removes series-level data label settings by series index.
        /// </summary>
        public PowerPointChart ClearSeriesDataLabels(int seriesIndex) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }

            bool applied = ApplySeriesByIndex(seriesIndex, RemoveDataLabels);

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Removes series-level data label settings by series name.
        /// </summary>
        public PowerPointChart ClearSeriesDataLabels(string seriesName, bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }

            bool applied = ApplySeriesByName(seriesName, ignoreCase, RemoveDataLabels);

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Applies a reusable data label template to a specific point by series index.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelTemplateForPoint(int seriesIndex, int pointIndex,
            PowerPointChartDataLabelTemplate template) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            if (pointIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(pointIndex));
            }
            if (template == null) {
                throw new ArgumentNullException(nameof(template));
            }

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                ApplySeriesLeaderLineTemplate(series, template);
                ApplyDataLabelTemplate(EnsureDataLabel(series, pointIndex), template);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Applies a reusable data label template to a specific point by series name.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelTemplateForPoint(string seriesName, int pointIndex,
            PowerPointChartDataLabelTemplate template, bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            if (pointIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(pointIndex));
            }
            if (template == null) {
                throw new ArgumentNullException(nameof(template));
            }

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                ApplySeriesLeaderLineTemplate(series, template);
                ApplyDataLabelTemplate(EnsureDataLabel(series, pointIndex), template);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Removes point-level data label overrides for a specific point by series index.
        /// </summary>
        public PowerPointChart ClearSeriesDataLabelForPoint(int seriesIndex, int pointIndex) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            if (pointIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(pointIndex));
            }

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                ClearDataLabel(series, pointIndex);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Removes point-level data label overrides for a specific point by series name.
        /// </summary>
        public PowerPointChart ClearSeriesDataLabelForPoint(string seriesName, int pointIndex, bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            if (pointIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(pointIndex));
            }

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                ClearDataLabel(series, pointIndex);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Configures a single data label point by series index and point index.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelForPoint(int seriesIndex, int pointIndex, bool? showValue = null,
            bool? showCategoryName = null, bool? showSeriesName = null, bool? showLegendKey = null,
            bool? showPercent = null, C.DataLabelPositionValues? position = null, string? numberFormat = null,
            bool sourceLinked = false) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            if (pointIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(pointIndex));
            }
            if (numberFormat != null && string.IsNullOrWhiteSpace(numberFormat)) {
                throw new ArgumentException("Number format cannot be empty.", nameof(numberFormat));
            }

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                C.DataLabel label = EnsureDataLabel(series, pointIndex);
                ApplyDataLabelOverrides(label, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent,
                    position, numberFormat, sourceLinked);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Configures a single data label point by series name and point index.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelForPoint(string seriesName, int pointIndex, bool? showValue = null,
            bool? showCategoryName = null, bool? showSeriesName = null, bool? showLegendKey = null,
            bool? showPercent = null, C.DataLabelPositionValues? position = null, string? numberFormat = null,
            bool sourceLinked = false, bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            if (pointIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(pointIndex));
            }
            if (numberFormat != null && string.IsNullOrWhiteSpace(numberFormat)) {
                throw new ArgumentException("Number format cannot be empty.", nameof(numberFormat));
            }

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                C.DataLabel label = EnsureDataLabel(series, pointIndex);
                ApplyDataLabelOverrides(label, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent,
                    position, numberFormat, sourceLinked);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Sets data label text style for a specific point by series index.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelTextStyleForPoint(int seriesIndex, int pointIndex,
            double? fontSizePoints = null, bool? bold = null, bool? italic = null,
            string? color = null, string? fontName = null) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            if (pointIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(pointIndex));
            }
            ValidateTextStyle(fontSizePoints, color, fontName);

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                C.DataLabel label = EnsureDataLabel(series, pointIndex);
                ApplyDataLabelTextStyle(label, fontSizePoints, bold, italic, color, fontName);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Sets data label text style for a specific point by series name.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelTextStyleForPoint(string seriesName, int pointIndex,
            double? fontSizePoints = null, bool? bold = null, bool? italic = null,
            string? color = null, string? fontName = null, bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            if (pointIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(pointIndex));
            }
            ValidateTextStyle(fontSizePoints, color, fontName);

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                C.DataLabel label = EnsureDataLabel(series, pointIndex);
                ApplyDataLabelTextStyle(label, fontSizePoints, bold, italic, color, fontName);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Sets data label shape styling for a specific point by series index.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelShapeStyleForPoint(int seriesIndex, int pointIndex,
            string? fillColor = null, string? lineColor = null, double? lineWidthPoints = null,
            bool noFill = false, bool noLine = false) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            if (pointIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(pointIndex));
            }
            ValidateAreaStyle(fillColor, lineColor, lineWidthPoints, noFill, noLine);

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                C.DataLabel label = EnsureDataLabel(series, pointIndex);
                ApplyDataLabelShapeStyle(label, fillColor, lineColor, lineWidthPoints, noFill, noLine);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Sets data label shape styling for a specific point by series name.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelShapeStyleForPoint(string seriesName, int pointIndex,
            string? fillColor = null, string? lineColor = null, double? lineWidthPoints = null,
            bool noFill = false, bool noLine = false, bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            if (pointIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(pointIndex));
            }
            ValidateAreaStyle(fillColor, lineColor, lineWidthPoints, noFill, noLine);

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                C.DataLabel label = EnsureDataLabel(series, pointIndex);
                ApplyDataLabelShapeStyle(label, fillColor, lineColor, lineWidthPoints, noFill, noLine);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Enables callout-style labels for a specific point by series index.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelCalloutsForPoint(int seriesIndex, int pointIndex, bool enabled = true,
            C.DataLabelPositionValues? position = null) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            if (pointIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(pointIndex));
            }

            C.DataLabelPositionValues? resolvedPosition = enabled ? position ?? C.DataLabelPositionValues.OutsideEnd : position;
            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                C.DataLabel label = EnsureDataLabel(series, pointIndex);
                ApplyDataLabelOverrides(label, showLegendKey: false, showValue: enabled, showCategoryName: false,
                    showSeriesName: false, showPercent: false, position: resolvedPosition, numberFormat: null,
                    sourceLinked: false);
                ApplyDataLabelLeaderLines(EnsureDataLabels(series), enabled, lineColor: null, lineWidthPoints: null);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Enables callout-style labels for a specific point by series name.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelCalloutsForPoint(string seriesName, int pointIndex, bool enabled = true,
            C.DataLabelPositionValues? position = null,
            bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            if (pointIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(pointIndex));
            }

            C.DataLabelPositionValues? resolvedPosition = enabled ? position ?? C.DataLabelPositionValues.OutsideEnd : position;
            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                C.DataLabel label = EnsureDataLabel(series, pointIndex);
                ApplyDataLabelOverrides(label, showLegendKey: false, showValue: enabled, showCategoryName: false,
                    showSeriesName: false, showPercent: false, position: resolvedPosition, numberFormat: null,
                    sourceLinked: false);
                ApplyDataLabelLeaderLines(EnsureDataLabels(series), enabled, lineColor: null, lineWidthPoints: null);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Sets the data label separator for a specific point by series index.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelSeparatorForPoint(int seriesIndex, int pointIndex, string? separator) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            if (pointIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(pointIndex));
            }
            if (separator != null && string.IsNullOrWhiteSpace(separator)) {
                throw new ArgumentException("Separator cannot be empty.", nameof(separator));
            }

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                C.DataLabel label = EnsureDataLabel(series, pointIndex);
                ApplyDataLabelSeparator(label, separator);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series index {seriesIndex} was not found.");
            }

            Save();
            return this;
        }

        /// <summary>
        ///     Sets the data label separator for a specific point by series name.
        /// </summary>
        public PowerPointChart SetSeriesDataLabelSeparatorForPoint(string seriesName, int pointIndex, string? separator,
            bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            if (pointIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(pointIndex));
            }
            if (separator != null && string.IsNullOrWhiteSpace(separator)) {
                throw new ArgumentException("Separator cannot be empty.", nameof(separator));
            }

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                C.DataLabel label = EnsureDataLabel(series, pointIndex);
                ApplyDataLabelSeparator(label, separator);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }
    }
}
