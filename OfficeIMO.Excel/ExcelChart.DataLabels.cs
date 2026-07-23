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
        /// Configures data labels for all supported chart series.
        /// </summary>
        public ExcelChart SetDataLabels(bool showValue = true, bool showCategoryName = false,
            bool showSeriesName = false, bool showLegendKey = false, bool showPercent = false) {
            return SetDataLabels(showValue, showCategoryName, showSeriesName, showLegendKey, showPercent, null, null, false);
        }

        /// <summary>
        /// Configures data labels for all supported chart series with optional formatting.
        /// </summary>
        public ExcelChart SetDataLabels(bool showValue, bool showCategoryName,
            bool showSeriesName, bool showLegendKey, bool showPercent,
            C.DataLabelPositionValues? position, string? numberFormat, bool sourceLinked = false) {
            if (numberFormat != null && string.IsNullOrWhiteSpace(numberFormat)) {
                throw new ArgumentException("Number format cannot be empty.", nameof(numberFormat));
            }
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            foreach (C.BarChart barChart in plotArea.Elements<C.BarChart>()) {
                ApplyDataLabels(barChart, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent,
                    position, numberFormat, sourceLinked);
            }
            foreach (C.Bar3DChart barChart in plotArea.Elements<C.Bar3DChart>()) {
                ApplyDataLabels(barChart, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent,
                    position, numberFormat, sourceLinked);
            }

            foreach (C.LineChart lineChart in plotArea.Elements<C.LineChart>()) {
                ApplyDataLabels(lineChart, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent,
                    position, numberFormat, sourceLinked);
            }
            foreach (C.Line3DChart lineChart in plotArea.Elements<C.Line3DChart>()) {
                ApplyDataLabels(lineChart, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent,
                    position, numberFormat, sourceLinked);
            }

            foreach (C.AreaChart areaChart in plotArea.Elements<C.AreaChart>()) {
                ApplyDataLabels(areaChart, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent,
                    position, numberFormat, sourceLinked);
            }
            foreach (C.Area3DChart areaChart in plotArea.Elements<C.Area3DChart>()) {
                ApplyDataLabels(areaChart, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent,
                    position, numberFormat, sourceLinked);
            }

            foreach (C.PieChart pieChart in plotArea.Elements<C.PieChart>()) {
                ApplyDataLabels(pieChart, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent,
                    position, numberFormat, sourceLinked);
            }
            foreach (C.Pie3DChart pieChart in plotArea.Elements<C.Pie3DChart>()) {
                ApplyDataLabels(pieChart, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent,
                    position, numberFormat, sourceLinked);
            }
            foreach (C.OfPieChart pieChart in plotArea.Elements<C.OfPieChart>()) {
                ApplyDataLabels(pieChart, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent,
                    position, numberFormat, sourceLinked);
            }

            foreach (C.DoughnutChart doughnutChart in plotArea.Elements<C.DoughnutChart>()) {
                ApplyDataLabels(doughnutChart, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent,
                    position, numberFormat, sourceLinked);
            }

            foreach (C.ScatterChart scatterChart in plotArea.Elements<C.ScatterChart>()) {
                ApplyDataLabels(scatterChart, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent,
                    position, numberFormat, sourceLinked);
            }

            foreach (C.BubbleChart bubbleChart in plotArea.Elements<C.BubbleChart>()) {
                ApplyDataLabels(bubbleChart, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent,
                    position, numberFormat, sourceLinked);
            }
            foreach (C.RadarChart radarChart in plotArea.Elements<C.RadarChart>()) {
                ApplyDataLabels(radarChart, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent,
                    position, numberFormat, sourceLinked);
            }
            foreach (C.StockChart stockChart in plotArea.Elements<C.StockChart>()) {
                ApplyDataLabels(stockChart, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent,
                    position, numberFormat, sourceLinked);
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets data label text style for all chart labels.
        /// </summary>
        public ExcelChart SetDataLabelTextStyle(double? fontSizePoints = null, bool? bold = null, bool? italic = null,
            string? color = null, string? fontName = null) {
            ValidateDataLabelTextStyle(fontSizePoints, color, fontName);

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            foreach (C.BarChart barChart in plotArea.Elements<C.BarChart>()) {
                ApplyDataLabelTextStyle(EnsureDataLabels(barChart), fontSizePoints, bold, italic, color, fontName);
            }
            foreach (C.Bar3DChart barChart in plotArea.Elements<C.Bar3DChart>()) {
                ApplyDataLabelTextStyle(EnsureDataLabels(barChart), fontSizePoints, bold, italic, color, fontName);
            }
            foreach (C.LineChart lineChart in plotArea.Elements<C.LineChart>()) {
                ApplyDataLabelTextStyle(EnsureDataLabels(lineChart), fontSizePoints, bold, italic, color, fontName);
            }
            foreach (C.Line3DChart lineChart in plotArea.Elements<C.Line3DChart>()) {
                ApplyDataLabelTextStyle(EnsureDataLabels(lineChart), fontSizePoints, bold, italic, color, fontName);
            }
            foreach (C.AreaChart areaChart in plotArea.Elements<C.AreaChart>()) {
                ApplyDataLabelTextStyle(EnsureDataLabels(areaChart), fontSizePoints, bold, italic, color, fontName);
            }
            foreach (C.Area3DChart areaChart in plotArea.Elements<C.Area3DChart>()) {
                ApplyDataLabelTextStyle(EnsureDataLabels(areaChart), fontSizePoints, bold, italic, color, fontName);
            }
            foreach (C.PieChart pieChart in plotArea.Elements<C.PieChart>()) {
                ApplyDataLabelTextStyle(EnsureDataLabels(pieChart), fontSizePoints, bold, italic, color, fontName);
            }
            foreach (C.Pie3DChart pieChart in plotArea.Elements<C.Pie3DChart>()) {
                ApplyDataLabelTextStyle(EnsureDataLabels(pieChart), fontSizePoints, bold, italic, color, fontName);
            }
            foreach (C.OfPieChart pieChart in plotArea.Elements<C.OfPieChart>()) {
                ApplyDataLabelTextStyle(EnsureDataLabels(pieChart), fontSizePoints, bold, italic, color, fontName);
            }
            foreach (C.DoughnutChart doughnutChart in plotArea.Elements<C.DoughnutChart>()) {
                ApplyDataLabelTextStyle(EnsureDataLabels(doughnutChart), fontSizePoints, bold, italic, color, fontName);
            }
            foreach (C.ScatterChart scatterChart in plotArea.Elements<C.ScatterChart>()) {
                ApplyDataLabelTextStyle(EnsureDataLabels(scatterChart), fontSizePoints, bold, italic, color, fontName);
            }
            foreach (C.BubbleChart bubbleChart in plotArea.Elements<C.BubbleChart>()) {
                ApplyDataLabelTextStyle(EnsureDataLabels(bubbleChart), fontSizePoints, bold, italic, color, fontName);
            }
            foreach (C.RadarChart radarChart in plotArea.Elements<C.RadarChart>()) {
                ApplyDataLabelTextStyle(EnsureDataLabels(radarChart), fontSizePoints, bold, italic, color, fontName);
            }
            foreach (C.StockChart stockChart in plotArea.Elements<C.StockChart>()) {
                ApplyDataLabelTextStyle(EnsureDataLabels(stockChart), fontSizePoints, bold, italic, color, fontName);
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets data label shape styling for all chart labels.
        /// </summary>
        public ExcelChart SetDataLabelShapeStyle(string? fillColor = null, string? lineColor = null, double? lineWidthPoints = null,
            bool noFill = false, bool noLine = false) {
            ValidateDataLabelShapeStyle(fillColor, lineColor, lineWidthPoints, noFill, noLine);

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            foreach (C.BarChart barChart in plotArea.Elements<C.BarChart>()) {
                ApplyDataLabelShapeStyle(EnsureDataLabels(barChart), fillColor, lineColor, lineWidthPoints, noFill, noLine);
            }
            foreach (C.Bar3DChart barChart in plotArea.Elements<C.Bar3DChart>()) {
                ApplyDataLabelShapeStyle(EnsureDataLabels(barChart), fillColor, lineColor, lineWidthPoints, noFill, noLine);
            }
            foreach (C.LineChart lineChart in plotArea.Elements<C.LineChart>()) {
                ApplyDataLabelShapeStyle(EnsureDataLabels(lineChart), fillColor, lineColor, lineWidthPoints, noFill, noLine);
            }
            foreach (C.Line3DChart lineChart in plotArea.Elements<C.Line3DChart>()) {
                ApplyDataLabelShapeStyle(EnsureDataLabels(lineChart), fillColor, lineColor, lineWidthPoints, noFill, noLine);
            }
            foreach (C.AreaChart areaChart in plotArea.Elements<C.AreaChart>()) {
                ApplyDataLabelShapeStyle(EnsureDataLabels(areaChart), fillColor, lineColor, lineWidthPoints, noFill, noLine);
            }
            foreach (C.Area3DChart areaChart in plotArea.Elements<C.Area3DChart>()) {
                ApplyDataLabelShapeStyle(EnsureDataLabels(areaChart), fillColor, lineColor, lineWidthPoints, noFill, noLine);
            }
            foreach (C.PieChart pieChart in plotArea.Elements<C.PieChart>()) {
                ApplyDataLabelShapeStyle(EnsureDataLabels(pieChart), fillColor, lineColor, lineWidthPoints, noFill, noLine);
            }
            foreach (C.Pie3DChart pieChart in plotArea.Elements<C.Pie3DChart>()) {
                ApplyDataLabelShapeStyle(EnsureDataLabels(pieChart), fillColor, lineColor, lineWidthPoints, noFill, noLine);
            }
            foreach (C.OfPieChart pieChart in plotArea.Elements<C.OfPieChart>()) {
                ApplyDataLabelShapeStyle(EnsureDataLabels(pieChart), fillColor, lineColor, lineWidthPoints, noFill, noLine);
            }
            foreach (C.DoughnutChart doughnutChart in plotArea.Elements<C.DoughnutChart>()) {
                ApplyDataLabelShapeStyle(EnsureDataLabels(doughnutChart), fillColor, lineColor, lineWidthPoints, noFill, noLine);
            }
            foreach (C.ScatterChart scatterChart in plotArea.Elements<C.ScatterChart>()) {
                ApplyDataLabelShapeStyle(EnsureDataLabels(scatterChart), fillColor, lineColor, lineWidthPoints, noFill, noLine);
            }
            foreach (C.BubbleChart bubbleChart in plotArea.Elements<C.BubbleChart>()) {
                ApplyDataLabelShapeStyle(EnsureDataLabels(bubbleChart), fillColor, lineColor, lineWidthPoints, noFill, noLine);
            }
            foreach (C.RadarChart radarChart in plotArea.Elements<C.RadarChart>()) {
                ApplyDataLabelShapeStyle(EnsureDataLabels(radarChart), fillColor, lineColor, lineWidthPoints, noFill, noLine);
            }
            foreach (C.StockChart stockChart in plotArea.Elements<C.StockChart>()) {
                ApplyDataLabelShapeStyle(EnsureDataLabels(stockChart), fillColor, lineColor, lineWidthPoints, noFill, noLine);
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets data label shape styling for a series by index.
        /// </summary>
        public ExcelChart SetSeriesDataLabelShapeStyle(int seriesIndex, string? fillColor = null, string? lineColor = null,
            double? lineWidthPoints = null, bool noFill = false, bool noLine = false) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            ValidateDataLabelShapeStyle(fillColor, lineColor, lineWidthPoints, noFill, noLine);

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
        /// Sets data label shape styling for a series by name.
        /// </summary>
        public ExcelChart SetSeriesDataLabelShapeStyle(string seriesName, string? fillColor = null, string? lineColor = null,
            double? lineWidthPoints = null, bool noFill = false, bool noLine = false, bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            ValidateDataLabelShapeStyle(fillColor, lineColor, lineWidthPoints, noFill, noLine);

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
        /// Configures data label leader lines for all chart labels.
        /// </summary>
        public ExcelChart SetDataLabelLeaderLines(bool showLeaderLines = true, string? lineColor = null, double? lineWidthPoints = null) {
            ValidateDataLabelLeaderLines(lineColor, lineWidthPoints);

            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            foreach (C.BarChart barChart in plotArea.Elements<C.BarChart>()) {
                ApplyDataLabelLeaderLines(EnsureDataLabels(barChart), showLeaderLines, lineColor, lineWidthPoints);
            }
            foreach (C.Bar3DChart barChart in plotArea.Elements<C.Bar3DChart>()) {
                ApplyDataLabelLeaderLines(EnsureDataLabels(barChart), showLeaderLines, lineColor, lineWidthPoints);
            }
            foreach (C.LineChart lineChart in plotArea.Elements<C.LineChart>()) {
                ApplyDataLabelLeaderLines(EnsureDataLabels(lineChart), showLeaderLines, lineColor, lineWidthPoints);
            }
            foreach (C.Line3DChart lineChart in plotArea.Elements<C.Line3DChart>()) {
                ApplyDataLabelLeaderLines(EnsureDataLabels(lineChart), showLeaderLines, lineColor, lineWidthPoints);
            }
            foreach (C.AreaChart areaChart in plotArea.Elements<C.AreaChart>()) {
                ApplyDataLabelLeaderLines(EnsureDataLabels(areaChart), showLeaderLines, lineColor, lineWidthPoints);
            }
            foreach (C.Area3DChart areaChart in plotArea.Elements<C.Area3DChart>()) {
                ApplyDataLabelLeaderLines(EnsureDataLabels(areaChart), showLeaderLines, lineColor, lineWidthPoints);
            }
            foreach (C.PieChart pieChart in plotArea.Elements<C.PieChart>()) {
                ApplyDataLabelLeaderLines(EnsureDataLabels(pieChart), showLeaderLines, lineColor, lineWidthPoints);
            }
            foreach (C.Pie3DChart pieChart in plotArea.Elements<C.Pie3DChart>()) {
                ApplyDataLabelLeaderLines(EnsureDataLabels(pieChart), showLeaderLines, lineColor, lineWidthPoints);
            }
            foreach (C.OfPieChart pieChart in plotArea.Elements<C.OfPieChart>()) {
                ApplyDataLabelLeaderLines(EnsureDataLabels(pieChart), showLeaderLines, lineColor, lineWidthPoints);
            }
            foreach (C.DoughnutChart doughnutChart in plotArea.Elements<C.DoughnutChart>()) {
                ApplyDataLabelLeaderLines(EnsureDataLabels(doughnutChart), showLeaderLines, lineColor, lineWidthPoints);
            }
            foreach (C.ScatterChart scatterChart in plotArea.Elements<C.ScatterChart>()) {
                ApplyDataLabelLeaderLines(EnsureDataLabels(scatterChart), showLeaderLines, lineColor, lineWidthPoints);
            }
            foreach (C.BubbleChart bubbleChart in plotArea.Elements<C.BubbleChart>()) {
                ApplyDataLabelLeaderLines(EnsureDataLabels(bubbleChart), showLeaderLines, lineColor, lineWidthPoints);
            }
            foreach (C.RadarChart radarChart in plotArea.Elements<C.RadarChart>()) {
                ApplyDataLabelLeaderLines(EnsureDataLabels(radarChart), showLeaderLines, lineColor, lineWidthPoints);
            }
            foreach (C.StockChart stockChart in plotArea.Elements<C.StockChart>()) {
                ApplyDataLabelLeaderLines(EnsureDataLabels(stockChart), showLeaderLines, lineColor, lineWidthPoints);
            }

            Save();
            return this;
        }

        /// <summary>
        /// Configures data label leader lines for a series by index.
        /// </summary>
        public ExcelChart SetSeriesDataLabelLeaderLines(int seriesIndex, bool showLeaderLines = true, string? lineColor = null, double? lineWidthPoints = null) {
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
        /// Configures data label leader lines for a series by name.
        /// </summary>
        public ExcelChart SetSeriesDataLabelLeaderLines(string seriesName, bool showLeaderLines = true, string? lineColor = null,
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
        /// Enables callout-style labels by positioning labels outside with leader lines.
        /// </summary>
        public ExcelChart SetDataLabelCallouts(bool enabled = true, C.DataLabelPositionValues? position = null,
            string? lineColor = null, double? lineWidthPoints = null) {
            var resolvedPosition = enabled ? position ?? C.DataLabelPositionValues.OutsideEnd : position;
            SetDataLabels(showValue: enabled, showCategoryName: false, showSeriesName: false, showLegendKey: false,
                showPercent: false, position: resolvedPosition, numberFormat: null, sourceLinked: false);
            return SetDataLabelLeaderLines(enabled, lineColor, lineWidthPoints);
        }

        /// <summary>
        /// Enables callout-style labels for a series by index.
        /// </summary>
        public ExcelChart SetSeriesDataLabelCallouts(int seriesIndex, bool enabled = true, C.DataLabelPositionValues? position = null,
            string? lineColor = null, double? lineWidthPoints = null) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            var resolvedPosition = enabled ? position ?? C.DataLabelPositionValues.OutsideEnd : position;
            SetSeriesDataLabels(seriesIndex, showValue: enabled, showCategoryName: false, showSeriesName: false,
                showLegendKey: false, showPercent: false, position: resolvedPosition, numberFormat: null, sourceLinked: false);
            return SetSeriesDataLabelLeaderLines(seriesIndex, enabled, lineColor, lineWidthPoints);
        }

        /// <summary>
        /// Enables callout-style labels for a series by name.
        /// </summary>
        public ExcelChart SetSeriesDataLabelCallouts(string seriesName, bool enabled = true, C.DataLabelPositionValues? position = null,
            string? lineColor = null, double? lineWidthPoints = null, bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            var resolvedPosition = enabled ? position ?? C.DataLabelPositionValues.OutsideEnd : position;
            SetSeriesDataLabels(seriesName, showValue: enabled, showCategoryName: false, showSeriesName: false,
                showLegendKey: false, showPercent: false, position: resolvedPosition, numberFormat: null, sourceLinked: false,
                ignoreCase: ignoreCase);
            return SetSeriesDataLabelLeaderLines(seriesName, enabled, lineColor, lineWidthPoints, ignoreCase);
        }

        /// <summary>
        /// Sets data label text style for a series by index.
        /// </summary>
        public ExcelChart SetSeriesDataLabelTextStyle(int seriesIndex, double? fontSizePoints = null, bool? bold = null,
            bool? italic = null, string? color = null, string? fontName = null) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            ValidateDataLabelTextStyle(fontSizePoints, color, fontName);

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
        /// Sets data label text style for a series by name.
        /// </summary>
        public ExcelChart SetSeriesDataLabelTextStyle(string seriesName, double? fontSizePoints = null, bool? bold = null,
            bool? italic = null, string? color = null, string? fontName = null, bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            ValidateDataLabelTextStyle(fontSizePoints, color, fontName);

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
        /// Applies a reusable data label template to a series by index.
        /// </summary>
        public ExcelChart SetSeriesDataLabelTemplate(int seriesIndex, ExcelChartDataLabelTemplate template) {
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
        /// Applies a reusable data label template to a series by name.
        /// </summary>
        public ExcelChart SetSeriesDataLabelTemplate(string seriesName, ExcelChartDataLabelTemplate template,
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
        /// Configures a single data label point by series index and point index.
        /// </summary>
        public ExcelChart SetSeriesDataLabelForPoint(int seriesIndex, int pointIndex, bool? showValue = null,
            bool? showCategoryName = null, bool? showSeriesName = null, bool? showLegendKey = null,
            bool? showPercent = null, bool? showBubbleSize = null,
            C.DataLabelPositionValues? position = null, string? numberFormat = null, bool sourceLinked = false) {
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
                    showBubbleSize, position, numberFormat, sourceLinked);
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
            C.DataLabelPositionValues? position = null, string? numberFormat = null, bool sourceLinked = false) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            if (pointIndices == null) {
                throw new ArgumentNullException(nameof(pointIndices));
            }
            if (numberFormat != null && string.IsNullOrWhiteSpace(numberFormat)) {
                throw new ArgumentException("Number format cannot be empty.", nameof(numberFormat));
            }

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                C.DataLabels labels = EnsureDataLabels(series);
                var labelsByIndex = new Dictionary<uint, C.DataLabel>();
                foreach (C.DataLabel existing in labels.Elements<C.DataLabel>()) {
                    uint? existingIndex = existing.GetFirstChild<C.Index>()?.Val?.Value;
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
                    if (!labelsByIndex.TryGetValue(index, out C.DataLabel? label)) {
                        label = new C.DataLabel(new C.Index { Val = index });
                        OpenXmlElement? insertBefore = labels.ChildElements.FirstOrDefault(child => child is not C.DataLabel);
                        if (insertBefore != null) {
                            labels.InsertBefore(label, insertBefore);
                        } else {
                            labels.Append(label);
                        }
                        labelsByIndex.Add(index, label);
                    }

                    ApplyDataLabelOverrides(label, showLegendKey, showValue, showCategoryName, showSeriesName,
                        showPercent, showBubbleSize, position, numberFormat, sourceLinked);
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
            C.DataLabelPositionValues? position = null, string? numberFormat = null, bool sourceLinked = false,
            bool ignoreCase = true) {
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
                    showBubbleSize, position, numberFormat, sourceLinked);
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
            double? fontSizePoints = null, bool? bold = null, bool? italic = null,
            string? color = null, string? fontName = null) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            if (pointIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(pointIndex));
            }
            ValidateDataLabelTextStyle(fontSizePoints, color, fontName);

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
        /// Sets data label text style for a specific point by series name.
        /// </summary>
        public ExcelChart SetSeriesDataLabelTextStyleForPoint(string seriesName, int pointIndex,
            double? fontSizePoints = null, bool? bold = null, bool? italic = null,
            string? color = null, string? fontName = null, bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            if (pointIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(pointIndex));
            }
            ValidateDataLabelTextStyle(fontSizePoints, color, fontName);

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
        /// Sets data label shape styling for a specific point by series index.
        /// </summary>
        public ExcelChart SetSeriesDataLabelShapeStyleForPoint(int seriesIndex, int pointIndex,
            string? fillColor = null, string? lineColor = null,
            double? lineWidthPoints = null, bool noFill = false, bool noLine = false) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            if (pointIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(pointIndex));
            }
            ValidateDataLabelShapeStyle(fillColor, lineColor, lineWidthPoints, noFill, noLine);

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
        /// Sets data label shape styling for a specific point by series name.
        /// </summary>
        public ExcelChart SetSeriesDataLabelShapeStyleForPoint(string seriesName, int pointIndex,
            string? fillColor = null, string? lineColor = null,
            double? lineWidthPoints = null, bool noFill = false, bool noLine = false,
            bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            if (pointIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(pointIndex));
            }
            ValidateDataLabelShapeStyle(fillColor, lineColor, lineWidthPoints, noFill, noLine);

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
        /// Sets the separator for all data labels.
        /// </summary>
        public ExcelChart SetDataLabelSeparator(string? separator) {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return this;
            }

            foreach (C.BarChart barChart in plotArea.Elements<C.BarChart>()) {
                ApplyDataLabelSeparator(EnsureDataLabels(barChart), separator);
            }
            foreach (C.Bar3DChart barChart in plotArea.Elements<C.Bar3DChart>()) {
                ApplyDataLabelSeparator(EnsureDataLabels(barChart), separator);
            }
            foreach (C.LineChart lineChart in plotArea.Elements<C.LineChart>()) {
                ApplyDataLabelSeparator(EnsureDataLabels(lineChart), separator);
            }
            foreach (C.Line3DChart lineChart in plotArea.Elements<C.Line3DChart>()) {
                ApplyDataLabelSeparator(EnsureDataLabels(lineChart), separator);
            }
            foreach (C.AreaChart areaChart in plotArea.Elements<C.AreaChart>()) {
                ApplyDataLabelSeparator(EnsureDataLabels(areaChart), separator);
            }
            foreach (C.Area3DChart areaChart in plotArea.Elements<C.Area3DChart>()) {
                ApplyDataLabelSeparator(EnsureDataLabels(areaChart), separator);
            }
            foreach (C.PieChart pieChart in plotArea.Elements<C.PieChart>()) {
                ApplyDataLabelSeparator(EnsureDataLabels(pieChart), separator);
            }
            foreach (C.Pie3DChart pieChart in plotArea.Elements<C.Pie3DChart>()) {
                ApplyDataLabelSeparator(EnsureDataLabels(pieChart), separator);
            }
            foreach (C.OfPieChart pieChart in plotArea.Elements<C.OfPieChart>()) {
                ApplyDataLabelSeparator(EnsureDataLabels(pieChart), separator);
            }
            foreach (C.DoughnutChart doughnutChart in plotArea.Elements<C.DoughnutChart>()) {
                ApplyDataLabelSeparator(EnsureDataLabels(doughnutChart), separator);
            }
            foreach (C.ScatterChart scatterChart in plotArea.Elements<C.ScatterChart>()) {
                ApplyDataLabelSeparator(EnsureDataLabels(scatterChart), separator);
            }
            foreach (C.BubbleChart bubbleChart in plotArea.Elements<C.BubbleChart>()) {
                ApplyDataLabelSeparator(EnsureDataLabels(bubbleChart), separator);
            }
            foreach (C.RadarChart radarChart in plotArea.Elements<C.RadarChart>()) {
                ApplyDataLabelSeparator(EnsureDataLabels(radarChart), separator);
            }
            foreach (C.StockChart stockChart in plotArea.Elements<C.StockChart>()) {
                ApplyDataLabelSeparator(EnsureDataLabels(stockChart), separator);
            }

            Save();
            return this;
        }

        /// <summary>
        /// Sets the separator for data labels in a series by index.
        /// </summary>
        public ExcelChart SetSeriesDataLabelSeparator(int seriesIndex, string? separator) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
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
        /// Sets the separator for data labels in a series by name.
        /// </summary>
        public ExcelChart SetSeriesDataLabelSeparator(string seriesName, string? separator, bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
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
        /// Sets the separator for a specific data label point by series index.
        /// </summary>
        public ExcelChart SetSeriesDataLabelSeparatorForPoint(int seriesIndex, int pointIndex, string? separator) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            if (pointIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(pointIndex));
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
        /// Sets the separator for a specific data label point by series name.
        /// </summary>
        public ExcelChart SetSeriesDataLabelSeparatorForPoint(string seriesName, int pointIndex, string? separator,
            bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            if (pointIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(pointIndex));
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

        /// <summary>
        /// Configures data labels for a single series by index.
        /// </summary>
        public ExcelChart SetSeriesDataLabels(int seriesIndex, bool showValue = true, bool showCategoryName = false,
            bool showSeriesName = false, bool showLegendKey = false, bool showPercent = false,
            C.DataLabelPositionValues? position = null, string? numberFormat = null, bool sourceLinked = false) {
            if (seriesIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(seriesIndex));
            }
            if (numberFormat != null && string.IsNullOrWhiteSpace(numberFormat)) {
                throw new ArgumentException("Number format cannot be empty.", nameof(numberFormat));
            }

            bool applied = ApplySeriesByIndex(seriesIndex, series => {
                ApplyDataLabels(series, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent,
                    position, numberFormat, sourceLinked);
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
            C.DataLabelPositionValues? position = null, string? numberFormat = null, bool sourceLinked = false,
            bool ignoreCase = true) {
            if (seriesName == null) {
                throw new ArgumentNullException(nameof(seriesName));
            }
            if (numberFormat != null && string.IsNullOrWhiteSpace(numberFormat)) {
                throw new ArgumentException("Number format cannot be empty.", nameof(numberFormat));
            }

            bool applied = ApplySeriesByName(seriesName, ignoreCase, series => {
                ApplyDataLabels(series, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent,
                    position, numberFormat, sourceLinked);
            });

            if (!applied) {
                throw new InvalidOperationException($"Series '{seriesName}' was not found.");
            }

            Save();
            return this;
        }

    }
}
