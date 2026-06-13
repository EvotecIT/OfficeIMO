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
        private static void ApplyDataLabels(OpenXmlCompositeElement chartElement, bool showLegendKey, bool showValue,
            bool showCategoryName, bool showSeriesName, bool showPercent) {
            ApplyDataLabels(chartElement, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent,
                null, null, false);
        }

        private static void ApplyDataLabels(OpenXmlCompositeElement chartElement, bool showLegendKey, bool showValue,
            bool showCategoryName, bool showSeriesName, bool showPercent,
            C.DataLabelPositionValues? position, string? numberFormat, bool sourceLinked) {
            C.DataLabels labels = EnsureDataLabels(chartElement);
            ReplaceChild(labels, new C.ShowLegendKey { Val = showLegendKey });
            ReplaceChild(labels, new C.ShowValue { Val = showValue });
            ReplaceChild(labels, new C.ShowCategoryName { Val = showCategoryName });
            ReplaceChild(labels, new C.ShowSeriesName { Val = showSeriesName });
            ReplaceChild(labels, new C.ShowPercent { Val = showPercent });
            ReplaceChild(labels, new C.ShowBubbleSize { Val = false });

            if (position != null) {
                ApplyDataLabelPosition(labels, chartElement, position.Value);
            }

            if (numberFormat != null) {
                ReplaceChild(labels, new C.NumberingFormat {
                    FormatCode = numberFormat,
                    SourceLinked = sourceLinked
                });
            }

            NormalizeDataLabelsOrder(labels);
        }

        private static void ApplyTextStyle(A.TextCharacterPropertiesType runProps, double? fontSizePoints, bool? bold,
            bool? italic, string? color, string? fontName) {
            if (fontSizePoints != null) {
                runProps.FontSize = (int)Math.Round(fontSizePoints.Value * 100);
            }
            if (bold != null) {
                runProps.Bold = bold.Value;
            }
            if (italic != null) {
                runProps.Italic = italic.Value;
            }
            if (color != null) {
                ApplyTextSolidFill(runProps, NormalizeHexColor(color));
            }
            if (fontName != null) {
                runProps.RemoveAllChildren<A.LatinFont>();
                runProps.Append(new A.LatinFont { Typeface = fontName });
            }
        }

        private static void ValidateDataLabelTextStyle(double? fontSizePoints, string? color, string? fontName) {
            if (fontSizePoints != null && fontSizePoints <= 0) {
                throw new ArgumentOutOfRangeException(nameof(fontSizePoints));
            }
            if (color != null && string.IsNullOrWhiteSpace(color)) {
                throw new ArgumentException("Label color cannot be empty.", nameof(color));
            }
            if (fontName != null && string.IsNullOrWhiteSpace(fontName)) {
                throw new ArgumentException("Font name cannot be empty.", nameof(fontName));
            }
        }

        private static void ValidateDataLabelShapeStyle(string? fillColor, string? lineColor, double? lineWidthPoints,
            bool noFill, bool noLine) {
            if (fillColor != null && string.IsNullOrWhiteSpace(fillColor)) {
                throw new ArgumentException("Fill color cannot be empty.", nameof(fillColor));
            }
            if (lineColor != null && string.IsNullOrWhiteSpace(lineColor)) {
                throw new ArgumentException("Line color cannot be empty.", nameof(lineColor));
            }
            if (lineWidthPoints != null && lineWidthPoints <= 0) {
                throw new ArgumentOutOfRangeException(nameof(lineWidthPoints));
            }
            if (noFill && fillColor != null) {
                throw new ArgumentException("Cannot set both fill color and noFill.", nameof(noFill));
            }
            if (noLine && (lineColor != null || lineWidthPoints != null)) {
                throw new ArgumentException("Cannot set line color/width when noLine is true.", nameof(noLine));
            }
        }

        private static void ValidateDataLabelLeaderLines(string? lineColor, double? lineWidthPoints) {
            if (lineColor != null && string.IsNullOrWhiteSpace(lineColor)) {
                throw new ArgumentException("Line color cannot be empty.", nameof(lineColor));
            }
            if (lineWidthPoints != null && lineWidthPoints <= 0) {
                throw new ArgumentOutOfRangeException(nameof(lineWidthPoints));
            }
        }

        private static void ValidateAxisGridlinesStyle(string? lineColor, double? lineWidthPoints) {
            if (lineColor != null && string.IsNullOrWhiteSpace(lineColor)) {
                throw new ArgumentException("Gridline color cannot be empty.", nameof(lineColor));
            }
            if (lineWidthPoints != null && lineWidthPoints <= 0) {
                throw new ArgumentOutOfRangeException(nameof(lineWidthPoints));
            }
        }

        private static void ValidateAxisScale(double? minimum, double? maximum, double? majorUnit, double? minorUnit,
            bool? logScale, double? logBase) {
            if (minimum != null && double.IsNaN(minimum.Value)) {
                throw new ArgumentOutOfRangeException(nameof(minimum));
            }
            if (maximum != null && double.IsNaN(maximum.Value)) {
                throw new ArgumentOutOfRangeException(nameof(maximum));
            }
            if (minimum != null && maximum != null && minimum.Value >= maximum.Value) {
                throw new ArgumentException("Minimum must be less than maximum.");
            }
            if (majorUnit != null && majorUnit.Value <= 0) {
                throw new ArgumentOutOfRangeException(nameof(majorUnit));
            }
            if (minorUnit != null && minorUnit.Value <= 0) {
                throw new ArgumentOutOfRangeException(nameof(minorUnit));
            }
            if (logScale == false && logBase != null) {
                throw new ArgumentException("Log base requires logScale to be enabled.", nameof(logBase));
            }

            bool effectiveLog = logScale == true || logBase != null;
            if (effectiveLog) {
                double baseValue = logBase ?? 10d;
                if (baseValue <= 1d) {
                    throw new ArgumentOutOfRangeException(nameof(logBase), "Log base must be greater than 1.");
                }
                if (minimum != null && minimum.Value <= 0) {
                    throw new ArgumentException("Minimum must be greater than 0 for log scale.", nameof(minimum));
                }
                if (maximum != null && maximum.Value <= 0) {
                    throw new ArgumentException("Maximum must be greater than 0 for log scale.", nameof(maximum));
                }
            }
        }

        private static void ValidateCrossesAtForAxis(OpenXmlCompositeElement axis, double? crossesAt) {
            if (crossesAt == null) {
                return;
            }

            C.Scaling? scaling = axis.GetFirstChild<C.Scaling>();
            if (scaling?.GetFirstChild<C.LogBase>() != null && crossesAt.Value <= 0) {
                throw new ArgumentOutOfRangeException(nameof(crossesAt), "Crosses-at value must be greater than 0 for log scale.");
            }
        }

        private static void ApplyAxisScale(OpenXmlCompositeElement axis, double? minimum, double? maximum,
            double? majorUnit, double? minorUnit, bool? reverseOrder, bool? logScale, double? logBase) {
            if (reverseOrder != null || minimum != null || maximum != null || logScale != null || logBase != null) {
                C.Scaling scaling = EnsureScaling(axis);
                if (reverseOrder != null) {
                    ReplaceChild(scaling, new C.Orientation {
                        Val = reverseOrder.Value ? C.OrientationValues.MaxMin : C.OrientationValues.MinMax
                    });
                }
                if (minimum != null) {
                    ReplaceChild(scaling, new C.MinAxisValue { Val = minimum.Value });
                }
                if (maximum != null) {
                    ReplaceChild(scaling, new C.MaxAxisValue { Val = maximum.Value });
                }

                bool effectiveLog = logScale == true || logBase != null;
                if (effectiveLog) {
                    double baseValue = logBase ?? 10d;
                    ReplaceChild(scaling, new C.LogBase { Val = baseValue });
                } else if (logScale == false) {
                    scaling.GetFirstChild<C.LogBase>()?.Remove();
                }

                NormalizeScalingOrder(scaling);
            }

            if (majorUnit != null) {
                ReplaceChild(axis, new C.MajorUnit { Val = majorUnit.Value });
            }
            if (minorUnit != null) {
                ReplaceChild(axis, new C.MinorUnit { Val = minorUnit.Value });
            }
        }

        private static C.Scaling EnsureScaling(OpenXmlCompositeElement axis) {
            C.Scaling scaling = axis.GetFirstChild<C.Scaling>() ?? new C.Scaling();
            if (scaling.Parent == null) {
                C.AxisId? axisId = axis.GetFirstChild<C.AxisId>();
                if (axisId != null) {
                    axis.InsertAfter(scaling, axisId);
                } else {
                    axis.InsertAt(scaling, 0);
                }
            }
            return scaling;
        }

        private static void NormalizeScalingOrder(C.Scaling scaling) {
            C.LogBase? logBase = scaling.GetFirstChild<C.LogBase>();
            C.Orientation? orientation = scaling.GetFirstChild<C.Orientation>();
            C.MaxAxisValue? max = scaling.GetFirstChild<C.MaxAxisValue>();
            C.MinAxisValue? min = scaling.GetFirstChild<C.MinAxisValue>();

            List<OpenXmlElement> otherChildren = scaling.ChildElements
                .Where(child => child is not C.LogBase
                                && child is not C.Orientation
                                && child is not C.MaxAxisValue
                                && child is not C.MinAxisValue)
                .ToList();

            scaling.RemoveAllChildren();

            if (logBase != null) {
                scaling.Append(logBase);
            }
            if (orientation != null) {
                scaling.Append(orientation);
            }
            if (max != null) {
                scaling.Append(max);
            }
            if (min != null) {
                scaling.Append(min);
            }
            foreach (OpenXmlElement child in otherChildren) {
                scaling.Append(child);
            }
        }

        private static void ValidateTrendline(C.TrendlineValues type, int? order, int? period,
            double? forward, double? backward, string? lineColor, double? lineWidthPoints) {
            if (order != null && (order <= 0 || order > byte.MaxValue)) {
                throw new ArgumentOutOfRangeException(nameof(order));
            }
            if (period != null && period <= 0) {
                throw new ArgumentOutOfRangeException(nameof(period));
            }
            if (forward != null && forward < 0) {
                throw new ArgumentOutOfRangeException(nameof(forward));
            }
            if (backward != null && backward < 0) {
                throw new ArgumentOutOfRangeException(nameof(backward));
            }
            if (lineColor != null && string.IsNullOrWhiteSpace(lineColor)) {
                throw new ArgumentException("Trendline color cannot be empty.", nameof(lineColor));
            }
            if (lineWidthPoints != null && lineWidthPoints <= 0) {
                throw new ArgumentOutOfRangeException(nameof(lineWidthPoints));
            }

            bool isPolynomial = type.Equals(C.TrendlineValues.Polynomial);
            bool isMovingAverage = type.Equals(C.TrendlineValues.MovingAverage);
            if (isPolynomial && order == null) {
                throw new ArgumentException("Polynomial trendlines require an order.", nameof(order));
            }
            if (!isPolynomial && order != null) {
                throw new ArgumentException("Order is only valid for polynomial trendlines.", nameof(order));
            }
            if (isMovingAverage && period == null) {
                throw new ArgumentException("Moving average trendlines require a period.", nameof(period));
            }
            if (!isMovingAverage && period != null) {
                throw new ArgumentException("Period is only valid for moving average trendlines.", nameof(period));
            }
        }

        private static C.DataLabels EnsureDataLabels(OpenXmlCompositeElement chartElement) {
            C.DataLabels labels = chartElement.GetFirstChild<C.DataLabels>() ?? new C.DataLabels();
            if (labels.Parent != null) {
                return labels;
            }

            OpenXmlElement? insertBefore;
            if (IsSeriesElement(chartElement)) {
                insertBefore = chartElement.GetFirstChild<C.Trendline>();
                insertBefore ??= chartElement.GetFirstChild<C.ErrorBars>();
                insertBefore ??= chartElement.GetFirstChild<C.CategoryAxisData>();
                insertBefore ??= chartElement.GetFirstChild<C.Values>();
                insertBefore ??= chartElement.GetFirstChild<C.XValues>();
                insertBefore ??= chartElement.GetFirstChild<C.YValues>();
                insertBefore ??= chartElement.GetFirstChild<C.BubbleSize>();
                insertBefore ??= chartElement.GetFirstChild<C.Smooth>();
                insertBefore ??= chartElement.GetFirstChild<C.ExtensionList>();
            } else {
                insertBefore = chartElement.GetFirstChild<C.GapWidth>();
                insertBefore ??= chartElement.GetFirstChild<C.GapDepth>();
                insertBefore ??= chartElement.GetFirstChild<C.Overlap>();
                insertBefore ??= chartElement.GetFirstChild<C.BubbleScale>();
                insertBefore ??= chartElement.GetFirstChild<C.ShowNegativeBubbles>();
                insertBefore ??= chartElement.GetFirstChild<C.SizeRepresents>();
                insertBefore ??= chartElement.GetFirstChild<C.HighLowLines>();
                insertBefore ??= chartElement.GetFirstChild<C.UpDownBars>();
                insertBefore ??= chartElement.GetFirstChild<C.Shape>();
                insertBefore ??= chartElement.GetFirstChild<C.AxisId>();
                insertBefore ??= chartElement.GetFirstChild<C.ExtensionList>();
            }

            if (insertBefore != null) {
                chartElement.InsertBefore(labels, insertBefore);
            } else {
                chartElement.Append(labels);
            }
            return labels;
        }

        private static void ApplyDataLabelTextStyle(OpenXmlCompositeElement labels, double? fontSizePoints, bool? bold,
            bool? italic, string? color, string? fontName) {
            ApplyTextStyle(EnsureTextPropertiesRunProperties(labels), fontSizePoints, bold, italic, color, fontName);

            if (labels is C.DataLabels dataLabels) {
                NormalizeDataLabelsOrder(dataLabels);
            } else if (labels is C.DataLabel dataLabel) {
                NormalizeDataLabelOrder(dataLabel);
            }
        }

        private static void ApplyDataLabelShapeStyle(OpenXmlCompositeElement labels, string? fillColor, string? lineColor,
            double? lineWidthPoints, bool noFill, bool noLine) {
            C.ChartShapeProperties props = EnsureDataLabelShapeProperties(labels);
            if (noFill) {
                ApplyNoFill(props);
            } else if (fillColor != null) {
                ApplySolidFill(props, NormalizeHexColor(fillColor));
            }

            if (noLine) {
                ApplyNoLine(props);
            } else if (lineColor != null || lineWidthPoints != null) {
                string? normalizedLine = lineColor != null ? NormalizeHexColor(lineColor) : null;
                ApplyOptionalLine(props, normalizedLine, lineWidthPoints);
            }

            if (labels is C.DataLabels dataLabels) {
                NormalizeDataLabelsOrder(dataLabels);
            } else if (labels is C.DataLabel dataLabel) {
                NormalizeDataLabelOrder(dataLabel);
            }
        }

        private static void ApplyAreaStyle(OpenXmlCompositeElement props, string? fillColor, string? lineColor,
            double? lineWidthPoints, bool noFill, bool noLine) {
            if (noFill) {
                ApplyNoFill(props);
            } else if (fillColor != null) {
                ApplySolidFill(props, NormalizeHexColor(fillColor));
            }

            if (noLine) {
                ApplyNoLine(props);
            } else if (lineColor != null || lineWidthPoints != null) {
                string? normalizedLine = lineColor != null ? NormalizeHexColor(lineColor) : null;
                ApplyOptionalLine(props, normalizedLine, lineWidthPoints);
            }
        }

        private static void ApplyDataLabelLeaderLines(C.DataLabels labels, bool showLeaderLines, string? lineColor,
            double? lineWidthPoints) {
            ReplaceChild(labels, new C.ShowLeaderLines { Val = showLeaderLines });

            if (lineColor != null || lineWidthPoints != null) {
                C.LeaderLines leaderLines = labels.GetFirstChild<C.LeaderLines>() ?? new C.LeaderLines();
                C.ChartShapeProperties props = leaderLines.GetFirstChild<C.ChartShapeProperties>() ?? new C.ChartShapeProperties();
                string? normalizedLine = lineColor != null ? NormalizeHexColor(lineColor) : null;
                ApplyOptionalLine(props, normalizedLine, lineWidthPoints);
                if (props.Parent == null) {
                    leaderLines.Append(props);
                }
                if (leaderLines.Parent == null) {
                    labels.Append(leaderLines);
                }
            }

            NormalizeDataLabelsOrder(labels);
        }

        private static void ApplyDataLabelOverrides(OpenXmlCompositeElement label, bool? showLegendKey, bool? showValue,
            bool? showCategoryName, bool? showSeriesName, bool? showPercent, bool? showBubbleSize,
            C.DataLabelPositionValues? position, string? numberFormat, bool sourceLinked) {
            if (showLegendKey != null) {
                ReplaceChild(label, new C.ShowLegendKey { Val = showLegendKey.Value });
            }
            if (showValue != null) {
                ReplaceChild(label, new C.ShowValue { Val = showValue.Value });
            }
            if (showCategoryName != null) {
                ReplaceChild(label, new C.ShowCategoryName { Val = showCategoryName.Value });
            }
            if (showSeriesName != null) {
                ReplaceChild(label, new C.ShowSeriesName { Val = showSeriesName.Value });
            }
            if (showPercent != null) {
                ReplaceChild(label, new C.ShowPercent { Val = showPercent.Value });
            }
            if (showBubbleSize != null) {
                ReplaceChild(label, new C.ShowBubbleSize { Val = showBubbleSize.Value });
            }
            if (position != null) {
                ApplyDataLabelPosition(label, label, position.Value);
            }
            if (numberFormat != null) {
                ReplaceChild(label, new C.NumberingFormat {
                    FormatCode = numberFormat,
                    SourceLinked = sourceLinked
                });
            }

            if (label is C.DataLabels dataLabels) {
                NormalizeDataLabelsOrder(dataLabels);
            } else if (label is C.DataLabel dataLabel) {
                NormalizeDataLabelOrder(dataLabel);
            }
        }

        private static void ApplyDataLabelSeparator(OpenXmlCompositeElement label, string? separator) {
            C.Separator? existing = label.GetFirstChild<C.Separator>();
            if (separator == null) {
                existing?.Remove();
                return;
            }

            existing?.Remove();
            label.Append(new C.Separator { Text = separator });

            if (label is C.DataLabels dataLabels) {
                NormalizeDataLabelsOrder(dataLabels);
            } else if (label is C.DataLabel dataLabel) {
                NormalizeDataLabelOrder(dataLabel);
            }
        }

        private static void ApplyDataLabelPosition(OpenXmlCompositeElement labelContainer, OpenXmlElement context, C.DataLabelPositionValues position) {
            if (TryNormalizeDataLabelPosition(context, position, out var normalizedPosition)) {
                ReplaceChild(labelContainer, new C.DataLabelPosition { Val = normalizedPosition });
                return;
            }

            labelContainer.RemoveAllChildren<C.DataLabelPosition>();
        }

        private static bool TryNormalizeDataLabelPosition(OpenXmlElement context, C.DataLabelPositionValues position, out C.DataLabelPositionValues normalizedPosition) {
            normalizedPosition = position;

            for (OpenXmlElement? current = context; current != null; current = current.Parent) {
                if (current is C.PieChart || current is C.Pie3DChart || current is C.OfPieChart || current is C.DoughnutChart
                    || current is C.PieChartSeries) {
                    return position != C.DataLabelPositionValues.BestFit;
                }

                if (position == C.DataLabelPositionValues.OutsideEnd
                    && (current is C.LineChart || current is C.Line3DChart || current is C.LineChartSeries)) {
                    normalizedPosition = C.DataLabelPositionValues.Top;
                    return true;
                }

                if (position == C.DataLabelPositionValues.OutsideEnd && current is C.BarChart barChart) {
                    var grouping = barChart.BarGrouping?.Val?.Value;
                    if (grouping == C.BarGroupingValues.Stacked || grouping == C.BarGroupingValues.PercentStacked) {
                        normalizedPosition = C.DataLabelPositionValues.Center;
                        return true;
                    }
                }
            }

            return true;
        }

        private static void ApplyDataLabelTemplate(OpenXmlCompositeElement series, ExcelChartDataLabelTemplate template) {
            if (template.NumberFormat != null && string.IsNullOrWhiteSpace(template.NumberFormat)) {
                throw new ArgumentException("Number format cannot be empty.", nameof(template.NumberFormat));
            }

            bool applyTextStyle = template.FontSizePoints != null
                || template.Bold != null
                || template.Italic != null
                || template.TextColor != null
                || template.FontName != null;
            bool applyShapeStyle = template.NoFill
                || template.NoLine
                || template.FillColor != null
                || template.LineColor != null
                || template.LineWidthPoints != null;
            bool applyLeaderLines = template.ShowLeaderLines != null
                || template.LeaderLineColor != null
                || template.LeaderLineWidthPoints != null;

            if (applyTextStyle) {
                ValidateDataLabelTextStyle(template.FontSizePoints, template.TextColor, template.FontName);
            }
            if (applyShapeStyle) {
                ValidateDataLabelShapeStyle(template.FillColor, template.LineColor, template.LineWidthPoints,
                    template.NoFill, template.NoLine);
            }
            if (applyLeaderLines) {
                ValidateDataLabelLeaderLines(template.LeaderLineColor, template.LeaderLineWidthPoints);
            }

            C.DataLabels labels = EnsureDataLabels(series);
            ApplyDataLabelOverrides(labels, template.ShowLegendKey, template.ShowValue, template.ShowCategoryName,
                template.ShowSeriesName, template.ShowPercent, template.ShowBubbleSize,
                template.Position, template.NumberFormat, template.SourceLinked);
            if (template.Separator != null) {
                ApplyDataLabelSeparator(labels, template.Separator);
            }

            if (applyTextStyle) {
                ApplyDataLabelTextStyle(labels, template.FontSizePoints, template.Bold, template.Italic,
                    template.TextColor, template.FontName);
            }
            if (applyShapeStyle) {
                ApplyDataLabelShapeStyle(labels, template.FillColor, template.LineColor, template.LineWidthPoints,
                    template.NoFill, template.NoLine);
            }
            if (applyLeaderLines) {
                bool showLeaderLines = template.ShowLeaderLines ?? true;
                ApplyDataLabelLeaderLines(labels, showLeaderLines, template.LeaderLineColor, template.LeaderLineWidthPoints);
            }
        }

        private static C.DataLabel EnsureDataLabel(OpenXmlCompositeElement series, int pointIndex) {
            C.DataLabels labels = EnsureDataLabels(series);
            uint idx = (uint)pointIndex;
            C.DataLabel? label = labels.Elements<C.DataLabel>()
                .FirstOrDefault(item => item.GetFirstChild<C.Index>()?.Val?.Value == idx);
            if (label == null) {
                label = new C.DataLabel();
                label.Append(new C.Index { Val = idx });

                OpenXmlElement? insertBefore = labels.ChildElements.FirstOrDefault(child => child is not C.DataLabel);
                if (insertBefore != null) {
                    labels.InsertBefore(label, insertBefore);
                } else {
                    labels.Append(label);
                }
            }
            return label;
        }

        private static void NormalizeDataLabelsOrder(C.DataLabels labels) {
            List<C.DataLabel> overrides = labels.Elements<C.DataLabel>().ToList();
            C.Delete? delete = labels.GetFirstChild<C.Delete>();
            C.NumberingFormat? numFmt = labels.GetFirstChild<C.NumberingFormat>();
            C.ChartShapeProperties? shapeProps = labels.GetFirstChild<C.ChartShapeProperties>();
            C.TextProperties? textProps = labels.GetFirstChild<C.TextProperties>();
            C.DataLabelPosition? position = labels.GetFirstChild<C.DataLabelPosition>();
            C.ShowLegendKey? showLegendKey = labels.GetFirstChild<C.ShowLegendKey>();
            C.ShowValue? showValue = labels.GetFirstChild<C.ShowValue>();
            C.ShowCategoryName? showCategoryName = labels.GetFirstChild<C.ShowCategoryName>();
            C.ShowSeriesName? showSeriesName = labels.GetFirstChild<C.ShowSeriesName>();
            C.ShowPercent? showPercent = labels.GetFirstChild<C.ShowPercent>();
            C.ShowBubbleSize? showBubbleSize = labels.GetFirstChild<C.ShowBubbleSize>();
            C.Separator? separator = labels.GetFirstChild<C.Separator>();
            C.ShowLeaderLines? showLeaderLines = labels.GetFirstChild<C.ShowLeaderLines>();
            C.LeaderLines? leaderLines = labels.GetFirstChild<C.LeaderLines>();
            C.ExtensionList? extLst = labels.GetFirstChild<C.ExtensionList>();

            List<OpenXmlElement> otherChildren = labels.ChildElements
                .Where(child => child is not C.DataLabel
                                && child is not C.Delete
                                && child is not C.NumberingFormat
                                && child is not C.ChartShapeProperties
                                && child is not C.TextProperties
                                && child is not C.DataLabelPosition
                                && child is not C.ShowLegendKey
                                && child is not C.ShowValue
                                && child is not C.ShowCategoryName
                                && child is not C.ShowSeriesName
                                && child is not C.ShowPercent
                                && child is not C.ShowBubbleSize
                                && child is not C.Separator
                                && child is not C.ShowLeaderLines
                                && child is not C.LeaderLines
                                && child is not C.ExtensionList)
                .ToList();

            labels.RemoveAllChildren();

            foreach (C.DataLabel child in overrides) {
                labels.Append(child);
            }
            if (delete != null) {
                labels.Append(delete);
            }
            if (numFmt != null) {
                labels.Append(numFmt);
            }
            if (shapeProps != null) {
                labels.Append(shapeProps);
            }
            if (textProps != null) {
                labels.Append(textProps);
            }
            if (position != null) {
                labels.Append(position);
            }
            if (showLegendKey != null) {
                labels.Append(showLegendKey);
            }
            if (showValue != null) {
                labels.Append(showValue);
            }
            if (showCategoryName != null) {
                labels.Append(showCategoryName);
            }
            if (showSeriesName != null) {
                labels.Append(showSeriesName);
            }
            if (showPercent != null) {
                labels.Append(showPercent);
            }
            if (showBubbleSize != null) {
                labels.Append(showBubbleSize);
            }
            if (separator != null) {
                labels.Append(separator);
            }
            if (showLeaderLines != null) {
                labels.Append(showLeaderLines);
            }
            if (leaderLines != null) {
                labels.Append(leaderLines);
            }
            if (extLst != null) {
                labels.Append(extLst);
            }
            foreach (OpenXmlElement child in otherChildren) {
                labels.Append(child);
            }
        }

        private static void NormalizeDataLabelOrder(C.DataLabel label) {
            C.Index? idx = label.GetFirstChild<C.Index>();
            if (idx == null) {
                return;
            }

            C.Delete? delete = label.GetFirstChild<C.Delete>();
            C.Layout? layout = label.GetFirstChild<C.Layout>();
            C.ChartText? chartText = label.GetFirstChild<C.ChartText>();
            C.NumberingFormat? numFmt = label.GetFirstChild<C.NumberingFormat>();
            C.ChartShapeProperties? shapeProps = label.GetFirstChild<C.ChartShapeProperties>();
            C.TextProperties? textProps = label.GetFirstChild<C.TextProperties>();
            C.DataLabelPosition? position = label.GetFirstChild<C.DataLabelPosition>();
            C.ShowLegendKey? showLegendKey = label.GetFirstChild<C.ShowLegendKey>();
            C.ShowValue? showValue = label.GetFirstChild<C.ShowValue>();
            C.ShowCategoryName? showCategoryName = label.GetFirstChild<C.ShowCategoryName>();
            C.ShowSeriesName? showSeriesName = label.GetFirstChild<C.ShowSeriesName>();
            C.ShowPercent? showPercent = label.GetFirstChild<C.ShowPercent>();
            C.ShowBubbleSize? showBubbleSize = label.GetFirstChild<C.ShowBubbleSize>();
            C.Separator? separator = label.GetFirstChild<C.Separator>();
            C.ShowLeaderLines? showLeaderLines = label.GetFirstChild<C.ShowLeaderLines>();
            C.LeaderLines? leaderLines = label.GetFirstChild<C.LeaderLines>();
            C.ExtensionList? extLst = label.GetFirstChild<C.ExtensionList>();

            List<OpenXmlElement> otherChildren = label.ChildElements
                .Where(child => child is not C.Index
                                && child is not C.Delete
                                && child is not C.Layout
                                && child is not C.ChartText
                                && child is not C.NumberingFormat
                                && child is not C.ChartShapeProperties
                                && child is not C.TextProperties
                                && child is not C.DataLabelPosition
                                && child is not C.ShowLegendKey
                                && child is not C.ShowValue
                                && child is not C.ShowCategoryName
                                && child is not C.ShowSeriesName
                                && child is not C.ShowPercent
                                && child is not C.ShowBubbleSize
                                && child is not C.Separator
                                && child is not C.ShowLeaderLines
                                && child is not C.LeaderLines
                                && child is not C.ExtensionList)
                .ToList();

            label.RemoveAllChildren();

            label.Append(idx);
            if (delete != null) {
                label.Append(delete);
            }
            if (layout != null) {
                label.Append(layout);
            }
            if (chartText != null) {
                label.Append(chartText);
            }
            if (numFmt != null) {
                label.Append(numFmt);
            }
            if (shapeProps != null) {
                label.Append(shapeProps);
            }
            if (textProps != null) {
                label.Append(textProps);
            }
            if (position != null) {
                label.Append(position);
            }
            if (showLegendKey != null) {
                label.Append(showLegendKey);
            }
            if (showValue != null) {
                label.Append(showValue);
            }
            if (showCategoryName != null) {
                label.Append(showCategoryName);
            }
            if (showSeriesName != null) {
                label.Append(showSeriesName);
            }
            if (showPercent != null) {
                label.Append(showPercent);
            }
            if (showBubbleSize != null) {
                label.Append(showBubbleSize);
            }
            if (separator != null) {
                label.Append(separator);
            }
            if (showLeaderLines != null) {
                label.Append(showLeaderLines);
            }
            if (leaderLines != null) {
                label.Append(leaderLines);
            }
            if (extLst != null) {
                label.Append(extLst);
            }
            foreach (OpenXmlElement child in otherChildren) {
                label.Append(child);
            }
        }

        private static void EnsureSeriesChildPosition(OpenXmlCompositeElement series, OpenXmlElement child, OpenXmlElement? insertBefore) {
            if (child.Parent != null) {
                child.Remove();
            }

            if (insertBefore != null) {
                series.InsertBefore(child, insertBefore);
            } else {
                series.Append(child);
            }
        }

    }
}
