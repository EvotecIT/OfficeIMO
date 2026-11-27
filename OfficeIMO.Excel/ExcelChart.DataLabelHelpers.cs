using System.Globalization;
using OfficeIMO.OpenXml.Internal;
using static OfficeIMO.Excel.OpenXml.ExcelChartXml;

namespace OfficeIMO.Excel {
    // Represents a chart on a worksheet.
    public sealed partial class ExcelChart {
        // Elements that follow c:dLbls in a series, in the order the insertion point is looked up.
        private static readonly string[] SeriesChildrenAfterDataLabels = {
            "trendline", "errBars", "cat", "val", "xVal", "yVal", "bubbleSize", "smooth", "extLst"
        };

        // Elements that follow c:dLbls in a chart group, in the order the insertion point is looked up.
        private static readonly string[] ChartChildrenAfterDataLabels = {
            "gapWidth", "gapDepth", "overlap", "bubbleScale", "showNegBubbles", "sizeRepresents", "hiLowLines",
            "upDownBars", "shape", "axId", "extLst"
        };

        // The order c:dLbls children are written in; c:dLbl overrides come first, all of them.
        private static readonly string[] DataLabelsChildOrder = {
            "dLbl", "delete", "numFmt", "spPr", "txPr", "dLblPos", "showLegendKey", "showVal", "showCatName",
            "showSerName", "showPercent", "showBubbleSize", "separator", "showLeaderLines", "leaderLines", "extLst"
        };

        // The order c:dLbl children are written in, from c:idx.
        private static readonly string[] DataLabelChildOrder = {
            "idx", "delete", "layout", "tx", "numFmt", "spPr", "txPr", "dLblPos", "showLegendKey", "showVal",
            "showCatName", "showSerName", "showPercent", "showBubbleSize", "separator", "showLeaderLines",
            "leaderLines", "extLst"
        };

        private static void ApplyDataLabels(OfficeXmlNode chartElement, bool showLegendKey, bool showValue,
            bool showCategoryName, bool showSeriesName, bool showPercent) {
            ApplyDataLabels(chartElement, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent,
                null, null, null);
        }

        private static void ApplyDataLabels(OfficeXmlNode chartElement, bool? showLegendKey, bool? showValue,
            bool? showCategoryName, bool? showSeriesName, bool? showPercent,
            OfficeChartDataLabelPosition? position, string? numberFormat, bool? sourceLinked) {
            OfficeXmlNode labels = EnsureDataLabels(chartElement);
            ApplyDataLabelOverrides(labels, showLegendKey, showValue, showCategoryName, showSeriesName, showPercent,
                showBubbleSize: null, position, numberFormat, sourceLinked);
        }

        // DrawingML text is stored in hundredths of a point. Validate before attaching XML.
        private static int ToNativeFontSize(OfficeLength size) {
            int value = checked((int)decimal.Round(size.ToEmus() / 127m, 0, MidpointRounding.AwayFromZero));
            if (value <= 0) throw new ArgumentOutOfRangeException(nameof(size), "Font size must remain positive at DrawingML precision.");
            return value;
        }

        private static void ApplyTextStyle(OfficeXmlNode runProps, OfficeLength? fontSize, bool? bold,
            bool? italic, OfficeColor? color, string? fontFamily) {
            if (fontSize != null) {
                runProps.SetAttribute("sz", ToNativeFontSize(fontSize.Value).ToString(CultureInfo.InvariantCulture));
            }
            if (bold != null) {
                runProps.Attr("b", bold.Value);
            }
            if (italic != null) {
                runProps.Attr("i", italic.Value);
            }
            if (color != null) {
                ApplyTextSolidFill(runProps, color);
            }
            if (fontFamily != null) {
                runProps.RemoveChildren(A, "latin");
                runProps.Append(runProps.New(A, "latin").Attr("typeface", fontFamily));
            }
        }

        private static void ValidateDataLabelTextStyle(OfficeLength? fontSize, OfficeColor? color, string? fontFamily) {
            if (fontSize != null && fontSize.Value <= OfficeLength.Zero) {
                throw new ArgumentOutOfRangeException(nameof(fontSize));
            }
            if (fontSize.HasValue) _ = ToNativeFontSize(fontSize.Value);
            if (fontFamily != null && string.IsNullOrWhiteSpace(fontFamily)) {
                throw new ArgumentException("Font name cannot be empty.", nameof(fontFamily));
            }
        }

        private static void ValidateDataLabelShapeStyle(OfficeColor? fillColor, OfficeColor? lineColor, OfficeLength? lineWidth,
            bool noFill, bool noLine) {
            OfficeChartStyleBounds.ValidateLineWidth(lineWidth?.ToPoints(), nameof(lineWidth));
            if (noFill && fillColor != null) {
                throw new ArgumentException("Cannot set both fill color and noFill.", nameof(noFill));
            }
            if (noLine && (lineColor != null || lineWidth != null)) {
                throw new ArgumentException("Cannot set line color/width when noLine is true.", nameof(noLine));
            }
        }

        private static void ValidateDataLabelLeaderLines(OfficeColor? lineColor, OfficeLength? lineWidth) {
            OfficeChartStyleBounds.ValidateLineWidth(lineWidth?.ToPoints(), nameof(lineWidth));
        }

        private static void ValidateAxisGridlinesStyle(OfficeColor? lineColor, OfficeLength? lineWidth) {
            OfficeChartStyleBounds.ValidateLineWidth(lineWidth?.ToPoints(), nameof(lineWidth));
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

        private static void ValidateCrossesAtForAxis(OfficeXmlNode axis, double? crossesAt) {
            if (crossesAt == null) {
                return;
            }

            OfficeXmlNode scaling = axis.Child(C, "scaling");
            if (scaling.Child(C, "logBase").Exists && crossesAt.Value <= 0) {
                throw new ArgumentOutOfRangeException(nameof(crossesAt), "Crosses-at value must be greater than 0 for log scale.");
            }
        }

        private static void ApplyAxisScale(OfficeXmlNode axis, double? minimum, double? maximum,
            double? majorUnit, double? minorUnit, bool? reverseOrder, bool? logScale, double? logBase) {
            if (reverseOrder != null || minimum != null || maximum != null || logScale != null || logBase != null) {
                OfficeXmlNode scaling = EnsureScaling(axis);
                if (reverseOrder != null) {
                    ReplaceChild(scaling, scaling.NewVal(C, "orientation", reverseOrder.Value ? "maxMin" : "minMax"));
                }
                if (minimum != null) {
                    ReplaceChild(scaling, scaling.NewVal(C, "min", minimum.Value));
                }
                if (maximum != null) {
                    ReplaceChild(scaling, scaling.NewVal(C, "max", maximum.Value));
                }

                bool effectiveLog = logScale == true || logBase != null;
                if (effectiveLog) {
                    double baseValue = logBase ?? 10d;
                    ReplaceChild(scaling, scaling.NewVal(C, "logBase", baseValue));
                } else if (logScale == false) {
                    scaling.RemoveChild(C, "logBase");
                }

                NormalizeScalingOrder(scaling);
            }

            if (majorUnit != null) {
                ReplaceChild(axis, axis.NewVal(C, "majorUnit", majorUnit.Value));
            }
            if (minorUnit != null) {
                ReplaceChild(axis, axis.NewVal(C, "minorUnit", minorUnit.Value));
            }
        }

        private static OfficeXmlNode EnsureScaling(OfficeXmlNode axis) {
            OfficeXmlNode scaling = axis.Child(C, "scaling");
            if (scaling.Exists) {
                return scaling;
            }

            scaling = axis.New(C, "scaling");
            OfficeXmlNode axisId = axis.Child(C, "axId");
            return axisId.Exists ? axisId.InsertAfterSelf(scaling) : axis.InsertBefore(scaling, axis.FirstChild);
        }

        private static readonly string[] ScalingChildOrder = { "logBase", "orientation", "max", "min" };

        private static void NormalizeScalingOrder(OfficeXmlNode scaling) {
            ReorderChildren(scaling, ScalingChildOrder, multipleIndex: -1);
        }

        /// <summary>
        /// Puts the chart children named in <paramref name="order"/> in that order, the first of each name only (all of
        /// them for the name at <paramref name="multipleIndex"/>), followed by every other child in document order.
        /// Later duplicates of a single-valued name are dropped, as rebuilding from typed child lookups drops them.
        /// </summary>
        private static void ReorderChildren(OfficeXmlNode parent, string[] order, int multipleIndex) {
            var captured = new OfficeXmlNode[order.Length];
            List<OfficeXmlNode>? multiple = null;
            List<OfficeXmlNode>? others = null;
            foreach (OfficeXmlNode child in parent.Children()) {
                int index = child.NamespaceId == C ? Array.IndexOf(order, child.LocalName) : -1;
                if (index < 0) {
                    (others ??= new List<OfficeXmlNode>()).Add(child);
                } else if (index == multipleIndex) {
                    (multiple ??= new List<OfficeXmlNode>()).Add(child);
                } else if (captured[index].IsNone) {
                    captured[index] = child;
                }
            }

            parent.RemoveAllChildren();

            for (int i = 0; i < order.Length; i++) {
                if (i == multipleIndex) {
                    if (multiple != null) {
                        foreach (OfficeXmlNode child in multiple) parent.Append(child);
                    }
                } else if (captured[i].Exists) {
                    parent.Append(captured[i]);
                }
            }
            if (others != null) {
                foreach (OfficeXmlNode child in others) parent.Append(child);
            }
        }

        private static void ValidateTrendline(OfficeChartTrendlineType type, int? order, int? period,
            double? forward, double? backward, OfficeColor? lineColor, OfficeLength? lineWidth) {
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
            OfficeChartStyleBounds.ValidateLineWidth(lineWidth?.ToPoints(), nameof(lineWidth));

            bool isPolynomial = type == OfficeChartTrendlineType.Polynomial;
            bool isMovingAverage = type == OfficeChartTrendlineType.MovingAverage;
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

        private static OfficeXmlNode EnsureDataLabels(OfficeXmlNode chartElement) {
            OfficeXmlNode labels = chartElement.Child(C, "dLbls");
            if (labels.Exists) {
                return labels;
            }

            labels = NewDataLabels(chartElement);
            OfficeXmlNode insertBefore = IsSeriesElement(chartElement)
                ? FirstExistingChild(chartElement, SeriesChildrenAfterDataLabels)
                : FirstExistingChild(chartElement, ChartChildrenAfterDataLabels);
            return chartElement.InsertBefore(labels, insertBefore);
        }

        private static void NormalizeDataLabelContainerOrder(OfficeXmlNode labels) {
            if (labels.Is(C, "dLbls")) {
                NormalizeDataLabelsOrder(labels);
            } else if (labels.Is(C, "dLbl")) {
                NormalizeDataLabelOrder(labels);
            }
        }

        private static void ApplyDataLabelTextStyle(OfficeXmlNode labels, OfficeLength? fontSize, bool? bold,
            bool? italic, OfficeColor? color, string? fontFamily) {
            ApplyTextStyle(EnsureTextPropertiesRunProperties(labels), fontSize, bold, italic, color, fontFamily);
            NormalizeDataLabelContainerOrder(labels);
        }

        private static void ApplyDataLabelShapeStyle(OfficeXmlNode labels, OfficeColor? fillColor, OfficeColor? lineColor,
            OfficeLength? lineWidth, bool noFill, bool noLine) {
            OfficeXmlNode props = EnsureDataLabelShapeProperties(labels);
            if (noFill) {
                ApplyNoFill(props);
            } else if (fillColor != null) {
                ApplySolidFill(props, fillColor);
            }

            if (noLine) {
                ApplyNoLine(props);
            } else if (lineColor != null || lineWidth != null) {
                OfficeColor? normalizedLine = lineColor != null ? lineColor : null;
                ApplyOptionalLine(props, normalizedLine, lineWidth);
            }

            NormalizeDataLabelContainerOrder(labels);
        }

        private static void ApplyAreaStyle(OfficeXmlNode props, OfficeColor? fillColor, OfficeColor? lineColor,
            OfficeLength? lineWidth, bool noFill, bool noLine) {
            if (noFill) {
                ApplyNoFill(props);
            } else if (fillColor != null) {
                ApplySolidFill(props, fillColor);
            }

            if (noLine) {
                ApplyNoLine(props);
            } else if (lineColor != null || lineWidth != null) {
                OfficeColor? normalizedLine = lineColor != null ? lineColor : null;
                ApplyOptionalLine(props, normalizedLine, lineWidth);
            }
        }

        private static void ApplyDataLabelLeaderLines(OfficeXmlNode labels, bool showLeaderLines, OfficeColor? lineColor,
            OfficeLength? lineWidth) {
            ReplaceChild(labels, labels.NewVal(C, "showLeaderLines", showLeaderLines));

            if (lineColor != null || lineWidth != null) {
                OfficeXmlNode leaderLines = labels.Child(C, "leaderLines");
                if (leaderLines.IsNone) leaderLines = labels.Append(labels.New(C, "leaderLines"));
                OfficeXmlNode props = leaderLines.Child(C, "spPr");
                if (props.IsNone) props = leaderLines.Append(leaderLines.New(C, "spPr"));
                OfficeColor? normalizedLine = lineColor != null ? lineColor : null;
                ApplyOptionalLine(props, normalizedLine, lineWidth);
            }

            NormalizeDataLabelsOrder(labels);
        }

        private static void ApplyDataLabelOverrides(OfficeXmlNode label, bool? showLegendKey, bool? showValue,
            bool? showCategoryName, bool? showSeriesName, bool? showPercent, bool? showBubbleSize,
            OfficeChartDataLabelPosition? position, string? numberFormat, bool? sourceLinked) {
            if (showLegendKey != null) {
                ReplaceChild(label, label.NewVal(C, "showLegendKey", showLegendKey.Value));
            }
            if (showValue != null) {
                ReplaceChild(label, label.NewVal(C, "showVal", showValue.Value));
            }
            if (showCategoryName != null) {
                ReplaceChild(label, label.NewVal(C, "showCatName", showCategoryName.Value));
            }
            if (showSeriesName != null) {
                ReplaceChild(label, label.NewVal(C, "showSerName", showSeriesName.Value));
            }
            if (showPercent != null) {
                ReplaceChild(label, label.NewVal(C, "showPercent", showPercent.Value));
            }
            if (showBubbleSize != null) {
                ReplaceChild(label, label.NewVal(C, "showBubbleSize", showBubbleSize.Value));
            }
            if (position != null) {
                ApplyDataLabelPosition(label, label, position.Value);
            }
            if (numberFormat != null || sourceLinked.HasValue) {
                OfficeXmlNode existing = label.Child(C, "numFmt");
                OfficeXmlNode format = existing.Exists
                    ? existing
                    : label.New(C, "numFmt").Attr("formatCode", "General").Attr("sourceLinked", false);
                if (numberFormat != null) format.SetAttribute("formatCode", numberFormat);
                if (sourceLinked.HasValue) format.Attr("sourceLinked", sourceLinked.Value);
                if (existing.IsNone) {
                    // c:delete and the formatting children are the two branches of one choice; adding a format leaves
                    // the formatting branch, as the SDK's AddChild did.
                    label.RemoveChildren(C, "delete");
                    label.AddOrdered(format, label.Is(C, "dLbls") ? CDLblsView.Order : CDLblView.Order);
                }
            }

            NormalizeDataLabelContainerOrder(label);
        }

        private static void ApplyDataLabelSeparator(OfficeXmlNode label, string? separator) {
            OfficeXmlNode existing = label.Child(C, "separator");
            if (separator == null) {
                if (existing.Exists) existing.Remove();
                return;
            }

            if (existing.Exists) existing.Remove();
            label.Append(label.NewText(C, "separator", separator));

            NormalizeDataLabelContainerOrder(label);
        }

        private static void ApplyDataLabelPosition(OfficeXmlNode labelContainer, OfficeXmlNode context, OfficeChartDataLabelPosition position) {
            if (TryNormalizeDataLabelPosition(context, position, out var normalizedPosition)) {
                ReplaceChild(labelContainer, labelContainer.NewVal(C, "dLblPos", normalizedPosition.ToOpenXmlToken()));
                return;
            }

            labelContainer.RemoveChildren(C, "dLblPos");
        }

        private static bool TryNormalizeDataLabelPosition(OfficeXmlNode context, OfficeChartDataLabelPosition position, out OfficeChartDataLabelPosition normalizedPosition) {
            normalizedPosition = position;

            for (OfficeXmlNode current = context; current.Exists; current = current.Parent) {
                if (current.Is(C, "doughnutChart")) {
                    return position != OfficeChartDataLabelPosition.BestFit;
                }

                if (position == OfficeChartDataLabelPosition.OutsideEnd
                    && (current.Is(C, "lineChart") || current.Is(C, "line3DChart") || IsLineSeries(current))) {
                    normalizedPosition = OfficeChartDataLabelPosition.Top;
                    return true;
                }

                if (position == OfficeChartDataLabelPosition.OutsideEnd && current.Is(C, "barChart")) {
                    string? grouping = current.ChildVal(C, "grouping");
                    if (grouping == "stacked" || grouping == "percentStacked") {
                        normalizedPosition = OfficeChartDataLabelPosition.Center;
                        return true;
                    }
                }
            }

            return true;
        }

        private static void ApplyDataLabelTemplate(OfficeXmlNode series, ExcelChartDataLabelTemplate template) {
            if (template.NumberFormat != null && string.IsNullOrWhiteSpace(template.NumberFormat)) {
                throw new ArgumentException("Number format cannot be empty.", nameof(template.NumberFormat));
            }

            bool applyTextStyle = template.FontSize != null
                || template.Bold != null
                || template.Italic != null
                || template.TextColor != null
                || template.FontFamily != null;
            bool applyShapeStyle = template.NoFill
                || template.NoLine
                || template.FillColor != null
                || template.LineColor != null
                || template.LineWidth != null;
            bool applyLeaderLines = template.ShowLeaderLines != null
                || template.LeaderLineColor != null
                || template.LeaderLineWidth != null;

            if (applyTextStyle) {
                ValidateDataLabelTextStyle(template.FontSize, template.TextColor, template.FontFamily);
            }
            if (applyShapeStyle) {
                ValidateDataLabelShapeStyle(template.FillColor, template.LineColor, template.LineWidth,
                    template.NoFill, template.NoLine);
            }
            if (applyLeaderLines) {
                ValidateDataLabelLeaderLines(template.LeaderLineColor, template.LeaderLineWidth);
            }

            OfficeXmlNode labels = EnsureDataLabels(series);
            ApplyDataLabelOverrides(labels, template.ShowLegendKey, template.ShowValue, template.ShowCategoryName,
                template.ShowSeriesName, template.ShowPercent, template.ShowBubbleSize,
                template.Position, template.NumberFormat, template.NumberFormat != null ? template.SourceLinked : (bool?)null);
            if (template.Separator != null) {
                ApplyDataLabelSeparator(labels, template.Separator);
            }

            if (applyTextStyle) {
                ApplyDataLabelTextStyle(labels, template.FontSize, template.Bold, template.Italic,
                    template.TextColor, template.FontFamily);
            }
            if (applyShapeStyle) {
                ApplyDataLabelShapeStyle(labels, template.FillColor, template.LineColor, template.LineWidth,
                    template.NoFill, template.NoLine);
            }
            if (applyLeaderLines) {
                bool showLeaderLines = template.ShowLeaderLines ?? true;
                ApplyDataLabelLeaderLines(labels, showLeaderLines, template.LeaderLineColor, template.LeaderLineWidth);
            }
        }

        private static OfficeXmlNode EnsureDataLabel(OfficeXmlNode series, int pointIndex) {
            OfficeXmlNode labels = EnsureDataLabels(series);
            uint idx = (uint)pointIndex;
            foreach (OfficeXmlNode item in labels.Children(C, "dLbl")) {
                if (item.Child(C, "idx").UInt32Val() == idx) {
                    return item;
                }
            }

            return InsertDataLabel(labels, idx);
        }

        // A new c:dLbl override with its c:idx, after the existing overrides.
        private static OfficeXmlNode InsertDataLabel(OfficeXmlNode labels, uint idx) {
            OfficeXmlNode label = labels.New(C, "dLbl");
            label.Append(label.NewVal(C, "idx", idx));

            OfficeXmlNode insertBefore = OfficeXmlNode.None;
            foreach (OfficeXmlNode child in labels.Children()) {
                if (!child.Is(C, "dLbl")) {
                    insertBefore = child;
                    break;
                }
            }
            return labels.InsertBefore(label, insertBefore);
        }

        private static void NormalizeDataLabelsOrder(OfficeXmlNode labels) {
            ReorderChildren(labels, DataLabelsChildOrder, multipleIndex: 0);
        }

        private static void NormalizeDataLabelOrder(OfficeXmlNode label) {
            if (label.Child(C, "idx").IsNone) {
                return;
            }

            ReorderChildren(label, DataLabelChildOrder, multipleIndex: -1);
        }

        private static OfficeXmlNode EnsureSeriesChildPosition(OfficeXmlNode series, OfficeXmlNode child, OfficeXmlNode insertBefore) {
            if (child.Parent.Exists) {
                child.Remove();
            }

            return series.InsertBefore(child, insertBefore);
        }

    }
}
