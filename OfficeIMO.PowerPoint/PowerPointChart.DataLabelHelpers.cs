using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
        //     Represents a chart on a slide.
        public partial class PowerPointChart : PowerPointShape {
        // The single settings of c:dLbls and c:dLbl in the order they are written. Their c:extLst is a type of its own in
        // the SDK, so it stays among the other children rather than going last.
        private static readonly string[] DataLabelsSettings = {
            "delete", "numFmt", "spPr", "txPr", "dLblPos", "showLegendKey", "showVal", "showCatName", "showSerName", "showPercent",
            "showBubbleSize", "separator", "showLeaderLines", "leaderLines"
        };

        private static readonly string[] DataLabelSettings = {
            "delete", "layout", "tx", "numFmt", "spPr", "txPr", "dLblPos", "showLegendKey", "showVal", "showCatName", "showSerName",
            "showPercent", "showBubbleSize", "separator", "showLeaderLines", "leaderLines"
        };

        // The chart groups chart-wide data label settings apply to, in the order they are applied.
        private static readonly string[] DataLabelGroupNames = {
            "barChart", "lineChart", "areaChart", "pieChart", "doughnutChart", "scatterChart", "bubbleChart"
        };

        /// <summary>The chart groups of the plot area that take chart-wide data labels: every bar chart, then every line chart, and so on.</summary>
        private static IEnumerable<OfficeXmlNode> DataLabelGroups(OfficeXmlNode plotArea) {
            foreach (string groupName in DataLabelGroupNames) {
                foreach (OfficeXmlNode group in plotArea.Children(CNs, groupName).ToList()) yield return group;
            }
        }

        private static void ApplyDataLabels(OfficeXmlNode chartElement, bool showLegendKey, bool showValue,
            bool showCategoryName, bool showSeriesName, bool showPercent) {
            OfficeXmlNode labels = EnsureDataLabels(chartElement);
            ReplaceValChild(labels, "showLegendKey", showLegendKey);
            ReplaceValChild(labels, "showVal", showValue);
            ReplaceValChild(labels, "showCatName", showCategoryName);
            ReplaceValChild(labels, "showSerName", showSeriesName);
            ReplaceValChild(labels, "showPercent", showPercent);
            ReplaceValChild(labels, "showBubbleSize", false);
            NormalizeDataLabelsOrder(labels);
        }

        private static void SetDataLabelPosition(OfficeXmlNode labels, OfficeChartDataLabelPosition? position) {
            OfficeXmlNode existing = labels.Child(CNs, "dLblPos");
            if (existing.Exists) existing.Remove();
            if (position != null) {
                ReplaceValChild(labels, "dLblPos", position.Value.ToOpenXmlToken());
            }
            NormalizeDataLabelsOrder(labels);
        }

        private static OfficeChartDataLabelPosition? GetPowerPointCompatibleDataLabelPosition(
            OfficeXmlNode chartElement,
            OfficeChartDataLabelPosition position) {
            if (chartElement.Is(CNs, "doughnutChart") && position == OfficeChartDataLabelPosition.BestFit) {
                // PowerPoint repairs doughnut charts that explicitly serialize bestFit; omitting it keeps the default behavior.
                return null;
            }

            return position;
        }

        private static void SetNumberFormatChild(OfficeXmlNode label, string formatCode, bool sourceLinked) {
            OfficeXmlNode format = ReplaceChild(label, "numFmt");
            format.SetAttribute("formatCode", formatCode);
            format.SetBooleanAttribute("sourceLinked", sourceLinked);
        }

        private static void SetDataLabelNumberFormat(OfficeXmlNode labels, string formatCode, bool sourceLinked) {
            SetNumberFormatChild(labels, formatCode, sourceLinked);
            NormalizeDataLabelsOrder(labels);
        }

        private static void ApplyDataLabelOverrides(OfficeXmlNode label, bool? showLegendKey, bool? showValue,
            bool? showCategoryName, bool? showSeriesName, bool? showPercent,
            OfficeChartDataLabelPosition? position, string? numberFormat, bool sourceLinked) {
            if (showLegendKey != null) {
                ReplaceValChild(label, "showLegendKey", showLegendKey.Value);
            }
            if (showValue != null) {
                ReplaceValChild(label, "showVal", showValue.Value);
            }
            if (showCategoryName != null) {
                ReplaceValChild(label, "showCatName", showCategoryName.Value);
            }
            if (showSeriesName != null) {
                ReplaceValChild(label, "showSerName", showSeriesName.Value);
            }
            if (showPercent != null) {
                ReplaceValChild(label, "showPercent", showPercent.Value);
            }
            ReplaceValChild(label, "showBubbleSize", false);
            if (position != null) {
                ReplaceValChild(label, "dLblPos", position.Value.ToOpenXmlToken());
            }
            if (numberFormat != null) {
                SetNumberFormatChild(label, numberFormat, sourceLinked);
            }

            NormalizeLabelOrder(label);
        }

        private void ApplyDataLabelTextStyle(OfficeXmlNode labels, OfficeLength? fontSize, bool? bold,
            bool? italic, OfficeColor? color, string? fontFamily) {
            ApplyTextStyle(EnsureTextPropertiesRunProperties(labels), fontSize, bold, italic, color, fontFamily);
            NormalizeLabelOrder(labels);
        }

        private void ApplyDataLabelShapeStyle(OfficeXmlNode labels, OfficeColor? fillColor, OfficeColor? lineColor,
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
                ApplyOptionalLine(props, lineColor, lineWidth);
            }

            NormalizeLabelOrder(labels);
        }

        private void ApplyDataLabelLeaderLines(OfficeXmlNode labels, bool showLeaderLines, OfficeColor? lineColor,
            OfficeLength? lineWidth) {
            ReplaceValChild(labels, "showLeaderLines", showLeaderLines);

            if (showLeaderLines || lineColor != null || lineWidth != null) {
                OfficeXmlNode leaderLines = GetOrAppend(labels, CNs, "leaderLines");
                if (lineColor != null || lineWidth != null) {
                    ApplyOptionalLine(GetOrAppend(leaderLines, CNs, "spPr"), lineColor, lineWidth);
                }
            } else {
                OfficeXmlNode existing = labels.Child(CNs, "leaderLines");
                if (existing.Exists) existing.Remove();
            }

            NormalizeDataLabelsOrder(labels);
        }

        private static void ApplyDataLabelSeparator(OfficeXmlNode label, string? separator) {
            OfficeXmlNode existing = label.Child(CNs, "separator");
            if (existing.Exists) existing.Remove();
            if (separator != null) {
                AppendNew(label, CNs, "separator").SetInnerText(separator);
            }
            NormalizeLabelOrder(label);
        }

        private void ApplyDataLabelTemplate(OfficeXmlNode series, PowerPointChartDataLabelTemplate template) {
            if (template.NumberFormat != null && string.IsNullOrWhiteSpace(template.NumberFormat)) {
                throw new ArgumentException("Number format cannot be empty.", nameof(template.NumberFormat));
            }
            if (template.Separator != null && string.IsNullOrWhiteSpace(template.Separator)) {
                throw new ArgumentException("Separator cannot be empty.", nameof(template.Separator));
            }

            ApplyDataLabelsTemplate(EnsureDataLabels(series), template);
        }

        private void ApplyDataLabelsTemplate(OfficeXmlNode labels, PowerPointChartDataLabelTemplate template) {
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
                ValidateTextStyle(template.FontSize, template.TextColor, template.FontFamily);
            }
            if (applyShapeStyle) {
                ValidateAreaStyle(template.FillColor, template.LineColor, template.LineWidth,
                    template.NoFill, template.NoLine);
            }
            if (applyLeaderLines) {
                ValidateDataLabelLeaderLines(template.LeaderLineColor, template.LeaderLineWidth);
            }

            ApplyDataLabelOverrides(labels, template.ShowLegendKey, template.ShowValue, template.ShowCategoryName,
                template.ShowSeriesName, template.ShowPercent, template.Position, template.NumberFormat,
                template.SourceLinked);

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
                ApplyDataLabelLeaderLines(labels, showLeaderLines, template.LeaderLineColor,
                    template.LeaderLineWidth);
            }
        }

        private void ApplyPointDataLabelTemplate(OfficeXmlNode label, PowerPointChartDataLabelTemplate template) {
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

            if (template.NumberFormat != null && string.IsNullOrWhiteSpace(template.NumberFormat)) {
                throw new ArgumentException("Number format cannot be empty.", nameof(template.NumberFormat));
            }
            if (template.Separator != null && string.IsNullOrWhiteSpace(template.Separator)) {
                throw new ArgumentException("Separator cannot be empty.", nameof(template.Separator));
            }
            if (applyTextStyle) {
                ValidateTextStyle(template.FontSize, template.TextColor, template.FontFamily);
            }
            if (applyShapeStyle) {
                ValidateAreaStyle(template.FillColor, template.LineColor, template.LineWidth,
                    template.NoFill, template.NoLine);
            }

            ApplyDataLabelOverrides(label, template.ShowLegendKey, template.ShowValue, template.ShowCategoryName,
                template.ShowSeriesName, template.ShowPercent, template.Position, template.NumberFormat,
                template.SourceLinked);

            if (template.Separator != null) {
                ApplyDataLabelSeparator(label, template.Separator);
            }
            if (applyTextStyle) {
                ApplyDataLabelTextStyle(label, template.FontSize, template.Bold, template.Italic,
                    template.TextColor, template.FontFamily);
            }
            if (applyShapeStyle) {
                ApplyDataLabelShapeStyle(label, template.FillColor, template.LineColor, template.LineWidth,
                    template.NoFill, template.NoLine);
            }
        }

        private void ApplySeriesLeaderLineTemplate(OfficeXmlNode series,
            PowerPointChartDataLabelTemplate template) {
            bool applyLeaderLines = template.ShowLeaderLines == true
                || template.LeaderLineColor != null
                || template.LeaderLineWidth != null;
            if (!applyLeaderLines) {
                return;
            }

            ValidateDataLabelLeaderLines(template.LeaderLineColor, template.LeaderLineWidth);
            ApplyDataLabelLeaderLines(EnsureDataLabels(series), template.ShowLeaderLines ?? true,
                template.LeaderLineColor, template.LeaderLineWidth);
        }

        private static OfficeXmlNode EnsureDataLabels(OfficeXmlNode chartElement) {
            OfficeXmlNode labels = chartElement.Child(CNs, "dLbls");
            if (labels.Exists) {
                return labels;
            }

            OfficeXmlNode insertBefore;
            if (IsSeriesElement(chartElement)) {
                insertBefore = FirstChildOf(chartElement, "trendline", "errBars", "cat", "val", "xVal", "yVal", "bubbleSize", "smooth");
            } else {
                insertBefore = FirstChildOf(chartElement, "gapWidth", "overlap", "bubble3D", "bubbleScale", "showNegBubbles", "sizeRepresents", "axId");
                if (insertBefore.IsNone) insertBefore = GenericExtensionList(chartElement);
            }

            return InsertBeforeOrAppend(chartElement, chartElement.CreateElement(CNs, "dLbls"), insertBefore);
        }

        private static OfficeXmlNode FindDataLabel(OfficeXmlNode labels, uint index) {
            foreach (OfficeXmlNode label in labels.Children(CNs, "dLbl")) {
                if (label.Child(CNs, "idx").UInt32Attribute("val") == index) return label;
            }
            return default;
        }

        private static OfficeXmlNode EnsureDataLabel(OfficeXmlNode series, int pointIndex) {
            OfficeXmlNode labels = EnsureDataLabels(series);
            uint index = (uint)pointIndex;
            OfficeXmlNode label = FindDataLabel(labels, index);

            if (label.IsNone) {
                OfficeXmlNode insertBefore = default;
                foreach (OfficeXmlNode child in labels.Children()) {
                    if (!child.Is(CNs, "dLbl")) {
                        insertBefore = child;
                        break;
                    }
                }
                label = InsertBeforeOrAppend(labels, labels.CreateElement(CNs, "dLbl"), insertBefore);
                AppendNew(label, CNs, "idx").SetUInt32Attribute("val", index);
            }

            return label;
        }

        private static void RemoveDataLabels(OfficeXmlNode chartElement) {
            OfficeXmlNode labels = chartElement.Child(CNs, "dLbls");
            if (labels.Exists) labels.Remove();
        }

        private static void ClearDataLabel(OfficeXmlNode series, int pointIndex) {
            OfficeXmlNode labels = series.Child(CNs, "dLbls");
            if (labels.IsNone) {
                return;
            }

            OfficeXmlNode label = FindDataLabel(labels, (uint)pointIndex);
            if (label.IsNone) {
                return;
            }

            label.Remove();
            NormalizeDataLabelsOrder(labels);
        }

        private static OfficeXmlNode EnsureDataLabelShapeProperties(OfficeXmlNode labels) => GetOrAppend(labels, CNs, "spPr");

        private static void NormalizeLabelOrder(OfficeXmlNode label) {
            if (label.Is(CNs, "dLbls")) {
                NormalizeDataLabelsOrder(label);
            } else if (label.Is(CNs, "dLbl")) {
                NormalizeDataLabelOrder(label);
            }
        }

        private static void NormalizeDataLabelsOrder(OfficeXmlNode labels) {
            var ordered = new List<OfficeXmlNode>();
            foreach (OfficeXmlNode label in labels.Children(CNs, "dLbl")) ordered.Add(label);
            AddSettingsAndOthers(labels, ordered, "dLbl", DataLabelsSettings);
            Reorder(labels, ordered);
        }

        private static void NormalizeDataLabelOrder(OfficeXmlNode label) {
            OfficeXmlNode index = label.Child(CNs, "idx");
            if (index.IsNone) {
                return;
            }

            var ordered = new List<OfficeXmlNode> { index };
            AddSettingsAndOthers(label, ordered, "idx", DataLabelSettings);
            Reorder(label, ordered);
        }

        // Adds the first of each setting in order, then every child that is neither a setting nor the leading element.
        // Further copies of a setting are dropped, as the SDK rebuild dropped them.
        private static void AddSettingsAndOthers(OfficeXmlNode parent, List<OfficeXmlNode> ordered, string leading, string[] settings) {
            foreach (string setting in settings) {
                OfficeXmlNode child = parent.Child(CNs, setting);
                if (child.Exists) ordered.Add(child);
            }
            foreach (OfficeXmlNode child in parent.Children()) {
                if (child.NamespaceId == CNs && (child.LocalName == leading || Array.IndexOf(settings, child.LocalName) >= 0)) continue;
                ordered.Add(child);
            }
        }

        private static void Reorder(OfficeXmlNode parent, List<OfficeXmlNode> ordered) {
            foreach (OfficeXmlNode child in parent.Children().ToList()) child.Remove();
            foreach (OfficeXmlNode child in ordered) parent.Append(child);
        }

    }
}
