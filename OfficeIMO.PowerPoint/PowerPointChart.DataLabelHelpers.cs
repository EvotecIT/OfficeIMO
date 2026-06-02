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
        private static void ApplyDataLabels(OpenXmlCompositeElement chartElement, bool showLegendKey, bool showValue,
            bool showCategoryName, bool showSeriesName, bool showPercent) {
            C.DataLabels labels = EnsureDataLabels(chartElement);
            ReplaceChild(labels, new C.ShowLegendKey { Val = showLegendKey });
            ReplaceChild(labels, new C.ShowValue { Val = showValue });
            ReplaceChild(labels, new C.ShowCategoryName { Val = showCategoryName });
            ReplaceChild(labels, new C.ShowSeriesName { Val = showSeriesName });
            ReplaceChild(labels, new C.ShowPercent { Val = showPercent });
            ReplaceChild(labels, new C.ShowBubbleSize { Val = false });
            NormalizeDataLabelsOrder(labels);
        }

        private static void SetDataLabelPosition(C.DataLabels labels, C.DataLabelPositionValues? position) {
            labels.GetFirstChild<C.DataLabelPosition>()?.Remove();
            if (position != null) {
                ReplaceChild(labels, new C.DataLabelPosition { Val = position.Value });
            }
            NormalizeDataLabelsOrder(labels);
        }

        private static C.DataLabelPositionValues? GetPowerPointCompatibleDataLabelPosition(
            OpenXmlElement chartElement,
            C.DataLabelPositionValues position) {
            if (chartElement is C.DoughnutChart && position == C.DataLabelPositionValues.BestFit) {
                // PowerPoint repairs doughnut charts that explicitly serialize bestFit; omitting it keeps the default behavior.
                return null;
            }

            return position;
        }

        private static void SetDataLabelNumberFormat(C.DataLabels labels, string formatCode, bool sourceLinked) {
            ReplaceChild(labels, new C.NumberingFormat {
                FormatCode = formatCode,
                SourceLinked = sourceLinked
            });
            NormalizeDataLabelsOrder(labels);
        }

        private static void ApplyDataLabelOverrides(OpenXmlCompositeElement label, bool? showLegendKey, bool? showValue,
            bool? showCategoryName, bool? showSeriesName, bool? showPercent,
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
            ReplaceChild(label, new C.ShowBubbleSize { Val = false });
            if (position != null) {
                ReplaceChild(label, new C.DataLabelPosition { Val = position.Value });
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
                ApplySolidFill(props, fillColor);
            }

            if (noLine) {
                ApplyNoLine(props);
            } else if (lineColor != null || lineWidthPoints != null) {
                ApplyOptionalLine(props, lineColor, lineWidthPoints);
            }

            if (labels is C.DataLabels dataLabels) {
                NormalizeDataLabelsOrder(dataLabels);
            } else if (labels is C.DataLabel dataLabel) {
                NormalizeDataLabelOrder(dataLabel);
            }
        }

        private static void ApplyDataLabelLeaderLines(C.DataLabels labels, bool showLeaderLines, string? lineColor,
            double? lineWidthPoints) {
            ReplaceChild(labels, new C.ShowLeaderLines { Val = showLeaderLines });

            if (showLeaderLines || lineColor != null || lineWidthPoints != null) {
                C.LeaderLines leaderLines = labels.GetFirstChild<C.LeaderLines>() ?? new C.LeaderLines();
                C.ChartShapeProperties props = leaderLines.GetFirstChild<C.ChartShapeProperties>() ?? new C.ChartShapeProperties();
                if (lineColor != null || lineWidthPoints != null) {
                    ApplyOptionalLine(props, lineColor, lineWidthPoints);
                }
                if ((lineColor != null || lineWidthPoints != null) && props.Parent == null) {
                    leaderLines.Append(props);
                }
                if (leaderLines.Parent == null) {
                    labels.Append(leaderLines);
                }
            } else {
                labels.GetFirstChild<C.LeaderLines>()?.Remove();
            }

            NormalizeDataLabelsOrder(labels);
        }

        private static void ApplyDataLabelSeparator(OpenXmlCompositeElement label, string? separator) {
            C.Separator? existing = label.GetFirstChild<C.Separator>();
            if (separator == null) {
                existing?.Remove();
                if (label is C.DataLabels dataLabels) {
                    NormalizeDataLabelsOrder(dataLabels);
                } else if (label is C.DataLabel dataLabel) {
                    NormalizeDataLabelOrder(dataLabel);
                }
                return;
            }

            existing?.Remove();
            label.Append(new C.Separator { Text = separator });
            if (label is C.DataLabels allDataLabels) {
                NormalizeDataLabelsOrder(allDataLabels);
            } else if (label is C.DataLabel dataLabel) {
                NormalizeDataLabelOrder(dataLabel);
            }
        }

        private static void ApplyDataLabelTemplate(OpenXmlCompositeElement series, PowerPointChartDataLabelTemplate template) {
            if (template.NumberFormat != null && string.IsNullOrWhiteSpace(template.NumberFormat)) {
                throw new ArgumentException("Number format cannot be empty.", nameof(template.NumberFormat));
            }
            if (template.Separator != null && string.IsNullOrWhiteSpace(template.Separator)) {
                throw new ArgumentException("Separator cannot be empty.", nameof(template.Separator));
            }

            ApplyDataLabelTemplate(EnsureDataLabels(series), template);
        }

        private static void ApplyDataLabelTemplate(C.DataLabels labels, PowerPointChartDataLabelTemplate template) {
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
                ValidateTextStyle(template.FontSizePoints, template.TextColor, template.FontName);
            }
            if (applyShapeStyle) {
                ValidateAreaStyle(template.FillColor, template.LineColor, template.LineWidthPoints,
                    template.NoFill, template.NoLine);
            }
            if (applyLeaderLines) {
                ValidateDataLabelLeaderLines(template.LeaderLineColor, template.LeaderLineWidthPoints);
            }

            ApplyDataLabelOverrides(labels, template.ShowLegendKey, template.ShowValue, template.ShowCategoryName,
                template.ShowSeriesName, template.ShowPercent, template.Position, template.NumberFormat,
                template.SourceLinked);

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
                ApplyDataLabelLeaderLines(labels, showLeaderLines, template.LeaderLineColor,
                    template.LeaderLineWidthPoints);
            }
        }

        private static void ApplyDataLabelTemplate(C.DataLabel label, PowerPointChartDataLabelTemplate template) {
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

            if (template.NumberFormat != null && string.IsNullOrWhiteSpace(template.NumberFormat)) {
                throw new ArgumentException("Number format cannot be empty.", nameof(template.NumberFormat));
            }
            if (template.Separator != null && string.IsNullOrWhiteSpace(template.Separator)) {
                throw new ArgumentException("Separator cannot be empty.", nameof(template.Separator));
            }
            if (applyTextStyle) {
                ValidateTextStyle(template.FontSizePoints, template.TextColor, template.FontName);
            }
            if (applyShapeStyle) {
                ValidateAreaStyle(template.FillColor, template.LineColor, template.LineWidthPoints,
                    template.NoFill, template.NoLine);
            }

            ApplyDataLabelOverrides(label, template.ShowLegendKey, template.ShowValue, template.ShowCategoryName,
                template.ShowSeriesName, template.ShowPercent, template.Position, template.NumberFormat,
                template.SourceLinked);

            if (template.Separator != null) {
                ApplyDataLabelSeparator(label, template.Separator);
            }
            if (applyTextStyle) {
                ApplyDataLabelTextStyle(label, template.FontSizePoints, template.Bold, template.Italic,
                    template.TextColor, template.FontName);
            }
            if (applyShapeStyle) {
                ApplyDataLabelShapeStyle(label, template.FillColor, template.LineColor, template.LineWidthPoints,
                    template.NoFill, template.NoLine);
            }
        }

        private static void ApplySeriesLeaderLineTemplate(OpenXmlCompositeElement series,
            PowerPointChartDataLabelTemplate template) {
            bool applyLeaderLines = template.ShowLeaderLines == true
                || template.LeaderLineColor != null
                || template.LeaderLineWidthPoints != null;
            if (!applyLeaderLines) {
                return;
            }

            ValidateDataLabelLeaderLines(template.LeaderLineColor, template.LeaderLineWidthPoints);
            ApplyDataLabelLeaderLines(EnsureDataLabels(series), template.ShowLeaderLines ?? true,
                template.LeaderLineColor, template.LeaderLineWidthPoints);
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
                insertBefore ??= chartElement.GetFirstChild<C.Overlap>();
                insertBefore ??= chartElement.GetFirstChild<C.BubbleScale>();
                insertBefore ??= chartElement.GetFirstChild<C.ShowNegativeBubbles>();
                insertBefore ??= chartElement.GetFirstChild<C.SizeRepresents>();
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

        private static C.DataLabel EnsureDataLabel(OpenXmlCompositeElement series, int pointIndex) {
            C.DataLabels labels = EnsureDataLabels(series);
            uint index = (uint)pointIndex;
            C.DataLabel? label = labels.Elements<C.DataLabel>()
                .FirstOrDefault(item => item.GetFirstChild<C.Index>()?.Val?.Value == index);

            if (label == null) {
                label = new C.DataLabel(new C.Index { Val = index });
                OpenXmlElement? insertBefore = labels.ChildElements.FirstOrDefault(child => child is not C.DataLabel);
                if (insertBefore != null) {
                    labels.InsertBefore(label, insertBefore);
                } else {
                    labels.Append(label);
                }
            }

            return label;
        }

        private static void RemoveDataLabels(OpenXmlCompositeElement chartElement) {
            chartElement.GetFirstChild<C.DataLabels>()?.Remove();
        }

        private static void ClearDataLabel(OpenXmlCompositeElement series, int pointIndex) {
            C.DataLabels? labels = series.GetFirstChild<C.DataLabels>();
            if (labels == null) {
                return;
            }

            uint index = (uint)pointIndex;
            C.DataLabel? label = labels.Elements<C.DataLabel>()
                .FirstOrDefault(item => item.GetFirstChild<C.Index>()?.Val?.Value == index);
            if (label == null) {
                return;
            }

            label.Remove();
            NormalizeDataLabelsOrder(labels);
        }

        private static C.ChartShapeProperties EnsureDataLabelShapeProperties(OpenXmlCompositeElement labels) {
            C.ChartShapeProperties props = labels.GetFirstChild<C.ChartShapeProperties>() ?? new C.ChartShapeProperties();
            if (props.Parent == null) {
                labels.Append(props);
            }

            return props;
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

            foreach (C.DataLabel dataLabel in overrides) {
                labels.Append(dataLabel);
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

            foreach (OpenXmlElement otherChild in otherChildren) {
                labels.Append(otherChild);
            }

            if (extLst != null) {
                labels.Append(extLst);
            }
        }

        private static void NormalizeDataLabelOrder(C.DataLabel label) {
            C.Index? index = label.GetFirstChild<C.Index>();
            if (index == null) {
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
            label.Append(index);
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
            foreach (OpenXmlElement otherChild in otherChildren) {
                label.Append(otherChild);
            }
            if (extLst != null) {
                label.Append(extLst);
            }
        }

    }
}
