using System;
using System.Linq;
using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal;

internal static partial class OfficeOpenXmlChartSeriesReader {
    internal sealed class LabelLayout {
        internal bool Values, Categories, SeriesNames, Percentages, LeaderLines;
        internal double? FontSize;
        internal OfficeFontStyle? FontStyle;
        internal OfficeColor? TextColor;
        internal bool Visible => Values || Categories || SeriesNames || Percentages;
        internal string? Separator, NumberFormat;
        internal OfficeChartDataLabelPosition Position;
        internal bool SameAs(LabelLayout other) => Values == other.Values && Categories == other.Categories &&
            SeriesNames == other.SeriesNames && Percentages == other.Percentages && Separator == other.Separator &&
            NumberFormat == other.NumberFormat && Position == other.Position &&
            LeaderLines == other.LeaderLines && FontSize == other.FontSize &&
            FontStyle == other.FontStyle && Nullable.Equals(TextColor, other.TextColor);
    }

    internal static LabelLayout ReadLabels(C.Chart chart, A.ColorScheme? scheme = null) {
        LabelLayout? selected = null;
        foreach (var layer in chart.PlotArea?.ChildElements.OfType<OpenXmlCompositeElement>()
            .Where(item => item.LocalName.EndsWith("Chart", StringComparison.Ordinal)) ?? Enumerable.Empty<OpenXmlCompositeElement>()) {
            var labels = layer.GetFirstChild<C.DataLabels>();
            if (labels?.GetFirstChild<C.Separator>() is C.Separator separator && string.IsNullOrEmpty(separator.Text))
                throw new NotSupportedException("An empty native data label separator cannot be projected.");
            var current = new LabelLayout {
                Values = LabelFlag<C.ShowValue>(labels), Categories = LabelFlag<C.ShowCategoryName>(labels),
                SeriesNames = LabelFlag<C.ShowSeriesName>(labels), Percentages = LabelFlag<C.ShowPercent>(labels),
                LeaderLines = LabelFlag<C.ShowLeaderLines>(labels),
                Separator = labels?.GetFirstChild<C.Separator>()?.Text,
                NumberFormat = labels?.GetFirstChild<C.NumberingFormat>()?.FormatCode?.Value,
                Position = labels?.GetFirstChild<C.DataLabelPosition>()?.Val?.InnerText switch {
                    null or "bestFit" => OfficeChartDataLabelPosition.BestFit,
                    "b" => OfficeChartDataLabelPosition.Bottom, "t" => OfficeChartDataLabelPosition.Top,
                    "l" => OfficeChartDataLabelPosition.Left, "r" => OfficeChartDataLabelPosition.Right,
                    "ctr" => OfficeChartDataLabelPosition.Center, "inBase" => OfficeChartDataLabelPosition.InsideBase,
                    "inEnd" => OfficeChartDataLabelPosition.InsideEnd, "outEnd" => OfficeChartDataLabelPosition.OutsideEnd,
                    _ => throw new NotSupportedException("The native data label position cannot be projected.")
                }
            };
            if (current.Visible && (layer is C.PieChart or C.DoughnutChart) &&
                current.Position != OfficeChartDataLabelPosition.BestFit)
                throw new NotSupportedException("The native radial data label position cannot be projected.");
            if (layer.ChildElements.OfType<OpenXmlCompositeElement>().Where(item => item.LocalName == "ser")
                .Any(item => item.GetFirstChild<C.DataLabels>() is C.DataLabels seriesLabels &&
                    (current.Visible || HasVisibleLabelContent(seriesLabels))))
                throw new NotSupportedException("Series-specific chart labels cannot be projected by this layout reader.");
            if (!current.Visible && labels != null && HasVisibleLabelContent(labels))
                throw new NotSupportedException("The native data label overrides cannot be projected.");
            if (!current.Visible) {
                current.Separator = null;
                current.NumberFormat = null;
                current.Position = OfficeChartDataLabelPosition.BestFit;
                current.LeaderLines = false;
            }
            if (current.Visible && (labels?.Elements<C.DataLabel>().Any() == true || LabelFlag<C.ShowLegendKey>(labels) ||
                LabelFlag<C.ShowBubbleSize>(labels) ||
                (current.LeaderLines && layer is C.PieChart or C.DoughnutChart && current.Position == OfficeChartDataLabelPosition.BestFit) ||
                (current.LeaderLines && !(layer is C.PieChart or C.DoughnutChart && current.Position == OfficeChartDataLabelPosition.OutsideEnd) &&
                    current.Position is not OfficeChartDataLabelPosition.BestFit and
                    not OfficeChartDataLabelPosition.Center and not OfficeChartDataLabelPosition.InsideBase and not OfficeChartDataLabelPosition.InsideEnd)))
                throw new NotSupportedException("The native data label overrides cannot be projected.");
            if (current.Visible && labels != null) {
                NativeText text = ReadNativeText(chart, labels, scheme);
                current.FontSize = text.Size;
                current.FontStyle = text.Style;
                current.TextColor = text.Color;
                if (HasUnsupportedSharedAxisNumberFormat(labels))
                    throw new NotSupportedException("The native data label format cannot be projected.");
                foreach (var child in labels.ChildElements) {
                    if (child is C.TextProperties or C.ShowValue or C.ShowCategoryName or C.ShowSeriesName or C.ShowPercent or C.ShowLegendKey or
                        C.ShowBubbleSize or C.ShowLeaderLines or C.Separator or C.NumberingFormat or C.DataLabelPosition) continue;
                    if (child is C.LeaderLines leader && !leader.HasChildren && !leader.HasAttributes) continue;
                    throw new NotSupportedException("The native data label appearance cannot be projected.");
                }
                if (labels.GetFirstChild<C.NumberingFormat>()?.SourceLinked?.Value == true)
                    throw new NotSupportedException("Source-linked data label formats require qualified workbook formatting.");
            }
            if (selected != null && !selected.SameAs(current))
                throw new NotSupportedException("Different label layouts across chart layers cannot be projected together.");
            selected = current;
        }
        return selected ?? new LabelLayout();
    }

    private static bool LabelFlag<T>(C.DataLabels? labels) where T : C.BooleanType =>
        labels?.GetFirstChild<T>() is T flag && flag.Val?.Value != false;

    private static bool HasVisibleLabelContent(C.DataLabels labels) =>
        LabelFlag<C.ShowValue>(labels) || LabelFlag<C.ShowCategoryName>(labels) ||
        LabelFlag<C.ShowSeriesName>(labels) || LabelFlag<C.ShowPercent>(labels) ||
        LabelFlag<C.ShowLegendKey>(labels) || LabelFlag<C.ShowBubbleSize>(labels) ||
        labels.Elements<C.DataLabel>().Any(label => label.GetFirstChild<C.ChartText>() != null ||
            label.Descendants<C.BooleanType>().Any(flag => flag.Val?.Value != false &&
                flag is C.ShowValue or C.ShowCategoryName or C.ShowSeriesName or C.ShowPercent or C.ShowLegendKey or C.ShowBubbleSize));
}
