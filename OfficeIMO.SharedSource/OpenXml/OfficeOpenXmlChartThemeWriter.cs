using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal;

/// <summary>Writes theme defaults into newly authored charts; imported charts keep their formatting.</summary>
internal static class OfficeOpenXmlChartThemeWriter {
    internal static void Apply(ChartPart part, OfficeTheme theme) {
        C.ChartSpace? space = part.ChartSpace;
        C.Chart? chart = space?.GetFirstChild<C.Chart>();
        if (space == null || chart == null) return;
        OfficeThemeChartStyle style = theme.ChartStyle;
        AddTextDefaults(space, style.AxisFont.Family, style.AxisFontSize, theme.Palette.Text);
        C.Title? title = chart.GetFirstChild<C.Title>();
        if (title != null) ApplyTitle(title, theme);
        C.Legend? legend = chart.GetFirstChild<C.Legend>();
        if (legend != null) {
            if (style.ShowLegend) AddTextDefaults(legend, style.AxisFont.Family, style.AxisFontSize, theme.Palette.Text);
            else legend.Remove();
        }
        C.PlotArea? plot = chart.GetFirstChild<C.PlotArea>();
        if (plot == null) return;
        foreach (OpenXmlCompositeElement axis in plot.ChildElements.OfType<OpenXmlCompositeElement>()
                     .Where(axis => axis is C.CategoryAxis or C.DateAxis or C.ValueAxis or C.SeriesAxis)) {
            AddTextDefaults(axis, style.AxisFont.Family, style.AxisFontSize, style.AxisTextColor.Resolve(theme));
            if (axis.GetFirstChild<C.ChartShapeProperties>() == null &&
                axis.GetFirstChild<C.Delete>()?.Val?.Value != true) {
                var outline = new A.Outline { Width = 6350 };
                if (axis is C.ValueAxis) outline.Append(new A.NoFill());
                else outline.Append(Fill(style.AxisLineColor.Resolve(theme)));
                axis.AddChild(new C.ChartShapeProperties(outline), true);
            }
            foreach (C.MajorGridlines grid in axis.Elements<C.MajorGridlines>()) {
                grid.RemoveAllChildren<C.ChartShapeProperties>();
                grid.AddChild(
                    new C.ChartShapeProperties(new A.Outline(Fill(style.GridlineColor.Resolve(theme))) {
                        Width = ToEmus(style.GridlineWidth)
                    }), true);
            }
        }
        foreach (C.BarChart bars in plot.Elements<C.BarChart>()) {
            C.GapWidth gap = bars.GetFirstChild<C.GapWidth>() ?? new C.GapWidth();
            if (gap.Parent == null) bars.AddChild(gap, true);
            gap.Val = checked((ushort)style.BarGapPercent);
        }
        foreach (C.LineChartSeries series in plot.Descendants<C.LineChartSeries>()) {
            C.ChartShapeProperties properties = series.GetFirstChild<C.ChartShapeProperties>() ?? new C.ChartShapeProperties();
            A.Outline outline = properties.GetFirstChild<A.Outline>() ?? new A.Outline();
            if (outline.Width == null) outline.Width = ToEmus(style.LineWidth);
            if (outline.Parent == null) properties.AddChild(outline, true);
            if (properties.Parent == null) series.AddChild(properties, true);
        }
        foreach (C.DataLabels labels in plot.Descendants<C.DataLabels>()) {
            C.ShowValue? showValue = labels.GetFirstChild<C.ShowValue>();
            if (showValue == null) labels.AddChild(new C.ShowValue { Val = style.ShowDataLabels }, true);
            AddTextDefaults(labels, style.AxisFont.Family, style.AxisFontSize, theme.Palette.Text);
        }
    }

    internal static void ApplyTitle(C.Title title, OfficeTheme theme) => AddTextDefaults(title,
        theme.ChartStyle.TitleFont.Family, theme.ChartStyle.TitleFontSize, theme.Palette.Text);

    internal static void RestoreAxisTextDefaults(A.TextCharacterPropertiesType properties, OfficeTheme theme) {
        properties.FontSize = checked((int)Math.Round(theme.ChartStyle.AxisFontSize.ToPoints() * 100D));
        properties.RemoveAllChildren<A.SolidFill>();
        properties.RemoveAllChildren<A.LatinFont>();
        properties.AddChild(Fill(theme.ChartStyle.AxisTextColor.Resolve(theme)), true);
        properties.AddChild(new A.LatinFont { Typeface = theme.ChartStyle.AxisFont.Family }, true);
    }

    private static void AddTextDefaults(OpenXmlCompositeElement owner, string family, OfficeLength size, OfficeColor color) {
        if (owner.GetFirstChild<C.TextProperties>() != null) return;
        owner.AddChild(new C.TextProperties(new A.BodyProperties(), new A.ListStyle(),
            new A.Paragraph(new A.ParagraphProperties(new A.DefaultRunProperties(
                Fill(color), new A.LatinFont { Typeface = family }) {
                FontSize = checked((int)Math.Round(size.ToPoints() * 100D))
            }), new A.EndParagraphRunProperties())), true);
    }

    private static A.SolidFill Fill(OfficeColor color) => new(new A.RgbColorModelHex { Val = color.ToRgbHex() });
    private static int ToEmus(OfficeLength value) => checked((int)Math.Round(value.ToPoints() * 12700D));
}
