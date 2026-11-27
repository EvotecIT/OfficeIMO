using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using static OfficeIMO.OpenXml.Internal.OfficeChartXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal;

/// <summary>Writes theme defaults into newly authored charts; imported charts keep their formatting.</summary>
internal static class OfficeOpenXmlChartThemeWriter {
    internal static void Apply(IOfficeChartPart part, OfficeTheme theme) {
        OfficeXmlNode space = part.ChartSpace;
        OfficeXmlNode chart = space.Child(NsC, "chart");
        if (space.IsNone || chart.IsNone) return;
        OfficeThemeChartStyle style = theme.ChartStyle;
        AddTextDefaults(space, style.AxisFont.Family, style.AxisFontSize, theme.Palette.Text);
        OfficeXmlNode title = chart.Child(NsC, "title");
        if (!title.IsNone) ApplyTitle(title, theme);
        OfficeXmlNode legend = chart.Child(NsC, "legend");
        if (!legend.IsNone) {
            if (style.ShowLegend) AddTextDefaults(legend, style.AxisFont.Family, style.AxisFontSize, theme.Palette.Text);
            else legend.Remove();
        }
        OfficeXmlNode plot = chart.Child(NsC, "plotArea");
        if (plot.IsNone) return;
        foreach (OfficeXmlNode axis in plot.Children()) {
            if (!(axis.NamespaceId == NsC && axis.LocalName is "catAx" or "dateAx" or "valAx" or "serAx")) continue;
            OfficeXmlChildOrder axisOrder = OrderOf(axis);
            AddTextDefaults(axis, style.AxisFont.Family, style.AxisFontSize, style.AxisTextColor.Resolve(theme));
            if (axis.Child(NsC, "spPr").IsNone &&
                axis.Child(NsC, "delete").BoolVal() != true) {
                OfficeXmlNode properties = axis.CreateElement(NsC, "spPr");
                OfficeXmlNode outline = properties.AppendNew(NsA, "ln");
                outline.SetAttribute("w", Text(6350));
                if (axis.Is(NsC, "valAx")) outline.AppendNew(NsA, "noFill");
                else AppendFill(outline, style.AxisLineColor.Resolve(theme));
                SetOrdered(axis, properties, axisOrder);
            }
            foreach (OfficeXmlNode grid in axis.Children(NsC, "majorGridlines")) {
                grid.RemoveChildren(NsC, "spPr");
                OfficeXmlNode properties = grid.CreateElement(NsC, "spPr");
                OfficeXmlNode outline = properties.AppendNew(NsA, "ln");
                outline.SetAttribute("w", Text(ToEmus(style.GridlineWidth)));
                AppendFill(outline, style.GridlineColor.Resolve(theme));
                SetOrdered(grid, properties, CChartLinesView.Order);
            }
        }
        foreach (OfficeXmlNode bars in plot.Children(NsC, "barChart")) {
            OfficeXmlNode gap = bars.Child(NsC, "gapWidth");
            if (gap.IsNone) gap = SetOrdered(bars, bars.CreateElement(NsC, "gapWidth"), CBarChartView.Order);
            gap.SetAttribute("val", checked((ushort)style.BarGapPercent).ToString(CultureInfo.InvariantCulture));
        }
        // c:ser of a line, 3-D line or stock chart, the SDK's LineChartSeries.
        foreach (OfficeXmlNode series in plot.DescendantsWhere(element => element.Is(NsC, "ser") &&
                     (element.Parent.Is(NsC, "lineChart") || element.Parent.Is(NsC, "line3DChart") || element.Parent.Is(NsC, "stockChart")))) {
            OfficeXmlNode properties = series.Child(NsC, "spPr");
            bool newProperties = properties.IsNone;
            if (newProperties) properties = series.CreateElement(NsC, "spPr");
            OfficeXmlNode outline = properties.Child(NsA, "ln");
            bool newOutline = outline.IsNone;
            if (newOutline) outline = properties.CreateElement(NsA, "ln");
            if (outline.Attribute("w") == null) outline.SetAttribute("w", Text(ToEmus(style.LineWidth)));
            if (newOutline) SetOrdered(properties, outline, DmlShapePropertiesView.Order);
            if (newProperties) SetOrdered(series, properties, CLineSerView.Order);
        }
        foreach (OfficeXmlNode labels in plot.ContentDescendants(NsC, "dLbls")) {
            if (labels.Child(NsC, "showVal").IsNone) {
                // c:delete and the formatting children are the two branches of one choice; adding c:showVal leaves the
                // formatting branch, as the SDK's AddChild did.
                labels.RemoveChildren(NsC, "delete");
                SetOrdered(labels, labels.CreateVal(NsC, "showVal", Text(style.ShowDataLabels)), CDLblsView.Order);
            }
            AddTextDefaults(labels, style.AxisFont.Family, style.AxisFontSize, theme.Palette.Text);
        }
    }

    internal static void ApplyTitle(OfficeXmlNode title, OfficeTheme theme) => AddTextDefaults(title,
        theme.ChartStyle.TitleFont.Family, theme.ChartStyle.TitleFontSize, theme.Palette.Text);

    internal static void RestoreAxisTextDefaults(OfficeXmlNode properties, OfficeTheme theme) {
        properties.SetAttribute("sz", Text(checked((int)Math.Round(theme.ChartStyle.AxisFontSize.ToPoints() * 100D))));
        properties.RemoveChildren(NsA, "solidFill");
        properties.RemoveChildren(NsA, "latin");
        OfficeXmlNode fill = properties.CreateElement(NsA, "solidFill");
        fill.AppendVal(NsA, "srgbClr", theme.ChartStyle.AxisTextColor.Resolve(theme).ToRgbHex());
        SetOrdered(properties, fill, DmlTextCharacterPropertiesView.Order);
        OfficeXmlNode latin = properties.CreateElement(NsA, "latin");
        latin.SetAttribute("typeface", theme.ChartStyle.AxisFont.Family);
        SetOrdered(properties, latin, DmlTextCharacterPropertiesView.Order);
    }

    private static void AddTextDefaults(OfficeXmlNode owner, string family, OfficeLength size, OfficeColor color) {
        if (!owner.Child(NsC, "txPr").IsNone) return;
        OfficeXmlNode text = owner.CreateElement(NsC, "txPr");
        text.AppendNew(NsA, "bodyPr");
        text.AppendNew(NsA, "lstStyle");
        OfficeXmlNode paragraph = text.AppendNew(NsA, "p");
        OfficeXmlNode paragraphProperties = paragraph.AppendNew(NsA, "pPr");
        OfficeXmlNode run = paragraphProperties.AppendNew(NsA, "defRPr");
        run.SetAttribute("sz", Text(checked((int)Math.Round(size.ToPoints() * 100D))));
        AppendFill(run, color);
        run.AppendNew(NsA, "latin").SetAttribute("typeface", family);
        paragraph.AppendNew(NsA, "endParaRPr");
        SetOrdered(owner, text, OrderOf(owner));
    }

    // The schema order of the elements that take text defaults.
    private static OfficeXmlChildOrder OrderOf(OfficeXmlNode owner) => owner.LocalName switch {
        "chartSpace" => CChartSpaceView.Order,
        "title" => CTitleView.Order,
        "legend" => CLegendView.Order,
        "catAx" => CCatAxView.Order,
        "dateAx" => CDateAxView.Order,
        "valAx" => CValAxView.Order,
        "serAx" => CSerAxView.Order,
        "dLbls" => CDLblsView.Order,
        _ => throw new InvalidOperationException("The element '" + owner.LocalName + "' takes no chart text defaults.")
    };

    // Appends a:solidFill holding a:srgbClr.
    private static void AppendFill(OfficeXmlNode parent, OfficeColor color) =>
        parent.AppendNew(NsA, "solidFill").AppendVal(NsA, "srgbClr", color.ToRgbHex());

    private static int ToEmus(OfficeLength value) => checked((int)Math.Round(value.ToPoints() * 12700D));

    // SDK callers
    internal static void Apply(ChartPart part, OfficeTheme theme) => Apply(new OfficeSdkChartPart(part), theme);
    internal static void ApplyTitle(C.Title title, OfficeTheme theme) => ApplyTitle(OfficeXmlNode.FromSdk(title), theme);
    internal static void RestoreAxisTextDefaults(A.TextCharacterPropertiesType properties, OfficeTheme theme) =>
        RestoreAxisTextDefaults(OfficeXmlNode.FromSdk(properties), theme);
}
