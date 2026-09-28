using System;
using System.Linq;
using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal;

/// <summary>Owns numeric settings on the referenced secondary value axis.</summary>
internal static class OfficeOpenXmlChartSecondaryAxis {
    internal static C.ValueAxis? Resolve(C.PlotArea? plot) {
        if (plot == null) return null;
        var groups = OfficeOpenXmlChartAxisGroups.Create(plot);
        var axes = plot.ChildElements.OfType<OpenXmlCompositeElement>()
            .Where(layer => layer.LocalName.EndsWith("Chart", StringComparison.Ordinal) &&
                groups.Read(layer) == OfficeChartAxisGroup.Secondary)
            .SelectMany(layer => layer.Elements<C.AxisId>())
            .Select(reference => groups.Resolve(reference.Val?.Value)).OfType<C.ValueAxis>().Distinct().ToArray();
        if (axes.Length > 1) throw new NotSupportedException("Multiple independent secondary value axes cannot be projected.");
        return axes.SingleOrDefault();
    }

    internal static OfficeChartValueAxisLayout? Read(C.PlotArea? plot) {
        var axis = Resolve(plot);
        if (axis == null) return null;
        var scaling = axis.GetFirstChild<C.Scaling>();
        var layout = new OfficeChartValueAxisLayout(
            minimum: scaling?.GetFirstChild<C.MinAxisValue>()?.Val?.Value,
            maximum: scaling?.GetFirstChild<C.MaxAxisValue>()?.Val?.Value,
            majorUnit: axis.GetFirstChild<C.MajorUnit>()?.Val?.Value,
            minorUnit: axis.GetFirstChild<C.MinorUnit>()?.Val?.Value,
            numberFormat: axis.GetFirstChild<C.NumberingFormat>()?.FormatCode?.Value,
            majorTickMark: ReadTick(axis.GetFirstChild<C.MajorTickMark>()?.Val?.Value),
            minorTickMark: ReadTick(axis.GetFirstChild<C.MinorTickMark>()?.Val?.Value));
        string? title = ReadTitle(axis.GetFirstChild<C.Title>());
        return title == null ? layout : layout.WithTitle(title);
    }

    internal static void Apply(C.Chart? chart, OfficeChartValueAxisLayout layout) {
        if (layout == null) throw new ArgumentNullException(nameof(layout));
        var axis = Resolve(chart?.PlotArea) ?? throw new InvalidOperationException("The chart has no referenced secondary value axis.");
        var scaling = axis.GetFirstChild<C.Scaling>() ?? new C.Scaling();
        if (scaling.Parent == null) axis.AddChild(scaling, true);
        scaling.RemoveAllChildren<C.MinAxisValue>();
        scaling.RemoveAllChildren<C.MaxAxisValue>();
        scaling.RemoveAllChildren<C.LogBase>();
        scaling.RemoveAllChildren<C.Orientation>();
        scaling.AddChild(new C.Orientation { Val = C.OrientationValues.MinMax }, true);
        if (layout.Minimum.HasValue) scaling.AddChild(new C.MinAxisValue { Val = layout.Minimum.Value }, true);
        if (layout.Maximum.HasValue) scaling.AddChild(new C.MaxAxisValue { Val = layout.Maximum.Value }, true);
        axis.RemoveAllChildren<C.MajorUnit>();
        axis.RemoveAllChildren<C.MinorUnit>();
        if (layout.MajorUnit.HasValue) axis.AddChild(new C.MajorUnit { Val = layout.MajorUnit.Value }, true);
        if (layout.MinorUnit.HasValue) axis.AddChild(new C.MinorUnit { Val = layout.MinorUnit.Value }, true);
        axis.RemoveAllChildren<C.NumberingFormat>();
        if (layout.NumberFormat != null) axis.AddChild(new C.NumberingFormat { FormatCode = layout.NumberFormat, SourceLinked = false }, true);
        if (layout.MajorTickMark.HasValue) {
            axis.RemoveAllChildren<C.MajorTickMark>();
            axis.AddChild(new C.MajorTickMark { Val = WriteTick(layout.MajorTickMark.Value) }, true);
        }
        if (layout.MinorTickMark.HasValue) {
            axis.RemoveAllChildren<C.MinorTickMark>();
            axis.AddChild(new C.MinorTickMark { Val = WriteTick(layout.MinorTickMark.Value) }, true);
        }
        if (layout.IsTitleSpecified) axis.GetFirstChild<C.Title>()?.Remove();
        if (layout.IsTitleSpecified && layout.Title != null) axis.AddChild(new C.Title(
            new C.ChartText(new C.RichText(new A.BodyProperties(), new A.ListStyle(),
                new A.Paragraph(new A.Run(new A.Text(layout.Title))))),
            new C.Layout(), new C.Overlay { Val = false }), true);
    }

    internal static void QualifyLinearProjection(C.PlotArea? plot, bool resolveSourceLinkedFormats = false) {
        var axis = Resolve(plot);
        if (axis == null) return;
        QualifyTitle(axis.GetFirstChild<C.Title>());
        if (axis.GetFirstChild<C.Delete>() is C.Delete secondaryDeletion && secondaryDeletion.Val?.Value != false)
            throw new NotSupportedException("A deleted secondary value axis cannot be projected.");
        var groups = OfficeOpenXmlChartAxisGroups.Create(plot!);
        var primaryLayer = plot!.ChildElements.OfType<OpenXmlCompositeElement>().FirstOrDefault(layer =>
            layer.LocalName.EndsWith("Chart", StringComparison.Ordinal) &&
            groups.Read(layer) == OfficeChartAxisGroup.Primary);
        var primaryValueAxis = primaryLayer?.Elements<C.AxisId>()
            .Select(reference => groups.Resolve(reference.Val?.Value)).OfType<C.ValueAxis>().SingleOrDefault();
        if ((primaryValueAxis?.GetFirstChild<C.Delete>() is C.Delete deletion && deletion.Val?.Value != false) ||
            primaryValueAxis?.GetFirstChild<C.TickLabelPosition>()?.Val?.Value == C.TickLabelPositionValues.None)
            throw new NotSupportedException("The primary and secondary value-axis visibility cannot be projected independently.");
        var scaling = axis.GetFirstChild<C.Scaling>();
        if ((!resolveSourceLinkedFormats && axis.GetFirstChild<C.NumberingFormat>()?.SourceLinked?.Value == true) ||
            axis.GetFirstChild<C.NumberingFormat>()?.SourceLinked?.Value != true &&
                OfficeOpenXmlChartSeriesReader.HasUnsupportedSharedAxisNumberFormat(axis) ||
            scaling?.GetFirstChild<C.LogBase>() != null ||
            scaling?.GetFirstChild<C.Orientation>()?.Val?.Value == C.OrientationValues.MaxMin ||
            axis.GetFirstChild<C.DisplayUnits>() != null || axis.GetFirstChild<C.CrossesAt>() != null)
            throw new NotSupportedException("The secondary numeric axis requires an unsupported projection.");
    }

    internal static void QualifyTitle(C.Title? title) {
        var richText = title?.GetFirstChild<C.ChartText>()?.GetFirstChild<C.RichText>();
        var cache = title?.GetFirstChild<C.ChartText>()?.GetFirstChild<C.StringReference>()?.StringCache;
        if (cache != null) {
            C.StringPoint[] points = cache.Elements<C.StringPoint>().ToArray();
            if (points.Length != 1 || string.IsNullOrWhiteSpace(points[0].NumericValue?.Text))
                throw new NotSupportedException("The secondary axis title cache must contain one text value.");
        }
        if (title != null &&
            (title.GetFirstChild<C.Layout>()?.GetFirstChild<C.ManualLayout>() != null ||
             title.GetFirstChild<C.Overlay>() is C.Overlay overlay && overlay.Val?.Value != false ||
             richText?.Elements<A.Paragraph>().Skip(1).Any() == true ||
             richText?.Descendants<A.Break>().Any() == true ||
             ReadTitle(title) == null))
            throw new NotSupportedException("The secondary axis title cannot be projected.");
    }

    internal static void QualifyTypefaceOnlyTitleAppearance(C.PlotArea? plot) {
        var title = Resolve(plot)?.GetFirstChild<C.Title>();
        if (title == null) return;
        if (title.GetFirstChild<C.ChartShapeProperties>() is C.ChartShapeProperties shape &&
            (shape.HasChildren || shape.HasAttributes))
            throw new NotSupportedException("The secondary axis title shape cannot be projected.");
        foreach (A.BodyProperties body in title.Descendants<A.BodyProperties>()) {
            if (body.HasAttributes || body.HasChildren)
                throw new NotSupportedException("The secondary axis title body appearance cannot be projected.");
        }
        foreach (A.ParagraphProperties paragraph in title.Descendants<A.ParagraphProperties>()) {
            if (paragraph.HasAttributes || paragraph.ChildElements.Any(child => child is not A.DefaultRunProperties))
                throw new NotSupportedException("The secondary axis title paragraph appearance cannot be projected.");
        }
        foreach (OpenXmlElement properties in title.Descendants().Where(element =>
            element is A.RunProperties or A.DefaultRunProperties or A.EndParagraphRunProperties)) {
            if (properties.GetAttributes().Any(attribute => attribute.LocalName is not
                    ("lang" or "altLang" or "dirty" or "smtClean" or "smtId") &&
                !(attribute.LocalName == "spc" && attribute.Value == "-1")) ||
                properties.ChildElements.Any(child => child is not A.LatinFont and not A.EastAsianFont and not A.ComplexScriptFont))
                throw new NotSupportedException("The secondary axis title text appearance cannot be projected.");
        }
    }

    private static OfficeChartAxisTickMark ReadTick(C.TickMarkValues? value) =>
        value == C.TickMarkValues.Inside ? OfficeChartAxisTickMark.Inside :
        value == C.TickMarkValues.Outside ? OfficeChartAxisTickMark.Outside :
        value == C.TickMarkValues.Cross ? OfficeChartAxisTickMark.Cross : OfficeChartAxisTickMark.None;

    private static string? ReadTitle(C.Title? title) {
        var text = title?.GetFirstChild<C.ChartText>();
        string? value = text?.GetFirstChild<C.RichText>()?.InnerText ??
            text?.GetFirstChild<C.StringReference>()?.StringCache?.InnerText;
        return string.IsNullOrWhiteSpace(value) ? null : value;
    }

    private static C.TickMarkValues WriteTick(OfficeChartAxisTickMark value) =>
        value == OfficeChartAxisTickMark.Inside ? C.TickMarkValues.Inside :
        value == OfficeChartAxisTickMark.Outside ? C.TickMarkValues.Outside :
        value == OfficeChartAxisTickMark.Cross ? C.TickMarkValues.Cross : C.TickMarkValues.None;
}
