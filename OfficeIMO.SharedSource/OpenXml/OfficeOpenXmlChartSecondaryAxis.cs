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
        return new OfficeChartValueAxisLayout(
            minimum: scaling?.GetFirstChild<C.MinAxisValue>()?.Val?.Value,
            maximum: scaling?.GetFirstChild<C.MaxAxisValue>()?.Val?.Value,
            majorUnit: axis.GetFirstChild<C.MajorUnit>()?.Val?.Value,
            minorUnit: axis.GetFirstChild<C.MinorUnit>()?.Val?.Value,
            numberFormat: axis.GetFirstChild<C.NumberingFormat>()?.FormatCode?.Value,
            majorTickMark: ReadTick(axis.GetFirstChild<C.MajorTickMark>()?.Val?.Value),
            minorTickMark: ReadTick(axis.GetFirstChild<C.MinorTickMark>()?.Val?.Value))
            .WithTitle(ReadTitle(axis.GetFirstChild<C.Title>()));
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
        axis.GetFirstChild<C.Title>()?.Remove();
        if (layout.Title != null) axis.AddChild(new C.Title(
            new C.ChartText(new C.RichText(new A.BodyProperties(), new A.ListStyle(),
                new A.Paragraph(new A.Run(new A.Text(layout.Title))))),
            new C.Layout(), new C.Overlay { Val = false }), true);
    }

    internal static void QualifyLinearProjection(C.PlotArea? plot, bool resolveSourceLinkedFormats = false) {
        var axis = Resolve(plot);
        if (axis == null) return;
        QualifyTitle(axis.GetFirstChild<C.Title>());
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
        if (title != null &&
            (title.GetFirstChild<C.Layout>()?.GetFirstChild<C.ManualLayout>() != null ||
             title.GetFirstChild<C.Overlay>()?.Val?.Value == true ||
             title.GetFirstChild<C.ChartText>() == null))
            throw new NotSupportedException("The secondary axis title cannot be projected.");
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
