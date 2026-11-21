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

    internal static OfficeChartValueAxisLayout? Read(C.PlotArea? plot, Func<string>? resolveSourceLinkedFormat = null) {
        var axis = Resolve(plot);
        if (axis == null) return null;
        var scaling = axis.GetFirstChild<C.Scaling>();
        var layout = new OfficeChartValueAxisLayout(
            minimum: scaling?.GetFirstChild<C.MinAxisValue>()?.Val?.Value,
            maximum: scaling?.GetFirstChild<C.MaxAxisValue>()?.Val?.Value,
            majorUnit: axis.GetFirstChild<C.MajorUnit>()?.Val?.Value,
            minorUnit: axis.GetFirstChild<C.MinorUnit>()?.Val?.Value,
            numberFormat: IsWorkbookLinked(axis) && resolveSourceLinkedFormat != null
                    ? ResolveLinkedFormat(axis, resolveSourceLinkedFormat)
                    : ReadDrawableNumberFormat(axis, "secondary axis"),
            majorTickMark: ReadTick(axis.GetFirstChild<C.MajorTickMark>()?.Val?.Value),
            minorTickMark: ReadTick(axis.GetFirstChild<C.MinorTickMark>()?.Val?.Value));
        double? logBase = scaling?.GetFirstChild<C.LogBase>()?.Val?.Value;
        layout = layout.WithLogarithmicScale(logBase is double value && value >= 2D && value <= 1000D ? value : null)
            .WithReversedScale(scaling?.GetFirstChild<C.Orientation>()?.Val?.Value == C.OrientationValues.MaxMin);
        string? title = ReadTitle(axis.GetFirstChild<C.Title>());
        return title == null ? layout : layout.WithTitle(title);
    }

    private static string? ResolveLinkedFormat(C.ValueAxis axis, Func<string> resolveSourceLinkedFormat) {
        try {
            return resolveSourceLinkedFormat();
        } catch (NotSupportedException) {
            RecordLinkedFormatFallback("secondary axis");
            return ReadDrawableNumberFormat(axis, "secondary axis");
        }
    }

    /// <summary>Reports that a workbook-linked format could not be read from the workbook, so the chart's cached code is used.</summary>
    internal static void RecordLinkedFormatFallback(string role) {
        OfficeOpenXmlChartFidelityScope.Record("The " + role + " number format linked to the workbook could not be resolved; the format cached in the chart is used.");
    }

    internal static void Apply(C.Chart? chart, OfficeChartValueAxisLayout layout) {
        ArgumentNullException.ThrowIfNull(layout);
        var axis = Resolve(chart?.PlotArea) ?? throw new InvalidOperationException("The chart has no referenced secondary value axis.");
        var scaling = axis.GetFirstChild<C.Scaling>() ?? new C.Scaling();
        if (scaling.Parent == null) axis.AddChild(scaling, true);
        scaling.RemoveAllChildren<C.MinAxisValue>();
        scaling.RemoveAllChildren<C.MaxAxisValue>();
        scaling.RemoveAllChildren<C.LogBase>();
        scaling.RemoveAllChildren<C.Orientation>();
        if (layout.LogBase.HasValue) scaling.AddChild(new C.LogBase { Val = layout.LogBase.Value }, true);
        scaling.AddChild(new C.Orientation { Val = layout.IsReversed ? C.OrientationValues.MaxMin : C.OrientationValues.MinMax }, true);
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

    internal static void QualifyLinearProjection(C.PlotArea? plot) {
        var axis = Resolve(plot);
        if (axis == null) return;
        QualifyTitle(axis.GetFirstChild<C.Title>());
        if (axis.GetFirstChild<C.Delete>() is C.Delete secondaryDeletion && secondaryDeletion.Val?.Value != false)
            throw new NotSupportedException("A deleted secondary value axis cannot be projected.");
        var groups = OfficeOpenXmlChartAxisGroups.Create(plot!);
        var primaryLayer = plot!.ChildElements.OfType<OpenXmlCompositeElement>().FirstOrDefault(layer =>
            layer.LocalName.EndsWith("Chart", StringComparison.Ordinal) &&
            groups.Read(layer) == OfficeChartAxisGroup.Primary);
        C.ValueAxis[] primaryValueAxes = primaryLayer?.Elements<C.AxisId>()
            .Select(reference => groups.Resolve(reference.Val?.Value)).OfType<C.ValueAxis>()
            .Distinct().Take(2).ToArray() ?? Array.Empty<C.ValueAxis>();
        if (primaryValueAxes.Length > 1)
            throw new NotSupportedException("Multiple independent primary value axes cannot be projected.");
        C.ValueAxis? primaryValueAxis = primaryValueAxes.SingleOrDefault();
        if ((primaryValueAxis?.GetFirstChild<C.Delete>() is C.Delete deletion && deletion.Val?.Value != false) ||
            primaryValueAxis?.GetFirstChild<C.TickLabelPosition>()?.Val?.Value == C.TickLabelPositionValues.None)
            throw new NotSupportedException("The primary and secondary value-axis visibility cannot be projected independently.");
        // An axis format the renderer cannot reproduce is drawn as General by Read.
        if (axis.GetFirstChild<C.DisplayUnits>() != null || axis.GetFirstChild<C.CrossesAt>() != null)
            throw new NotSupportedException("The secondary numeric axis requires an unsupported projection.");
    }

    internal static bool HasUnsupportedSharedAxisNumberFormat(OpenXmlCompositeElement axis) {
        C.NumberingFormat? numbering = axis.GetFirstChild<C.NumberingFormat>();
        // Source-linked formats can differ from formatCode only when a workbook is attached.
        if (IsWorkbookLinked(axis)) return true;
        return !IsDrawableNumberFormat(numbering?.FormatCode?.Value);
    }

    /// <summary>
    /// Returns the number format the shared renderer draws for an axis or label set: its own code when the
    /// renderer reproduces it, otherwise null (General) after recording the fallback.
    /// </summary>
    internal static string? ReadDrawableNumberFormat(OpenXmlCompositeElement owner, string role = "axis") {
        string? format = owner.GetFirstChild<C.NumberingFormat>()?.FormatCode?.Value;
        // A linked axis shows the plotted cells' format, which the value or category caches record.
        if (IsWorkbookLinked(owner) && owner is C.ValueAxis axis) format = ReadCachedValueFormat(axis) ?? format;
        if (IsWorkbookLinked(owner) && owner is C.CategoryAxis or C.DateAxis) format = ReadCachedCategoryFormat(owner) ?? format;
        if (!IsDrawableNumberFormat(format)) {
            OfficeOpenXmlChartFidelityScope.Record("The " + role + " number format '" + format + "' is drawn as General.");
            return null;
        }
        // Office keeps a linked code in step with the plotted cells when it saves the chart.
        if (IsWorkbookLinked(owner) && !IsGeneral(format))
            OfficeOpenXmlChartFidelityScope.Record("The source-linked " + role + " number format uses the code cached in the chart, '" + format + "'.");
        return format;
    }

    private static string? ReadCachedCategoryFormat(OpenXmlCompositeElement axis) {
        uint? id = axis.GetFirstChild<C.AxisId>()?.Val?.Value;
        if (!id.HasValue || axis.Parent is not C.PlotArea plot) return null;
        return plot.ChildElements.OfType<OpenXmlCompositeElement>()
            .Where(layer => layer.LocalName.EndsWith("Chart", StringComparison.Ordinal) &&
                layer.Elements<C.AxisId>().Any(reference => reference.Val?.Value == id))
            .SelectMany(layer => layer.ChildElements.Where(series => series.LocalName == "ser"))
            .Select(series => (series.GetFirstChild<C.CategoryAxisData>() ?? (OpenXmlElement?)series.GetFirstChild<C.XValues>())?
                .GetFirstChild<C.NumberReference>()?.NumberingCache?.FormatCode?.Text)
            .FirstOrDefault(code => !string.IsNullOrWhiteSpace(code));
    }

    private static string? ReadCachedValueFormat(C.ValueAxis axis) =>
        ReadCachedValueFormats(axis).FirstOrDefault(code => !string.IsNullOrWhiteSpace(code));

    /// <summary>
    /// The cached format codes of the values a value axis plots: the X values for the horizontal axis of a scatter or
    /// bubble group, otherwise the series values.
    /// </summary>
    internal static IEnumerable<string?> ReadCachedValueFormats(C.ValueAxis axis) {
        uint? id = axis.AxisId?.Val?.Value;
        if (!id.HasValue || axis.Parent is not C.PlotArea plot) return Enumerable.Empty<string?>();
        bool horizontal = axis.AxisPosition?.Val?.Value is C.AxisPositionValues position &&
            (position == C.AxisPositionValues.Bottom || position == C.AxisPositionValues.Top);
        return plot.ChildElements.OfType<OpenXmlCompositeElement>()
            .Where(layer => layer.LocalName.EndsWith("Chart", StringComparison.Ordinal) &&
                layer.Elements<C.AxisId>().Any(reference => reference.Val?.Value == id))
            .SelectMany(layer => layer.ChildElements.Where(series => series.LocalName == "ser")
                .Select(series => horizontal && layer is C.ScatterChart or C.BubbleChart
                    ? series.GetFirstChild<C.XValues>()
                    : series.GetFirstChild<C.Values>() ?? (OpenXmlElement?)series.GetFirstChild<C.YValues>()))
            .Select(values => values?.Descendants<C.FormatCode>().FirstOrDefault()?.Text);
    }

    internal static bool IsWorkbookLinked(OpenXmlCompositeElement owner) =>
        owner.GetFirstChild<C.NumberingFormat>()?.SourceLinked?.Value == true &&
        owner.Ancestors<C.ChartSpace>().FirstOrDefault()?.GetFirstChild<C.ExternalData>() != null;

    private static bool IsGeneral(string? format) =>
        string.IsNullOrWhiteSpace(format) || string.Equals(format, "General", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Whether the spreadsheet cell formatter the chart renderer uses can display values with this format.
    /// </summary>
    internal static bool IsDrawableNumberFormat(string? format) {
        if (IsGeneral(format)) return true;
        if (format!.Length > OfficeChartLayout.MaxNumberFormatLength) return false;
        if (!OfficeIMO.Spreadsheet.SpreadsheetNumberFormatSyntax.Parse(format).IsValid) return false;
        // Sample a value from each numeric section; the text section never formats a chart value, and a
        // date section has no negative serials, which Office shows as ####.
        foreach (double sample in new[] { 45123.25, -45123.25, 0D })
            if (OfficeIMO.Spreadsheet.SpreadsheetNumberFormatDisplay.FormatSerialValue(sample, format!) == null &&
                !IsEmptySection(format!, sample) && !(sample < 0D && IsDateSection(format!, sample))) return false;
        return true;
    }

    private static bool IsDateSection(string format, double sample) =>
        OfficeIMO.Spreadsheet.SpreadsheetNumberFormatDisplay.IsDateTimeSection(
            OfficeIMO.Spreadsheet.SpreadsheetNumberFormatDisplay.SelectNumberFormatSection(format, sample < 0 ? 1 : sample == 0 ? 2 : 0, sample, out _));

    private static bool IsEmptySection(string format, double sample) =>
        OfficeIMO.Spreadsheet.SpreadsheetNumberFormatDisplay.SelectNumberFormatSection(format,
            sample < 0 ? 1 : sample == 0 ? 2 : 0, sample, out _).Trim().Length == 0;

    internal static void QualifyTitle(C.Title? title) {
        var richText = title?.GetFirstChild<C.ChartText>()?.GetFirstChild<C.RichText>();
        var cache = title?.GetFirstChild<C.ChartText>()?.GetFirstChild<C.StringReference>()?.StringCache;
        if (cache != null) {
            C.StringPoint[] points = cache.Elements<C.StringPoint>().Take(2).ToArray();
            if (points.Length != 1 || string.IsNullOrWhiteSpace(points[0].NumericValue?.Text))
                throw new NotSupportedException("The secondary axis title cache must contain one text value.");
        }
        if (title != null &&
            (title.GetFirstChild<C.Layout>()?.GetFirstChild<C.ManualLayout>() != null ||
             title.GetFirstChild<C.Overlay>() is C.Overlay overlay && overlay.Val?.Value != false))
            OfficeOpenXmlChartFidelityScope.Approximate("Manual or overlaid secondary axis title layouts cannot be projected.",
                "A manually placed or overlaid secondary axis title is drawn in its automatic position.");
        if (title != null &&
            (richText?.Elements<A.Paragraph>().Skip(1).Any() == true ||
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
                properties.ChildElements.Any(child => child is not A.LatinFont))
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
