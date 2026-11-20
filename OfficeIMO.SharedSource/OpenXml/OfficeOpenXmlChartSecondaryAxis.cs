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
        string? title = ReadTitle(axis.GetFirstChild<C.Title>());
        return title == null ? layout : layout.WithTitle(title);
    }

    private static string? ResolveLinkedFormat(C.ValueAxis axis, Func<string> resolveSourceLinkedFormat) {
        try {
            return resolveSourceLinkedFormat();
        } catch (NotSupportedException) {
            return ReadDrawableNumberFormat(axis, "secondary axis");
        }
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
        var scaling = axis.GetFirstChild<C.Scaling>();
        // An axis format the renderer cannot reproduce is drawn as General by Read.
        if (scaling?.GetFirstChild<C.LogBase>() != null ||
            scaling?.GetFirstChild<C.Orientation>()?.Val?.Value == C.OrientationValues.MaxMin ||
            axis.GetFirstChild<C.DisplayUnits>() != null || axis.GetFirstChild<C.CrossesAt>() != null)
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
        // A linked value axis shows the plotted cells' format, which the value caches record.
        if (IsWorkbookLinked(owner) && owner is C.ValueAxis axis) format = ReadCachedValueFormat(axis) ?? format;
        // Category axes draw numeric categories as spreadsheet serial dates.
        if (owner is C.CategoryAxis or C.DateAxis && IsDateFormat(format)) return format;
        if (!IsDrawableNumberFormat(format)) {
            OfficeOpenXmlChartFidelityScope.Record("The " + role + " number format '" + format + "' is drawn as General.");
            return null;
        }
        // Office keeps a linked code in step with the plotted cells when it saves the chart.
        if (IsWorkbookLinked(owner) && !IsGeneral(format))
            OfficeOpenXmlChartFidelityScope.Record("The source-linked " + role + " number format uses the code cached in the chart, '" + format + "'.");
        return format;
    }

    private static string? ReadCachedValueFormat(C.ValueAxis axis) {
        uint? id = axis.AxisId?.Val?.Value;
        if (!id.HasValue || axis.Parent is not C.PlotArea plot) return null;
        return plot.ChildElements.OfType<OpenXmlCompositeElement>()
            .Where(layer => layer.LocalName.EndsWith("Chart", StringComparison.Ordinal) &&
                layer.Elements<C.AxisId>().Any(reference => reference.Val?.Value == id))
            .SelectMany(layer => layer.ChildElements.Where(series => series.LocalName == "ser"))
            .Select(series => (series.GetFirstChild<C.Values>() ?? (OpenXmlElement?)series.GetFirstChild<C.YValues>())?
                .Descendants<C.FormatCode>().FirstOrDefault()?.Text)
            .FirstOrDefault(code => !string.IsNullOrWhiteSpace(code));
    }

    private static bool IsWorkbookLinked(OpenXmlCompositeElement owner) =>
        owner.GetFirstChild<C.NumberingFormat>()?.SourceLinked?.Value == true &&
        owner.Ancestors<C.ChartSpace>().FirstOrDefault()?.GetFirstChild<C.ExternalData>() != null;

    private static bool IsGeneral(string? format) =>
        string.IsNullOrWhiteSpace(format) || string.Equals(format, "General", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Whether the shared chart renderer reproduces a spreadsheet number format: digit placeholders,
    /// grouping and scaling commas, percent, literals, up to four sections (a zero section may be
    /// literal-only, text only in the fourth), colour tags and <c>[$symbol-locale]</c> currency tags.
    /// Conditions, dates, fractions and exponents are not reproduced.
    /// </summary>
    internal static bool IsDrawableNumberFormat(string? format) {
        if (IsGeneral(format)) return true;

        bool inQuotedLiteral = false;
        bool escaped = false;
        bool sectionHasPlaceholder = false;
        int section = 0;
        for (int index = 0; index < format!.Length; index++) {
            char value = format[index];
            if (escaped) {
                if (value == '%') return false;
                escaped = false;
                continue;
            }
            if (value == '\\') { escaped = true; continue; }
            if (value == '"') { inQuotedLiteral = !inQuotedLiteral; continue; }
            if (inQuotedLiteral) {
                if (value == '%') return false;
                continue;
            }
            if (value == '_' || value == '*') {
                // Padding and fill take the next character, which never formats the value.
                index++;
                continue;
            }
            if (value == '0' || value == '#' || value == '?') {
                if (section == 3) return false;
                sectionHasPlaceholder = true;
                continue;
            }
            if (value == ';') {
                if (section == 0 && !sectionHasPlaceholder) return false;
                if (++section > 3) return false;
                sectionHasPlaceholder = false;
                continue;
            }
            if (value == '[') {
                int close = format.IndexOf(']', index + 1);
                if (close < 0 || !IsDrawableBracket(format.Substring(index + 1, close - index - 1))) return false;
                index = close;
                continue;
            }
            if (value == '@') {
                if (section != 3) return false;
                continue;
            }
            if (value == '/' || value == ']') return false;
            if (value != 'E' && value != 'e') continue;
            int next = index + 1;
            if (next < format.Length && (format[next] == '+' || format[next] == '-')) next++;
            if (next < format.Length && (format[next] == '0' || format[next] == '#' || format[next] == '?'))
                return false;
        }

        return !inQuotedLiteral && !escaped && (section > 0 || sectionHasPlaceholder);
    }

    private static bool IsDateFormat(string? format) {
        if (IsGeneral(format)) return false;
        string section = format!.Split(';')[0];
        bool inQuotedLiteral = false, hasDate = false;
        for (int index = 0; index < section.Length; index++) {
            char value = section[index];
            if (value == '"') { inQuotedLiteral = !inQuotedLiteral; continue; }
            if (inQuotedLiteral) continue;
            if (value == '\\') { index++; continue; }
            if (value == '[') {
                int close = section.IndexOf(']', index + 1);
                if (close < 0) return false;
                index = close;
                continue;
            }
            if (value is '0' or '#' or '?' or '%' or '@') return false;
            if (char.ToLowerInvariant(value) is 'y' or 'm' or 'd' or 'h' or 's') hasDate = true;
        }
        return hasDate && !inQuotedLiteral;
    }

    private static readonly string[] FormatColorNames = { "Black", "Blue", "Cyan", "Green", "Magenta", "Red", "White", "Yellow" };

    private static bool IsDrawableBracket(string content) {
        // [$€-413] prints its currency symbol; [$-409] only names a locale.
        if (content.StartsWith("$", StringComparison.Ordinal)) return true;
        if (FormatColorNames.Any(name => string.Equals(content, name, StringComparison.OrdinalIgnoreCase))) return true;
        return content.Length > 5 && content.StartsWith("Color", StringComparison.OrdinalIgnoreCase) &&
            int.TryParse(content.Substring(5), System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out int index) &&
            index >= 1 && index <= 56;
    }

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
