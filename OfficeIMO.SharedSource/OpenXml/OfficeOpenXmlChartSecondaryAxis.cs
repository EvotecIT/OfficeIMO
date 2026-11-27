using DocumentFormat.OpenXml;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using static OfficeIMO.OpenXml.Internal.OfficeChartXml;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal;

/// <summary>Owns numeric settings on the referenced secondary value axis.</summary>
internal static class OfficeOpenXmlChartSecondaryAxis {
    internal static OfficeXmlNode Resolve(OfficeXmlNode plot) {
        if (plot.IsNone) return default;
        var groups = OfficeOpenXmlChartAxisGroups.Create(plot);
        var axes = new List<OfficeXmlNode>();
        foreach (OfficeXmlNode layer in plot.Children()) {
            if (!layer.LocalName.EndsWith("Chart", StringComparison.Ordinal) ||
                groups.Read(layer) != OfficeChartAxisGroup.Secondary) continue;
            foreach (OfficeXmlNode reference in layer.Children(NsC, "axId")) {
                OfficeXmlNode axis = groups.Resolve(reference.UIntVal());
                if (axis.IsChart("valAx") && !axes.Contains(axis)) axes.Add(axis);
            }
        }
        if (axes.Count > 1) throw new NotSupportedException("Multiple independent secondary value axes cannot be projected.");
        return axes.Count == 1 ? axes[0] : default;
    }

    internal static OfficeChartValueAxisLayout? Read(OfficeXmlNode plot, Func<string>? resolveSourceLinkedFormat = null) {
        OfficeXmlNode axis = Resolve(plot);
        if (axis.IsNone) return null;
        OfficeXmlNode scaling = axis.Child(NsC, "scaling");
        var layout = new OfficeChartValueAxisLayout(
            minimum: scaling.Child(NsC, "min").DoubleVal(),
            maximum: scaling.Child(NsC, "max").DoubleVal(),
            majorUnit: axis.Child(NsC, "majorUnit").DoubleVal(),
            minorUnit: axis.Child(NsC, "minorUnit").DoubleVal(),
            numberFormat: IsWorkbookLinked(axis) && resolveSourceLinkedFormat != null
                    ? ResolveLinkedFormat(axis, resolveSourceLinkedFormat)
                    : ReadDrawableNumberFormat(axis, "secondary axis"),
            majorTickMark: ReadTick(axis.ChildVal(NsC, "majorTickMark")),
            minorTickMark: ReadTick(axis.ChildVal(NsC, "minorTickMark")));
        double? logBase = scaling.Child(NsC, "logBase").DoubleVal();
        layout = layout.WithLogarithmicScale(logBase is double value && value >= 2D && value <= 1000D ? value : null)
            .WithReversedScale(scaling.ChildVal(NsC, "orientation") == "maxMin");
        string? title = ReadTitle(axis.Child(NsC, "title"));
        return title == null ? layout : layout.WithTitle(title);
    }

    private static string? ResolveLinkedFormat(OfficeXmlNode axis, Func<string> resolveSourceLinkedFormat) {
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

    internal static void Apply(OfficeXmlNode chart, OfficeChartValueAxisLayout layout) {
        ArgumentNullException.ThrowIfNull(layout);
        OfficeXmlNode axis = Resolve(chart.Child(NsC, "plotArea"));
        if (axis.IsNone) throw new InvalidOperationException("The chart has no referenced secondary value axis.");
        OfficeXmlNode scaling = axis.Child(NsC, "scaling");
        if (scaling.IsNone) scaling = SetOrdered(axis, axis.CreateElement(NsC, "scaling"), CValAxView.Order);
        scaling.RemoveChildren(NsC, "min");
        scaling.RemoveChildren(NsC, "max");
        scaling.RemoveChildren(NsC, "logBase");
        scaling.RemoveChildren(NsC, "orientation");
        if (layout.LogBase.HasValue) SetOrdered(scaling, scaling.CreateVal(NsC, "logBase", Text(layout.LogBase.Value)), CScalingView.Order);
        SetOrdered(scaling, scaling.CreateVal(NsC, "orientation", layout.IsReversed ? "maxMin" : "minMax"), CScalingView.Order);
        if (layout.Minimum.HasValue) SetOrdered(scaling, scaling.CreateVal(NsC, "min", Text(layout.Minimum.Value)), CScalingView.Order);
        if (layout.Maximum.HasValue) SetOrdered(scaling, scaling.CreateVal(NsC, "max", Text(layout.Maximum.Value)), CScalingView.Order);
        axis.RemoveChildren(NsC, "majorUnit");
        axis.RemoveChildren(NsC, "minorUnit");
        if (layout.MajorUnit.HasValue) SetOrdered(axis, axis.CreateVal(NsC, "majorUnit", Text(layout.MajorUnit.Value)), CValAxView.Order);
        if (layout.MinorUnit.HasValue) SetOrdered(axis, axis.CreateVal(NsC, "minorUnit", Text(layout.MinorUnit.Value)), CValAxView.Order);
        axis.RemoveChildren(NsC, "numFmt");
        if (layout.NumberFormat != null) {
            OfficeXmlNode numberFormat = axis.CreateElement(NsC, "numFmt");
            numberFormat.SetAttribute("formatCode", layout.NumberFormat);
            numberFormat.SetAttribute("sourceLinked", Text(false));
            SetOrdered(axis, numberFormat, CValAxView.Order);
        }
        if (layout.MajorTickMark.HasValue) {
            axis.RemoveChildren(NsC, "majorTickMark");
            SetOrdered(axis, axis.CreateVal(NsC, "majorTickMark", WriteTick(layout.MajorTickMark.Value)), CValAxView.Order);
        }
        if (layout.MinorTickMark.HasValue) {
            axis.RemoveChildren(NsC, "minorTickMark");
            SetOrdered(axis, axis.CreateVal(NsC, "minorTickMark", WriteTick(layout.MinorTickMark.Value)), CValAxView.Order);
        }
        if (layout.IsTitleSpecified) {
            OfficeXmlNode existing = axis.Child(NsC, "title");
            if (!existing.IsNone) existing.Remove();
        }
        if (layout.IsTitleSpecified && layout.Title != null) {
            OfficeXmlNode title = axis.CreateElement(NsC, "title");
            OfficeXmlNode rich = title.AppendNew(NsC, "tx").AppendNew(NsC, "rich");
            rich.AppendNew(NsA, "bodyPr");
            rich.AppendNew(NsA, "lstStyle");
            rich.AppendNew(NsA, "p").AppendNew(NsA, "r").AppendNew(NsA, "t").SetInnerText(layout.Title);
            title.AppendNew(NsC, "layout");
            title.AppendVal(NsC, "overlay", Text(false));
            SetOrdered(axis, title, CValAxView.Order);
        }
    }

    internal static void QualifyLinearProjection(OfficeXmlNode plot) {
        OfficeXmlNode axis = Resolve(plot);
        if (axis.IsNone) return;
        QualifyTitle(axis.Child(NsC, "title"));
        OfficeXmlNode secondaryDeletion = axis.Child(NsC, "delete");
        if (!secondaryDeletion.IsNone && secondaryDeletion.BoolVal() != false)
            throw new NotSupportedException("A deleted secondary value axis cannot be projected.");
        var groups = OfficeOpenXmlChartAxisGroups.Create(plot);
        OfficeXmlNode primaryLayer = default;
        foreach (OfficeXmlNode layer in plot.Children()) {
            if (layer.LocalName.EndsWith("Chart", StringComparison.Ordinal) &&
                groups.Read(layer) == OfficeChartAxisGroup.Primary) {
                primaryLayer = layer;
                break;
            }
        }
        var primaryValueAxes = new List<OfficeXmlNode>(2);
        if (!primaryLayer.IsNone) {
            foreach (OfficeXmlNode reference in primaryLayer.Children(NsC, "axId")) {
                OfficeXmlNode candidate = groups.Resolve(reference.UIntVal());
                if (!candidate.IsChart("valAx") || primaryValueAxes.Contains(candidate)) continue;
                primaryValueAxes.Add(candidate);
                if (primaryValueAxes.Count == 2) break;
            }
        }
        if (primaryValueAxes.Count > 1)
            throw new NotSupportedException("Multiple independent primary value axes cannot be projected.");
        OfficeXmlNode primaryValueAxis = primaryValueAxes.Count == 1 ? primaryValueAxes[0] : default;
        OfficeXmlNode deletion = primaryValueAxis.Child(NsC, "delete");
        if ((!deletion.IsNone && deletion.BoolVal() != false) ||
            primaryValueAxis.ChildVal(NsC, "tickLblPos") == "none")
            throw new NotSupportedException("The primary and secondary value-axis visibility cannot be projected independently.");
        // An axis format the renderer cannot reproduce is drawn as General by Read.
        if (!axis.Child(NsC, "dispUnits").IsNone || !axis.Child(NsC, "crossesAt").IsNone)
            throw new NotSupportedException("The secondary numeric axis requires an unsupported projection.");
    }

    internal static bool HasUnsupportedSharedAxisNumberFormat(OfficeXmlNode axis) {
        OfficeXmlNode numbering = axis.Child(NsC, "numFmt");
        // Source-linked formats can differ from formatCode only when a workbook is attached.
        if (IsWorkbookLinked(axis)) return true;
        return !IsDrawableNumberFormat(numbering.Attribute("formatCode"));
    }

    /// <summary>
    /// Returns the number format the shared renderer draws for an axis or label set: its own code when the
    /// renderer reproduces it, otherwise null (General) after recording the fallback.
    /// </summary>
    internal static string? ReadDrawableNumberFormat(OfficeXmlNode owner, string role = "axis") {
        string? format = owner.Child(NsC, "numFmt").Attribute("formatCode");
        // A linked axis shows the plotted cells' format, which the value or category caches record.
        if (IsWorkbookLinked(owner) && owner.IsChart("valAx")) format = ReadCachedValueFormat(owner) ?? format;
        if (IsWorkbookLinked(owner) && (owner.IsChart("catAx") || owner.IsChart("dateAx"))) format = ReadCachedCategoryFormat(owner) ?? format;
        if (!IsDrawableNumberFormat(format)) {
            OfficeOpenXmlChartFidelityScope.Record("The " + role + " number format '" + format + "' is drawn as General.");
            return null;
        }
        // Office keeps a linked code in step with the plotted cells when it saves the chart.
        if (IsWorkbookLinked(owner) && !IsGeneral(format))
            OfficeOpenXmlChartFidelityScope.Record("The source-linked " + role + " number format uses the code cached in the chart, '" + format + "'.");
        return format;
    }

    private static string? ReadCachedCategoryFormat(OfficeXmlNode axis) {
        uint? id = axis.Child(NsC, "axId").UIntVal();
        OfficeXmlNode plot = axis.Parent;
        if (!id.HasValue || !plot.IsChart("plotArea")) return null;
        foreach (OfficeXmlNode layer in plot.Children()) {
            if (!layer.LocalName.EndsWith("Chart", StringComparison.Ordinal) || !ReferencesAxis(layer, id.Value)) continue;
            foreach (OfficeXmlNode series in layer.Children()) {
                if (series.LocalName != "ser") continue;
                OfficeXmlNode source = series.Child(NsC, "cat");
                if (source.IsNone) source = series.Child(NsC, "xVal");
                string? code = source.Child(NsC, "numRef").Child(NsC, "numCache").ChildText(NsC, "formatCode");
                if (!string.IsNullOrWhiteSpace(code)) return code;
            }
        }
        return null;
    }

    private static bool ReferencesAxis(OfficeXmlNode layer, uint id) {
        foreach (OfficeXmlNode reference in layer.Children(NsC, "axId")) {
            if (reference.UIntVal() == id) return true;
        }
        return false;
    }

    private static string? ReadCachedValueFormat(OfficeXmlNode axis) =>
        ReadCachedValueFormats(axis).FirstOrDefault(code => !string.IsNullOrWhiteSpace(code));

    /// <summary>
    /// The cached format codes of the values a value axis plots: the X values for the horizontal axis of a scatter or
    /// bubble group, otherwise the series values.
    /// </summary>
    internal static IEnumerable<string?> ReadCachedValueFormats(OfficeXmlNode axis) {
        uint? id = axis.Child(NsC, "axId").UIntVal();
        OfficeXmlNode plot = axis.Parent;
        if (!id.HasValue || !plot.IsChart("plotArea")) return Enumerable.Empty<string?>();
        string? position = axis.ChildVal(NsC, "axPos");
        bool horizontal = position == "b" || position == "t";
        return ReadCachedValueFormats(plot, id.Value, horizontal);
    }

    private static IEnumerable<string?> ReadCachedValueFormats(OfficeXmlNode plot, uint id, bool horizontal) {
        foreach (OfficeXmlNode layer in plot.Children()) {
            if (!layer.LocalName.EndsWith("Chart", StringComparison.Ordinal) || !ReferencesAxis(layer, id)) continue;
            bool xValues = horizontal && (layer.IsChart("scatterChart") || layer.IsChart("bubbleChart"));
            foreach (OfficeXmlNode series in layer.Children()) {
                if (series.LocalName != "ser") continue;
                OfficeXmlNode values;
                if (xValues) {
                    values = series.Child(NsC, "xVal");
                } else {
                    values = series.Child(NsC, "val");
                    if (values.IsNone) values = series.Child(NsC, "yVal");
                }
                OfficeXmlNode code = values.IsNone ? default : values.FirstDescendant(NsC, "formatCode");
                yield return code.IsNone ? null : code.InnerText;
            }
        }
    }

    internal static bool IsWorkbookLinked(OfficeXmlNode owner) =>
        owner.Child(NsC, "numFmt").BoolAttr("sourceLinked") == true &&
        !owner.Ancestor(NsC, "chartSpace").Child(NsC, "externalData").IsNone;

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

    internal static void QualifyTitle(OfficeXmlNode title) {
        OfficeXmlNode text = title.Child(NsC, "tx");
        OfficeXmlNode richText = text.Child(NsC, "rich");
        OfficeXmlNode cache = text.Child(NsC, "strRef").Child(NsC, "strCache");
        if (!cache.IsNone) {
            var points = new List<OfficeXmlNode>(2);
            foreach (OfficeXmlNode point in cache.Children(NsC, "pt")) {
                points.Add(point);
                if (points.Count == 2) break;
            }
            if (points.Count != 1 || string.IsNullOrWhiteSpace(points[0].ChildText(NsC, "v")))
                throw new NotSupportedException("The secondary axis title cache must contain one text value.");
        }
        OfficeXmlNode overlay = title.Child(NsC, "overlay");
        if (!title.IsNone &&
            (!title.Child(NsC, "layout").Child(NsC, "manualLayout").IsNone ||
             !overlay.IsNone && overlay.BoolVal() != false))
            OfficeOpenXmlChartFidelityScope.Approximate("Manual or overlaid secondary axis title layouts cannot be projected.",
                "A manually placed or overlaid secondary axis title is drawn in its automatic position.");
        if (!title.IsNone &&
            (richText.Children(NsA, "p").Count > 1 ||
             !richText.FirstDescendant(NsA, "br").IsNone ||
             ReadTitle(title) == null))
            throw new NotSupportedException("The secondary axis title cannot be projected.");
    }

    internal static void QualifyTypefaceOnlyTitleAppearance(OfficeXmlNode plot) {
        OfficeXmlNode title = Resolve(plot).Child(NsC, "title");
        if (title.IsNone) return;
        OfficeXmlNode shape = title.Child(NsC, "spPr");
        if (!shape.IsNone && (shape.HasChildElements() || shape.HasAttributes))
            throw new NotSupportedException("The secondary axis title shape cannot be projected.");
        foreach (OfficeXmlNode body in title.Descendants(NsA, "bodyPr")) {
            if (body.HasAttributes || body.HasChildElements())
                throw new NotSupportedException("The secondary axis title body appearance cannot be projected.");
        }
        foreach (OfficeXmlNode paragraph in title.Descendants(NsA, "pPr")) {
            if (paragraph.HasAttributes || paragraph.Children().ToList().Any(child => !child.IsDrawing("defRPr")))
                throw new NotSupportedException("The secondary axis title paragraph appearance cannot be projected.");
        }
        foreach (OfficeXmlNode properties in title.DescendantsWhere(element =>
            element.IsDrawing("rPr") || element.IsDrawing("defRPr") || element.IsDrawing("endParaRPr"))) {
            if (properties.Attributes().Any(attribute => attribute.LocalName is not
                    ("lang" or "altLang" or "dirty" or "smtClean" or "smtId") &&
                !(attribute.LocalName == "spc" && attribute.Value == "-1")) ||
                properties.Children().ToList().Any(child => !child.IsDrawing("latin")))
                throw new NotSupportedException("The secondary axis title text appearance cannot be projected.");
        }
    }

    private static OfficeChartAxisTickMark ReadTick(string? value) =>
        value == "in" ? OfficeChartAxisTickMark.Inside :
        value == "out" ? OfficeChartAxisTickMark.Outside :
        value == "cross" ? OfficeChartAxisTickMark.Cross : OfficeChartAxisTickMark.None;

    private static string? ReadTitle(OfficeXmlNode title) {
        OfficeXmlNode text = title.Child(NsC, "tx");
        OfficeXmlNode rich = text.Child(NsC, "rich");
        OfficeXmlNode cache = text.Child(NsC, "strRef").Child(NsC, "strCache");
        string? value = !rich.IsNone ? rich.InnerText : cache.IsNone ? null : cache.InnerText;
        return string.IsNullOrWhiteSpace(value) ? null : value;
    }

    private static string WriteTick(OfficeChartAxisTickMark value) =>
        value == OfficeChartAxisTickMark.Inside ? "in" :
        value == OfficeChartAxisTickMark.Outside ? "out" :
        value == OfficeChartAxisTickMark.Cross ? "cross" : "none";

    // SDK callers
    internal static C.ValueAxis? Resolve(C.PlotArea? plot) => (C.ValueAxis?)Resolve(OfficeXmlNode.FromSdk(plot)).SdkElement;

    internal static OfficeChartValueAxisLayout? Read(C.PlotArea? plot, Func<string>? resolveSourceLinkedFormat = null) =>
        Read(OfficeXmlNode.FromSdk(plot), resolveSourceLinkedFormat);

    internal static void Apply(C.Chart? chart, OfficeChartValueAxisLayout layout) => Apply(OfficeXmlNode.FromSdk(chart), layout);

    internal static void QualifyLinearProjection(C.PlotArea? plot) => QualifyLinearProjection(OfficeXmlNode.FromSdk(plot));

    internal static void QualifyTypefaceOnlyTitleAppearance(C.PlotArea? plot) => QualifyTypefaceOnlyTitleAppearance(OfficeXmlNode.FromSdk(plot));

    internal static string? ReadDrawableNumberFormat(OpenXmlCompositeElement owner, string role = "axis") =>
        ReadDrawableNumberFormat(OfficeXmlNode.FromSdk(owner), role);
}
