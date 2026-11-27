using System.Runtime.CompilerServices;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using static OfficeIMO.OpenXml.Internal.OfficeChartXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal;

internal sealed partial class OfficeOpenXmlChartTextReader {
    internal readonly struct NativeText {
        internal NativeText(double? size, OfficeFontStyle? style, OfficeColor? color, string? family) { Size = size; Style = style; Color = color; Family = family; }
        internal double? Size { get; }
        internal OfficeFontStyle? Style { get; }
        internal OfficeColor? Color { get; }
        internal string? Family { get; }
        internal bool SameAs(NativeText other) => Size == other.Size && Style == other.Style && Nullable.Equals(Color, other.Color) &&
            string.Equals(Family, other.Family, StringComparison.OrdinalIgnoreCase);
    }

    // Chart parts of OfficeIMO's own trees, registered by the host that holds them.
    private static readonly ConditionalWeakTable<OfficeXmlTree, IOfficeChartPart> TreeChartParts = new();

    /// <summary>
    /// Records the chart part a tree holds, so text read from a chart element in that tree resolves its color map and
    /// chart style the way an SDK chart part's does.
    /// </summary>
    internal static void AttachChartPart(OfficeXmlTree tree, IOfficeChartPart part) {
        TreeChartParts.Remove(tree);
        TreeChartParts.Add(tree, part);
    }

    /// <summary>
    /// The chart part whose <c>c:chartSpace</c> root holds <paramref name="chart"/>, or null when the chart's parent is
    /// not a part's chart space root.
    /// </summary>
    internal static IOfficeChartPart? ChartPartOf(OfficeXmlNode chart) => ChartPartOfSpace(chart.Parent);

    /// <summary>The chart part whose root is <paramref name="space"/>, or null when it is not a part's chart space root.</summary>
    internal static IOfficeChartPart? ChartPartOfSpace(OfficeXmlNode space) {
        if (!space.Is(NsC, "chartSpace") || !space.Parent.IsNone) return null;
        if (space.SdkElement is { } element)
            return OfficeSdkXml.PartOf(element) is DocumentFormat.OpenXml.Packaging.ChartPart part ? new OfficeSdkChartPart(part) : null;
        if (space.Tree is { } tree && space.Node == tree.Root && TreeChartParts.TryGetValue(tree, out IOfficeChartPart? held)) return held;
        return null;
    }

    internal static NativeText ReadNativeText(OfficeXmlNode chart, OfficeXmlNode owner, OfficeXmlNode scheme, bool allowMultiline = false) {
        if (owner.IsNone) return default;
        IOfficeChartPart? chartPart = ChartPartOf(chart);
        OfficeXmlNode colorMap = chartPart != null
            ? OfficeOpenXmlThemeColorResolver.ResolveChartColorMap(chartPart) : default;
        NativeText inherited = ReadTextPropertyDefaults(default, chart.Parent.Child(NsC, "txPr"), scheme, colorMap);
        return ReadNativeText(owner, scheme, colorMap, inherited, allowMultiline);
    }

    private static NativeText ReadNativeText(OfficeXmlNode owner, OfficeXmlNode scheme,
        OfficeXmlNode colorMap, NativeText inherited, bool allowMultiline = false) {
        inherited = ReadTextPropertyDefaults(inherited, owner.Child(NsC, "txPr"), scheme, colorMap);
        OfficeXmlNode rich = owner.Child(NsC, "tx").Child(NsC, "rich");
        if (rich.IsNone) return inherited;
        QualifyTextLayout(rich, allowMultiline);
        Dictionary<string, OfficeXmlNode> listLevels = ReadListLevels(rich);
        NativeText? selected = null;
        foreach (OfficeXmlNode paragraph in rich.Children(NsA, "p")) {
            OfficeXmlNode paragraphProperties = paragraph.Child(NsA, "pPr");
            int level = paragraphProperties.IntAttr("lvl") ?? 0;
            listLevels.TryGetValue($"lvl{level + 1}pPr", out OfficeXmlNode list);
            NativeText paragraphText = ApplyNativeText(inherited, list.Child(NsA, "defRPr"), scheme, colorMap);
            paragraphText = ApplyNativeText(paragraphText, paragraphProperties.Child(NsA, "defRPr"), scheme, colorMap);
            foreach (OfficeXmlNode run in paragraph.Children()) {
                if (!(run.NamespaceId == NsA && run.LocalName is "r" or "fld")) continue;
                if (string.IsNullOrEmpty(run.Child(NsA, "t").InnerText)) continue;
                NativeText current = ApplyNativeText(paragraphText, run.Child(NsA, "rPr"), scheme, colorMap);
                if (selected.HasValue && !selected.Value.SameAs(current))
                    throw new NotSupportedException("Mixed chart text formatting cannot be projected.");
                selected = current;
            }
        }
        return selected ?? inherited;
    }

    internal static NativeText ReadTextPropertyDefaults(NativeText inherited, OfficeXmlNode properties, OfficeXmlNode scheme, OfficeXmlNode colorMap) {
        if (properties.IsNone) return inherited;
        QualifyTextLayout(properties);
        Dictionary<string, OfficeXmlNode> listLevels = ReadListLevels(properties);
        NativeText? selected = null;
        foreach (OfficeXmlNode paragraph in properties.Children(NsA, "p")) {
            OfficeXmlNode paragraphProperties = paragraph.Child(NsA, "pPr");
            int level = paragraphProperties.IntAttr("lvl") ?? 0;
            listLevels.TryGetValue($"lvl{level + 1}pPr", out OfficeXmlNode list);
            NativeText current = ApplyNativeText(inherited, list.Child(NsA, "defRPr"), scheme, colorMap);
            current = ApplyNativeText(current, paragraphProperties.Child(NsA, "defRPr"), scheme, colorMap);
            var explicitRuns = new List<OfficeXmlNode>();
            foreach (OfficeXmlNode run in paragraph.Children(NsA, "r"))
                if (!string.IsNullOrEmpty(run.Child(NsA, "t").InnerText)) explicitRuns.Add(run);
            if (explicitRuns.Count == 0) current = ApplyNativeText(current, paragraph.Child(NsA, "endParaRPr"), scheme, colorMap);
            foreach (OfficeXmlNode run in explicitRuns) {
                NativeText runText = ApplyNativeText(current, run.Child(NsA, "rPr"), scheme, colorMap);
                if (selected.HasValue && !selected.Value.SameAs(runText)) throw new NotSupportedException("Mixed chart text defaults cannot be projected.");
                selected = runText;
            }
            if (explicitRuns.Count == 0) {
                if (selected.HasValue && !selected.Value.SameAs(current)) throw new NotSupportedException("Mixed chart text defaults cannot be projected.");
                selected = current;
            }
        }
        return selected ?? inherited;
    }

    private static Dictionary<string, OfficeXmlNode> ReadListLevels(OfficeXmlNode text) {
        var levels = new Dictionary<string, OfficeXmlNode>(StringComparer.Ordinal);
        foreach (OfficeXmlNode child in text.Child(NsA, "lstStyle").Children()) {
            if (levels.ContainsKey(child.LocalName))
                throw new NotSupportedException("Duplicate chart text list levels cannot be projected.");
            levels.Add(child.LocalName, child);
        }
        return levels;
    }

    internal static NativeText ReadUniformNativeText(OfficeXmlNode chart, IEnumerable<OfficeXmlNode> owners, OfficeXmlNode scheme, bool allowMultiline = false) {
        IOfficeChartPart? chartPart = ChartPartOf(chart);
        OfficeXmlNode colorMap = chartPart != null
            ? OfficeOpenXmlThemeColorResolver.ResolveChartColorMap(chartPart) : default;
        NativeText inherited = ReadTextPropertyDefaults(default, chart.Parent.Child(NsC, "txPr"), scheme, colorMap);
        NativeText? selected = null;
        foreach (OfficeXmlNode owner in owners) {
            NativeText current = ReadNativeText(owner, scheme, colorMap, inherited, allowMultiline);
            if (selected.HasValue && !selected.Value.SameAs(current))
                throw new NotSupportedException("Independent chart text formatting cannot be projected by the shared text role.");
            selected = current;
        }
        return selected ?? default;
    }

    private static void QualifyTextLayout(OfficeXmlNode text, bool allowMultiline = false) {
        if (!allowMultiline && text.Is(NsC, "rich") && (text.Children(NsA, "p").Count > 1 || !text.FirstDescendant(NsA, "br").IsNone))
            throw new NotSupportedException("Multiline native chart text cannot be projected.");
        OfficeXmlNode body = text.Child(NsA, "bodyPr");
        if (!body.IsNone && (body.Attributes().Any(attribute => !IsNeutralTextBodyAttribute(attribute, text)) ||
                body.Children().ToList().Any(child => !(child.NamespaceId == NsA && child.LocalName is "noAutofit" or "spAutoFit"))))
            throw new NotSupportedException("The native chart text body layout cannot be projected.");
        foreach (OfficeXmlNode paragraphProperties in text.DescendantsWhere(element => element.Is(NsA, "pPr") ||
            element.Parent.Is(NsA, "lstStyle"))) {
            if (paragraphProperties.Attributes().Any(attribute => attribute.LocalName != "lvl") ||
                paragraphProperties.Children().ToList().Any(child => !child.Is(NsA, "defRPr")))
                throw new NotSupportedException("The native chart paragraph layout cannot be projected.");
        }
    }

    // Office writes these body settings on every chart title and label. For single-line,
    // centred, unrotated text they leave the shared layout unchanged. Office writes the
    // -60000000 sentinel for automatic axis-label rotation and a quarter turn on the title of a
    // vertical axis, which is how the shared renderer draws both.
    private static bool IsNeutralTextBodyAttribute(OfficeXmlAttribute attribute, OfficeXmlNode text) => attribute.LocalName switch {
        "rot" => attribute.Value is "0" or "-60000000" || attribute.Value == "-5400000" && IsVerticalAxisTitle(text),
        "vert" => attribute.Value == "horz",
        "wrap" => attribute.Value == "square" || attribute.Value == "none",
        "anchor" => attribute.Value == "ctr",
        "anchorCtr" or "spcFirstLastPara" or "vertOverflow" or "horzOverflow" => true,
        "lIns" or "tIns" or "rIns" or "bIns" => int.TryParse(attribute.Value, System.Globalization.NumberStyles.Integer,
            System.Globalization.CultureInfo.InvariantCulture, out int inset) && inset >= 0 && inset <= 91440,
        _ => false
    };

    private static bool IsVerticalAxisTitle(OfficeXmlNode text) {
        OfficeXmlNode axis = text.Ancestor(NsC, "title").Parent;
        string? position = axis.ChildVal(NsC, "axPos");
        return axis.NamespaceId == NsC && axis.LocalName is "valAx" or "catAx" or "dateAx" &&
            (position == "l" || position == "r");
    }

    private static NativeText ApplyNativeText(NativeText inherited, OfficeXmlNode properties, OfficeXmlNode scheme, OfficeXmlNode colorMap) {
        if (properties.IsNone) return inherited;
        double? size = inherited.Size;
        OfficeFontStyle? style = inherited.Style;
        OfficeColor? color = inherited.Color;
        string? family = inherited.Family;
        string? latinTypeface = properties.Child(NsA, "latin").Attribute("typeface");
        if (latinTypeface != null) {
            family = ResolveNativeTypeface(latinTypeface, scheme);
        }
        foreach (OfficeXmlAttribute attribute in properties.Attributes()) {
            if (attribute.LocalName == "sz") {
                if (!int.TryParse(attribute.Value, out int hundredths) || hundredths <= 0)
                    throw new NotSupportedException("The chart text size cannot be projected.");
                size = hundredths / 100d;
            } else if (attribute.LocalName is "b" or "i") {
                OfficeFontStyle flag = attribute.LocalName == "b" ? OfficeFontStyle.Bold : OfficeFontStyle.Italic;
                style = attribute.Value is "1" or "true" ? (style ?? OfficeFontStyle.Regular) | flag : (style ?? OfficeFontStyle.Regular) & ~flag;
            } else if (attribute.LocalName == "u") {
                if (attribute.Value is not "none" and not "sng") throw new NotSupportedException("The chart text underline cannot be projected.");
                style = attribute.Value == "sng" ? (style ?? OfficeFontStyle.Regular) | OfficeFontStyle.Underline : (style ?? OfficeFontStyle.Regular) & ~OfficeFontStyle.Underline;
            } else if (attribute.LocalName == "strike") {
                if (attribute.Value is not "noStrike" and not "sngStrike") throw new NotSupportedException("The chart text strike cannot be projected.");
                style = attribute.Value == "sngStrike" ? (style ?? OfficeFontStyle.Regular) | OfficeFontStyle.Strikethrough : (style ?? OfficeFontStyle.Regular) & ~OfficeFontStyle.Strikethrough;
            } else if (attribute.LocalName == "spc" && attribute.Value is "-1" or "0") {
                // Some producers serialize this visually negligible default on every chart run.
            } else if (attribute.LocalName == "baseline" && attribute.Value == "0" ||
                attribute.LocalName == "kern" && int.TryParse(attribute.Value, out int kerning) && kerning >= 0) {
                // Office writes no superscript and its kerning threshold on every chart run.
            } else if (attribute.LocalName is not "lang" and not "altLang" and not "dirty" and not "smtClean" and not "smtId") {
                throw new NotSupportedException("The chart text attributes cannot be projected.");
            }
        }
        foreach (OfficeXmlNode child in properties.Children()) {
            bool drawing = child.NamespaceId == NsA;
            if (drawing && child.LocalName == "solidFill") {
                if (OfficeOpenXmlThemeColorResolver.HasUnsupportedTransforms(child))
                    throw new NotSupportedException("The chart text colour transforms cannot be projected.");
                color = OfficeOpenXmlThemeColorResolver.ResolveColor(child, scheme, colorMap: colorMap) ?? throw new NotSupportedException("The chart text colour cannot be resolved.");
            } else if (drawing && child.LocalName is "ea" or "cs") {
                string? scriptTypeface = child.Attribute("typeface");
                // Office pairs +mn-lt with +mn-ea and +mn-cs (and +mj-lt with +mj-ea and +mj-cs). The
                // script slots only apply to East Asian and complex-script runs of the same theme group.
                if (IsSameThemeFontGroup(latinTypeface, scriptTypeface)) continue;
                if (string.IsNullOrWhiteSpace(latinTypeface) ||
                    !string.Equals(family, ResolveNativeTypeface(scriptTypeface, scheme), StringComparison.OrdinalIgnoreCase))
                    throw new NotSupportedException("Different chart text script typefaces cannot be projected by one font family.");
            } else if (!(drawing && child.LocalName is "latin" or "ea" or "cs"))
                throw new NotSupportedException("The chart text appearance cannot be projected.");
        }
        return new NativeText(size, style, color, family);
    }

    private static bool IsSameThemeFontGroup(string? latinTypeface, string? scriptTypeface) =>
        latinTypeface is "+mn-lt" && scriptTypeface is "+mn-ea" or "+mn-cs" ||
        latinTypeface is "+mj-lt" && scriptTypeface is "+mj-ea" or "+mj-cs";

    private static string ResolveNativeTypeface(string? typeface, OfficeXmlNode scheme) {
        if (string.IsNullOrWhiteSpace(typeface))
            throw new NotSupportedException("The chart text typeface cannot be projected.");
        if (!typeface!.StartsWith("+", StringComparison.Ordinal)) return typeface;
        OfficeXmlNode fonts = scheme.Parent.Child(NsA, "fontScheme");
        string? resolved = typeface switch {
            "+mj-lt" => SchemeTypeface(fonts, "majorFont"),
            "+mn-lt" => SchemeTypeface(fonts, "minorFont"),
            _ => null
        };
        return !string.IsNullOrWhiteSpace(resolved) ? resolved! :
            throw new NotSupportedException("The chart text typeface cannot be projected without a resolved font scheme.");
    }

    // SDK callers
    internal static NativeText ReadNativeText(C.Chart chart, DocumentFormat.OpenXml.OpenXmlElement? owner, A.ColorScheme? scheme, bool allowMultiline = false) =>
        ReadNativeText(OfficeXmlNode.FromSdk(chart), OfficeXmlNode.FromSdk(owner), OfficeXmlNode.FromSdk(scheme), allowMultiline);
    internal static NativeText ReadUniformNativeText(C.Chart chart, IEnumerable<DocumentFormat.OpenXml.OpenXmlElement> owners, A.ColorScheme? scheme, bool allowMultiline = false) =>
        ReadUniformNativeText(OfficeXmlNode.FromSdk(chart), owners.Select(owner => OfficeXmlNode.FromSdk(owner)), OfficeXmlNode.FromSdk(scheme), allowMultiline);
}
