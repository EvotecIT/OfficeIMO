using System;
using System.Linq;
using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal;

internal static partial class OfficeOpenXmlChartSeriesReader {
    private readonly struct NativeText {
        internal NativeText(double? size, OfficeFontStyle? style, OfficeColor? color) { Size = size; Style = style; Color = color; }
        internal double? Size { get; }
        internal OfficeFontStyle? Style { get; }
        internal OfficeColor? Color { get; }
        internal bool SameAs(NativeText other) => Size == other.Size && Style == other.Style && Nullable.Equals(Color, other.Color);
    }

    private static NativeText ReadNativeText(C.Chart chart, OpenXmlElement? owner, A.ColorScheme? scheme) {
        if (owner == null) return default;
        NativeText inherited = ReadTextPropertyDefaults(default, chart.Parent?.GetFirstChild<C.TextProperties>(), scheme);
        inherited = ReadTextPropertyDefaults(inherited, owner.GetFirstChild<C.TextProperties>(), scheme);
        var rich = owner.GetFirstChild<C.ChartText>()?.GetFirstChild<C.RichText>();
        if (rich == null) return inherited;
        QualifyTextLayout(rich);
        NativeText? selected = null;
        foreach (var paragraph in rich.Elements<A.Paragraph>()) {
            int level = paragraph.ParagraphProperties?.Level?.Value ?? 0;
            var list = rich.GetFirstChild<A.ListStyle>()?.ChildElements.FirstOrDefault(element => element.LocalName == $"lvl{level + 1}pPr");
            NativeText paragraphText = ApplyNativeText(inherited, list?.GetFirstChild<A.DefaultRunProperties>(), scheme);
            paragraphText = ApplyNativeText(paragraphText, paragraph.ParagraphProperties?.GetFirstChild<A.DefaultRunProperties>(), scheme);
            foreach (var run in paragraph.ChildElements.Where(element => element is A.Run or A.Field)) {
                if (string.IsNullOrEmpty(run.GetFirstChild<A.Text>()?.Text)) continue;
                NativeText current = ApplyNativeText(paragraphText, run.GetFirstChild<A.RunProperties>(), scheme);
                if (selected.HasValue && !selected.Value.SameAs(current))
                    throw new NotSupportedException("Mixed chart text formatting cannot be projected.");
                selected = current;
            }
        }
        return selected ?? inherited;
    }

    private static NativeText ReadTextPropertyDefaults(NativeText inherited, C.TextProperties? properties, A.ColorScheme? scheme) {
        if (properties == null) return inherited;
        QualifyTextLayout(properties);
        NativeText? selected = null;
        foreach (var paragraph in properties.Elements<A.Paragraph>()) {
            int level = paragraph.ParagraphProperties?.Level?.Value ?? 0;
            var list = properties.GetFirstChild<A.ListStyle>()?.ChildElements.FirstOrDefault(element => element.LocalName == $"lvl{level + 1}pPr");
            NativeText current = ApplyNativeText(inherited, list?.GetFirstChild<A.DefaultRunProperties>(), scheme);
            current = ApplyNativeText(current, paragraph.ParagraphProperties?.GetFirstChild<A.DefaultRunProperties>(), scheme);
            var explicitRuns = paragraph.Elements<A.Run>().Where(run => !string.IsNullOrEmpty(run.GetFirstChild<A.Text>()?.Text)).ToArray();
            if (explicitRuns.Length == 0) current = ApplyNativeText(current, paragraph.GetFirstChild<A.EndParagraphRunProperties>(), scheme);
            foreach (var run in explicitRuns) {
                NativeText runText = ApplyNativeText(current, run.GetFirstChild<A.RunProperties>(), scheme);
                if (selected.HasValue && !selected.Value.SameAs(runText)) throw new NotSupportedException("Mixed chart text defaults cannot be projected.");
                selected = runText;
            }
            if (explicitRuns.Length == 0) {
                if (selected.HasValue && !selected.Value.SameAs(current)) throw new NotSupportedException("Mixed chart text defaults cannot be projected.");
                selected = current;
            }
        }
        return selected ?? inherited;
    }

    private static NativeText ReadUniformNativeText(C.Chart chart, System.Collections.Generic.IEnumerable<OpenXmlElement> owners, A.ColorScheme? scheme) {
        NativeText? selected = null;
        foreach (var owner in owners) {
            NativeText current = ReadNativeText(chart, owner, scheme);
            if (selected.HasValue && !selected.Value.SameAs(current))
                throw new NotSupportedException("Independent chart text formatting cannot be projected by the shared text role.");
            selected = current;
        }
        return selected ?? default;
    }

    private static void QualifyTextLayout(OpenXmlCompositeElement text) {
        if (text is C.RichText && (text.Elements<A.Paragraph>().Skip(1).Any() || text.Descendants<A.Break>().Any()))
            throw new NotSupportedException("Multiline native chart text cannot be projected.");
        var body = text.GetFirstChild<A.BodyProperties>();
        if (body != null && (body.GetAttributes().Any(attribute =>
                attribute.LocalName != "rot" || attribute.Value != "0") ||
                body.ChildElements.Any(child => child is not A.NoAutoFit)))
            throw new NotSupportedException("The native chart text body layout cannot be projected.");
        foreach (var paragraphProperties in text.Descendants().Where(element => element is A.ParagraphProperties ||
            element.Parent is A.ListStyle)) {
            if (paragraphProperties.GetAttributes().Any(attribute => attribute.LocalName != "lvl") ||
                paragraphProperties.ChildElements.Any(child => child is not A.DefaultRunProperties))
                throw new NotSupportedException("The native chart paragraph layout cannot be projected.");
        }
    }

    private static NativeText ApplyNativeText(NativeText inherited, OpenXmlElement? properties, A.ColorScheme? scheme) {
        if (properties == null) return inherited;
        double? size = inherited.Size;
        OfficeFontStyle? style = inherited.Style;
        OfficeColor? color = inherited.Color;
        string? latinTypeface = properties.GetFirstChild<A.LatinFont>()?.Typeface?.Value;
        foreach (var attribute in properties.GetAttributes()) {
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
            } else if (attribute.LocalName is not "lang" and not "altLang" and not "dirty" and not "smtClean" and not "smtId") {
                throw new NotSupportedException("The chart text attributes cannot be projected.");
            }
        }
        foreach (var child in properties.ChildElements) {
            if (child is A.SolidFill fill) {
                if (OfficeOpenXmlThemeColorResolver.HasUnsupportedTransforms(fill))
                    throw new NotSupportedException("The chart text colour transforms cannot be projected.");
                color = OfficeOpenXmlThemeColorResolver.ResolveColor(fill, scheme) ?? throw new NotSupportedException("The chart text colour cannot be resolved.");
            } else if (child is A.EastAsianFont or A.ComplexScriptFont) {
                string? scriptTypeface = child is A.EastAsianFont eastAsian ? eastAsian.Typeface?.Value : ((A.ComplexScriptFont)child).Typeface?.Value;
                if (string.IsNullOrWhiteSpace(latinTypeface) || !string.Equals(latinTypeface, scriptTypeface, StringComparison.OrdinalIgnoreCase))
                    throw new NotSupportedException("Different chart text script typefaces cannot be projected by one font family.");
            } else if (child is not A.LatinFont and not A.EastAsianFont and not A.ComplexScriptFont)
                throw new NotSupportedException("The chart text appearance cannot be projected.");
        }
        return new NativeText(size, style, color);
    }

    private static System.Collections.Generic.IEnumerable<OpenXmlElement> TextAxes(C.Chart chart) =>
        chart.PlotArea?.ChildElements.Where(axis => axis is C.CategoryAxis or C.ValueAxis or C.DateAxis).Cast<OpenXmlElement>() ?? Enumerable.Empty<OpenXmlElement>();
}
