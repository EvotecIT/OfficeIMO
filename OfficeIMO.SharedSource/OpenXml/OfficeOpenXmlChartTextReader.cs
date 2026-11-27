using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using static OfficeIMO.OpenXml.Internal.OfficeChartXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal;

/// <summary>Resolves chart typefaces through the native text inheritance and owner theme.</summary>
internal sealed partial class OfficeOpenXmlChartTextReader {
    private readonly OfficeXmlNode _scheme;
    internal OfficeOpenXmlChartTextReader(OfficeXmlNode scheme) => _scheme = scheme;

    internal static bool HasUncachedFormulaTitle(OfficeXmlNode chart) =>
        chart.Descendants(NsC, "title").Any(title => title.Child(NsC, "tx").Child(NsC, "strRef") is { IsNone: false } reference &&
            !reference.Child(NsC, "f").IsNone && string.IsNullOrWhiteSpace(reference.Child(NsC, "strCache").InnerText));

        internal bool TryReadSharedTextStyle(OfficeXmlNode chart,
            out OfficeChartStyle? style) {
            string? chartDefaultTypeface = ReadChartDefaultTypeface(chart);
            List<OfficeXmlNode> bodyTextAreas = chart.DescendantsWhere(IsRelevantBodyTextArea);
            string?[] bodyFonts = bodyTextAreas
                .Select(textArea => ReadBodyTypeface(textArea,
                    chartDefaultTypeface))
                .ToArray();
            if (bodyFonts.Length > 0
                && (bodyFonts.Any(string.IsNullOrWhiteSpace)
                    || bodyFonts.Distinct(StringComparer.OrdinalIgnoreCase)
                        .Count() != 1)) {
                style = null;
                return false;
            }
            string? bodyFont = bodyFonts.Length == 0
                ? null
                : bodyFonts[0];
            OfficeXmlNode title = chart.Child(NsC, "title");
            string? titleFont = ReadTitleTypeface(title,
                chartDefaultTypeface);
            if (!title.IsNone
                && HasText(title)
                && string.IsNullOrWhiteSpace(titleFont)) {
                style = null;
                return false;
            }
            if (!TryReadAxisTitleTypeface(chart, chartDefaultTypeface,
                    out _)) {
                style = null;
                return false;
            }
            style = bodyFont == null && titleFont == null
                ? null
                : new OfficeChartStyle(fontFamily: bodyFont,
                    titleFontFamily: titleFont);
            return true;
        }

        internal static string? ReadChartDefaultTypeface(OfficeXmlNode chart) =>
            chart.Parent.Child(NsC, "txPr")
                .Descendants(NsA, "latin")
                .Select(font => font.Attribute("typeface"))
                .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));

        internal bool TryReadAxisTitleTypeface(OfficeXmlNode chart,
            string? chartDefaultTypeface, out string? axisTitleFont) {
            OfficeXmlNode plotArea = chart.Child(NsC, "plotArea");
            var titles = new List<OfficeXmlNode>();
            foreach (OfficeXmlNode axis in plotArea.Children()) {
                if (!IsAxis(axis)) continue;
                OfficeXmlNode title = axis.Child(NsC, "title");
                if (title.IsNone) continue;
                if (HasText(title) ||
                    title.Child(NsC, "tx").Child(NsC, "strRef").Child(NsC, "strCache")
                        .Descendants(NsC, "pt").Any(point => !string.IsNullOrEmpty(point.Child(NsC, "v").InnerText)))
                    titles.Add(title);
            }
            string?[] fonts = titles.Select(title =>
                    ReadTitleTypeface(title, chartDefaultTypeface))
                .ToArray();
            if (fonts.Any(string.IsNullOrWhiteSpace)) {
                axisTitleFont = null;
                return false;
            }
            string[] distinct = fonts.Select(font => font!)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();
            axisTitleFont = distinct.Length == 1 ? distinct[0] : null;
            return distinct.Length <= 1;
        }

        private static bool IsAxis(OfficeXmlNode element) =>
            element.NamespaceId == NsC && element.LocalName is "catAx" or "dateAx" or "valAx" or "serAx";

        // Whether any a:t below the element holds text.
        private static bool HasText(OfficeXmlNode element) =>
            element.Descendants(NsA, "t").Any(text => !string.IsNullOrEmpty(text.InnerText));

        private string? ReadBodyTypeface(
            OfficeXmlNode textArea, string? chartDefaultTypeface) {
            string?[] typefaces = textArea.Descendants(NsC, "txPr")
                .Where(properties => properties.Ancestor(NsC, "title").IsNone)
                .SelectMany(properties => properties.Descendants(NsA, "latin"))
                .Select(font => font.Attribute("typeface"))
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .ToArray();
            if (typefaces.Length == 0) {
                return ResolveChartTypeface(chartDefaultTypeface,
                    useMajorWhenMissing: false);
            }
            string?[] resolved = typefaces.Select(value =>
                    ResolveChartTypeface(value, useMajorWhenMissing: false))
                .ToArray();
            if (resolved.Any(string.IsNullOrWhiteSpace)) return null;
            string[] explicitFonts = resolved.Select(value => value!)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();
            return explicitFonts.Length == 1 ? explicitFonts[0] : null;
        }

        private string? ReadTitleTypeface(OfficeXmlNode title,
            string? chartDefaultTypeface) {
            if (title.IsNone) return null;
            string? defaultTypeface = title.Children(NsC, "txPr").ToList()
                .SelectMany(properties => properties.Descendants(NsA, "latin"))
                .Select(font => font.Attribute("typeface"))
                .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value))
                ?? chartDefaultTypeface;
            OfficeXmlNode[] textRuns = title.Descendants(NsC, "rich")
                .SelectMany(richText => richText.DescendantsWhere(element =>
                    element.NamespaceId == NsA && element.LocalName is "r" or "fld"))
                .Where(run => !run.Child(NsA, "t").IsNone)
                .ToArray();
            if (textRuns.Length == 0) {
                return ResolveChartTypeface(defaultTypeface,
                    useMajorWhenMissing: true);
            }

            string?[] resolvedFonts = textRuns.Select(run =>
                    ResolveChartTypeface(ReadTitleRunTypeface(run)
                            ?? defaultTypeface,
                        useMajorWhenMissing: true))
                .ToArray();
            if (resolvedFonts.Any(string.IsNullOrWhiteSpace)) {
                return null;
            }
            string[] explicitFonts = resolvedFonts.Select(value => value!)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();
            return explicitFonts.Length == 1 ? explicitFonts[0] : null;
        }

        private string? ResolveChartTypeface(string? typeface,
            bool useMajorWhenMissing) {
            if (!string.IsNullOrWhiteSpace(typeface)
                && !IsThemeFontToken(typeface)) {
                return typeface;
            }
            OfficeXmlNode scheme = _scheme;
            if (scheme.IsNone) return null;
            if (string.IsNullOrWhiteSpace(typeface)) {
                return useMajorWhenMissing
                    ? SchemeTypeface(scheme, "majorFont")
                    : SchemeTypeface(scheme, "minorFont");
            }
            if (typeface!.StartsWith("+mj-", StringComparison.OrdinalIgnoreCase)) {
                return SchemeTypeface(scheme, "majorFont");
            }
            if (typeface.StartsWith("+mn-", StringComparison.OrdinalIgnoreCase)) {
                return SchemeTypeface(scheme, "minorFont");
            }
            return null;
        }

        // The Latin typeface of a font scheme's major or minor font.
        internal static string? SchemeTypeface(OfficeXmlNode fontScheme, string font) =>
            fontScheme.Child(NsA, font).Child(NsA, "latin").Attribute("typeface");

        private static string? ReadTitleRunTypeface(OfficeXmlNode run) {
            string? runTypeface = run.Child(NsA, "rPr")
                .Child(NsA, "latin").Attribute("typeface");
            if (!string.IsNullOrWhiteSpace(runTypeface)) return runTypeface;
            OfficeXmlNode paragraph = run.Ancestor(NsA, "p");
            OfficeXmlNode paragraphProperties = paragraph.Child(NsA, "pPr");
            string? paragraphTypeface = paragraphProperties
                .Child(NsA, "defRPr")
                .Child(NsA, "latin").Attribute("typeface");
            if (!string.IsNullOrWhiteSpace(paragraphTypeface)) {
                return paragraphTypeface;
            }

            int level = paragraphProperties.IntAttr("lvl") ?? 0;
            OfficeXmlNode listStyle = run.Ancestor(NsC, "rich").Child(NsA, "lstStyle");
            string levelName = $"lvl{level + 1}pPr";
            OfficeXmlNode levelProperties = default;
            foreach (OfficeXmlNode element in listStyle.Children()) {
                if (string.Equals(element.LocalName, levelName, StringComparison.Ordinal)) {
                    levelProperties = element;
                    break;
                }
            }
            return levelProperties.Child(NsA, "defRPr")
                .Child(NsA, "latin").Attribute("typeface");
        }

        private static bool IsRelevantBodyTextArea(OfficeXmlNode element) {
            if (element.NamespaceId == NsC && element.LocalName is "legend" or "catAx" or "valAx"
                or "dateAx" or "serAx" or "dispUnitsLbl"
                or "dTable" or "trendlineLbl") {
                return true;
            }
            // Data labels have their own layout font and may legitimately differ
            // from the legend and axes without changing the shared body font.
            return false;
        }

        private static bool IsThemeFontToken(string? typeface) =>
            !string.IsNullOrEmpty(typeface)
            && typeface![0] == '+';

    // SDK callers
    internal OfficeOpenXmlChartTextReader(A.FontScheme? scheme) : this(OfficeXmlNode.FromSdk(scheme)) { }
    internal static bool HasUncachedFormulaTitle(C.Chart chart) => HasUncachedFormulaTitle(OfficeXmlNode.FromSdk(chart));
    internal bool TryReadSharedTextStyle(C.Chart chart, out OfficeChartStyle? style) => TryReadSharedTextStyle(OfficeXmlNode.FromSdk(chart), out style);
    internal static string? ReadChartDefaultTypeface(C.Chart chart) => ReadChartDefaultTypeface(OfficeXmlNode.FromSdk(chart));
    internal bool TryReadAxisTitleTypeface(C.Chart chart, string? chartDefaultTypeface, out string? axisTitleFont) =>
        TryReadAxisTitleTypeface(OfficeXmlNode.FromSdk(chart), chartDefaultTypeface, out axisTitleFont);
}
