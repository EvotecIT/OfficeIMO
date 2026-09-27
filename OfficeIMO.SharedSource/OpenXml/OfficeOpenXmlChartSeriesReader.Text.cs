using System;
using System.Linq;
using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal;

/// <summary>Resolves chart typefaces through the native text inheritance and owner theme.</summary>
internal sealed class OfficeOpenXmlChartTextReader {
    private readonly A.FontScheme? _scheme;
    internal OfficeOpenXmlChartTextReader(A.FontScheme? scheme) => _scheme = scheme;
        internal bool TryReadSharedTextStyle(C.Chart chart,
            out OfficeChartStyle? style) {
            string? chartDefaultTypeface = ReadChartDefaultTypeface(chart);
            OpenXmlElement[] bodyTextAreas = chart.Descendants()
                .Where(IsRelevantBodyTextArea)
                .ToArray();
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
            C.Title? title = chart.GetFirstChild<C.Title>();
            string? titleFont = ReadTitleTypeface(title,
                chartDefaultTypeface);
            if (title != null
                && title.Descendants<A.Text>()
                    .Any(text => !string.IsNullOrEmpty(text.Text))
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

        internal static string? ReadChartDefaultTypeface(C.Chart chart) =>
            chart.Parent?.GetFirstChild<C.TextProperties>()?
                .Descendants<A.LatinFont>()
                .Select(font => font.Typeface?.Value)
                .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));

        internal bool TryReadAxisTitleTypeface(C.Chart chart,
            string? chartDefaultTypeface, out string? axisTitleFont) {
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            C.Title[] titles = plotArea == null
                ? Array.Empty<C.Title>()
                : plotArea.Elements()
                    .Where(axis => axis is C.CategoryAxis
                        or C.DateAxis or C.ValueAxis or C.SeriesAxis)
                    .Select(axis => axis.GetFirstChild<C.Title>())
                    .Where(title => title != null
                        && title.Descendants<A.Text>()
                            .Any(text => !string.IsNullOrEmpty(text.Text)))
                    .Cast<C.Title>()
                    .ToArray();
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

        private string? ReadBodyTypeface(
            OpenXmlElement textArea, string? chartDefaultTypeface) {
            string?[] typefaces = textArea.Descendants<C.TextProperties>()
                .Where(properties => !properties.Ancestors<C.Title>().Any())
                .SelectMany(properties => properties.Descendants<A.LatinFont>())
                .Select(font => font.Typeface?.Value)
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

        private string? ReadTitleTypeface(C.Title? title,
            string? chartDefaultTypeface) {
            if (title == null) return null;
            string? defaultTypeface = title.Elements<C.TextProperties>()
                .SelectMany(properties => properties.Descendants<A.LatinFont>())
                .Select(font => font.Typeface?.Value)
                .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value))
                ?? chartDefaultTypeface;
            OpenXmlElement[] textRuns = title.Descendants<C.RichText>()
                .SelectMany(richText => richText.Descendants()
                    .Where(element => element is A.Run or A.Field))
                .Where(run => run.GetFirstChild<A.Text>() != null)
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
            A.FontScheme? scheme = _scheme;
            if (scheme == null) return null;
            if (string.IsNullOrWhiteSpace(typeface)) {
                return useMajorWhenMissing
                    ? scheme.MajorFont?.LatinFont?.Typeface?.Value
                    : scheme.MinorFont?.LatinFont?.Typeface?.Value;
            }
            if (typeface!.StartsWith("+mj-", StringComparison.OrdinalIgnoreCase)) {
                return scheme.MajorFont?.LatinFont?.Typeface?.Value;
            }
            if (typeface.StartsWith("+mn-", StringComparison.OrdinalIgnoreCase)) {
                return scheme.MinorFont?.LatinFont?.Typeface?.Value;
            }
            return null;
        }

        private static string? ReadTitleRunTypeface(OpenXmlElement run) {
            string? runTypeface = run.GetFirstChild<A.RunProperties>()?
                .GetFirstChild<A.LatinFont>()?.Typeface?.Value;
            if (!string.IsNullOrWhiteSpace(runTypeface)) return runTypeface;
            A.Paragraph? paragraph = run.Ancestors<A.Paragraph>()
                .FirstOrDefault();
            string? paragraphTypeface = paragraph?
                .GetFirstChild<A.ParagraphProperties>()?
                .GetFirstChild<A.DefaultRunProperties>()?
                .GetFirstChild<A.LatinFont>()?.Typeface?.Value;
            if (!string.IsNullOrWhiteSpace(paragraphTypeface)) {
                return paragraphTypeface;
            }

            int level = paragraph?.GetFirstChild<A.ParagraphProperties>()?
                .Level?.Value ?? 0;
            A.ListStyle? listStyle = run.Ancestors<C.RichText>()
                .FirstOrDefault()?.GetFirstChild<A.ListStyle>();
            OpenXmlCompositeElement? levelProperties = listStyle?.ChildElements
                .OfType<OpenXmlCompositeElement>()
                .FirstOrDefault(element => string.Equals(element.LocalName,
                    $"lvl{level + 1}pPr", StringComparison.Ordinal));
            return levelProperties?.GetFirstChild<A.DefaultRunProperties>()?
                .GetFirstChild<A.LatinFont>()?.Typeface?.Value;
        }

        private static bool IsRelevantBodyTextArea(OpenXmlElement element) {
            if (element is C.Legend or C.CategoryAxis or C.ValueAxis
                or C.DateAxis or C.SeriesAxis or C.DisplayUnitsLabel
                or C.DataTable or C.TrendlineLabel) {
                return true;
            }
            if (element is not C.DataLabels labels) return false;
            return labels.GetFirstChild<C.ShowValue>()?.Val?.Value == true
                || labels.GetFirstChild<C.ShowCategoryName>()?.Val?.Value == true
                || labels.GetFirstChild<C.ShowSeriesName>()?.Val?.Value == true
                || labels.GetFirstChild<C.ShowPercent>()?.Val?.Value == true
                || labels.GetFirstChild<C.ShowBubbleSize>()?.Val?.Value == true
                || labels.GetFirstChild<C.ShowLegendKey>()?.Val?.Value == true;
        }

        private static bool IsThemeFontToken(string? typeface) =>
            !string.IsNullOrEmpty(typeface)
            && typeface![0] == '+';

}
