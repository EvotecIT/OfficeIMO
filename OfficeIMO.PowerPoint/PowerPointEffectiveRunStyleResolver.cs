using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    internal readonly struct PowerPointEffectiveRunStyle {
        internal PowerPointEffectiveRunStyle(
            bool? bold,
            bool? italic,
            PowerPointUnderlineStyle? underlineStyle,
            PowerPointStrikeStyle? strikeStyle,
            PowerPointCapitalization? capitalization,
            double? baselinePercent,
            double? fontSizePoints,
            string? fontName,
            string? color,
            string? language) {
            Bold = bold;
            Italic = italic;
            UnderlineStyle = underlineStyle;
            StrikeStyle = strikeStyle;
            Capitalization = capitalization;
            BaselinePercent = baselinePercent;
            FontSizePoints = fontSizePoints;
            FontName = fontName;
            Color = color;
            Language = language;
        }

        internal bool? Bold { get; }
        internal bool? Italic { get; }
        internal PowerPointUnderlineStyle? UnderlineStyle { get; }
        internal PowerPointStrikeStyle? StrikeStyle { get; }
        internal PowerPointCapitalization? Capitalization { get; }
        internal double? BaselinePercent { get; }
        internal double? FontSizePoints { get; }
        internal string? FontName { get; }
        internal string? Color { get; }
        internal string? Language { get; }
    }

    internal static class PowerPointEffectiveRunStyleResolver {
        internal static PowerPointEffectiveRunStyle Resolve(
            PowerPointTextRun run,
            PowerPointParagraph paragraph,
            A.ListStyle? listStyle,
            OpenXmlCompositeElement? masterTextStyle) {
            IReadOnlyList<A.TextCharacterPropertiesType> sources = ResolveSources(run, paragraph, listStyle, masterTextStyle);
            bool? bold = sources.Select(source => source.Bold == null ? (bool?)null : source.Bold.Value)
                .FirstOrDefault(value => value.HasValue);
            bool? italic = sources.Select(source => source.Italic == null ? (bool?)null : source.Italic.Value)
                .FirstOrDefault(value => value.HasValue);
            A.TextUnderlineValues? underline = sources.Select(source => source.Underline?.Value)
                .FirstOrDefault(value => value.HasValue);
            A.TextStrikeValues? strike = sources.Select(source => source.Strike?.Value)
                .FirstOrDefault(value => value.HasValue);
            A.TextCapsValues? capitalization = sources.Select(source => source.Capital?.Value)
                .FirstOrDefault(value => value.HasValue);
            int? baseline = sources.Select(source => source.Baseline?.Value)
                .FirstOrDefault(value => value.HasValue);
            int? fontSize = sources.Select(source => source.FontSize?.Value)
                .FirstOrDefault(value => value.HasValue);
            string? fontName = sources.Select(source => source.GetFirstChild<A.LatinFont>()?.Typeface?.Value)
                .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
            string? color = ResolveColor(run, sources);
            string? language = sources.Select(source => source.Language?.Value)
                .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));

            return new PowerPointEffectiveRunStyle(
                bold,
                italic,
                underline?.ToOfficeEnum(),
                ToStrikeStyle(strike),
                ToCapitalization(capitalization),
                baseline.HasValue ? baseline.Value / 1000D : (double?)null,
                fontSize.HasValue ? fontSize.Value / 100D : (double?)null,
                fontName,
                color,
                language);
        }

        private static string? ResolveColor(
            PowerPointTextRun run,
            IReadOnlyList<A.TextCharacterPropertiesType> sources) {
            A.ColorScheme? colorScheme = ResolveColorScheme(run.OwnerPart);
            OpenXmlCompositeElement? colorMap = ResolveColorMap(run.OwnerPart);
            foreach (A.TextCharacterPropertiesType source in sources) {
                A.SolidFill? fill = source.GetFirstChild<A.SolidFill>();
                if (fill == null) continue;
                OpenXmlElement effectiveFill = ApplyColorMap(fill, colorMap);
                OfficeColor? color = OfficeOpenXmlThemeColorResolver.ResolveColor(effectiveFill, colorScheme);
                if (color.HasValue) return color.Value.ToRgbHex();
            }

            return null;
        }

        private static OpenXmlElement ApplyColorMap(A.SolidFill fill, OpenXmlCompositeElement? colorMap) {
            A.SchemeColor? scheme = fill.GetFirstChild<A.SchemeColor>();
            string? schemeName = scheme?.GetAttribute("val", string.Empty).Value;
            string? mapped = MapSchemeColor(schemeName, colorMap);
            if (string.IsNullOrWhiteSpace(mapped)
                || string.Equals(mapped, schemeName, StringComparison.OrdinalIgnoreCase)) return fill;
            var clone = (A.SolidFill)fill.CloneNode(true);
            clone.GetFirstChild<A.SchemeColor>()!
                .SetAttribute(new OpenXmlAttribute("val", string.Empty, mapped!));
            return clone;
        }

        private static string? MapSchemeColor(string? scheme, OpenXmlCompositeElement? colorMap) {
            if (string.IsNullOrWhiteSpace(scheme) || colorMap == null) return scheme;
            string normalized = scheme!.Trim().ToLowerInvariant();
            string? attribute = normalized switch {
                "background1" or "bg1" => "bg1",
                "text1" or "tx1" => "tx1",
                "background2" or "bg2" => "bg2",
                "text2" or "tx2" => "tx2",
                "accent1" => "accent1",
                "accent2" => "accent2",
                "accent3" => "accent3",
                "accent4" => "accent4",
                "accent5" => "accent5",
                "accent6" => "accent6",
                "hyperlink" or "hlink" => "hlink",
                "followedhyperlink" or "folhlink" => "folHlink",
                _ => null
            };
            if (attribute == null) return scheme;
            string? mapped = colorMap.GetAttribute(attribute, string.Empty).Value;
            return string.IsNullOrWhiteSpace(mapped) ? scheme : mapped;
        }

        private static A.ColorScheme? ResolveColorScheme(OpenXmlPartContainer? ownerPart) => ownerPart switch {
            SlidePart slidePart => slidePart.ThemeOverridePart?.ThemeOverride?.ColorScheme
                ?? slidePart.SlideLayoutPart?.ThemeOverridePart?.ThemeOverride?.ColorScheme
                ?? slidePart.SlideLayoutPart?.SlideMasterPart?.ThemePart?.Theme?.ThemeElements?.ColorScheme,
            SlideLayoutPart layoutPart => layoutPart.ThemeOverridePart?.ThemeOverride?.ColorScheme
                ?? layoutPart.SlideMasterPart?.ThemePart?.Theme?.ThemeElements?.ColorScheme,
            SlideMasterPart masterPart => masterPart.ThemePart?.Theme?.ThemeElements?.ColorScheme,
            _ => null
        };

        private static OpenXmlCompositeElement? ResolveColorMap(OpenXmlPartContainer? ownerPart) => ownerPart switch {
            SlidePart slidePart => slidePart.Slide?.ColorMapOverride?.GetFirstChild<A.OverrideColorMapping>()
                ?? slidePart.SlideLayoutPart?.SlideLayout?.ColorMapOverride?.GetFirstChild<A.OverrideColorMapping>()
                ?? (OpenXmlCompositeElement?)slidePart.SlideLayoutPart?.SlideMasterPart?.SlideMaster?.ColorMap,
            SlideLayoutPart layoutPart => layoutPart.SlideLayout?.ColorMapOverride?.GetFirstChild<A.OverrideColorMapping>()
                ?? (OpenXmlCompositeElement?)layoutPart.SlideMasterPart?.SlideMaster?.ColorMap,
            SlideMasterPart masterPart => masterPart.SlideMaster?.ColorMap,
            _ => null
        };

        private static IReadOnlyList<A.TextCharacterPropertiesType> ResolveSources(
            PowerPointTextRun run,
            PowerPointParagraph paragraph,
            A.ListStyle? listStyle,
            OpenXmlCompositeElement? masterTextStyle) {
            var sources = new List<A.TextCharacterPropertiesType>();
            if (run.RunProperties != null) sources.Add(run.RunProperties);
            A.DefaultRunProperties? paragraphDefaults = paragraph.Paragraph.ParagraphProperties?
                .GetFirstChild<A.DefaultRunProperties>();
            if (paragraphDefaults != null) sources.Add(paragraphDefaults);
            int level = paragraph.Paragraph.ParagraphProperties?.Level?.Value ?? 0;
            sources.AddRange(FindDefaultRunProperties(listStyle, level));
            sources.AddRange(FindDefaultRunProperties(masterTextStyle, level));
            return sources;
        }

        private static IEnumerable<A.DefaultRunProperties> FindDefaultRunProperties(
            OpenXmlCompositeElement? container,
            int level) {
            A.DefaultRunProperties? levelDefaults = container?
                .ChildElements
                .OfType<A.TextParagraphPropertiesType>()
                .FirstOrDefault(properties => GetTextLevel(properties) == level)?
                .GetFirstChild<A.DefaultRunProperties>();
            if (levelDefaults != null) yield return levelDefaults;
            A.DefaultRunProperties? fallbackDefaults = container?
                .GetFirstChild<A.DefaultParagraphProperties>()?
                .GetFirstChild<A.DefaultRunProperties>();
            if (fallbackDefaults != null) yield return fallbackDefaults;
        }

        private static int GetTextLevel(A.TextParagraphPropertiesType properties) => properties switch {
            A.Level1ParagraphProperties => 0,
            A.Level2ParagraphProperties => 1,
            A.Level3ParagraphProperties => 2,
            A.Level4ParagraphProperties => 3,
            A.Level5ParagraphProperties => 4,
            A.Level6ParagraphProperties => 5,
            A.Level7ParagraphProperties => 6,
            A.Level8ParagraphProperties => 7,
            A.Level9ParagraphProperties => 8,
            _ => -1
        };

        private static PowerPointStrikeStyle? ToStrikeStyle(A.TextStrikeValues? value) {
            if (!value.HasValue) return null;
            if (value.Value == A.TextStrikeValues.NoStrike) return PowerPointStrikeStyle.None;
            if (value.Value == A.TextStrikeValues.SingleStrike) return PowerPointStrikeStyle.Single;
            if (value.Value == A.TextStrikeValues.DoubleStrike) return PowerPointStrikeStyle.Double;
            return null;
        }

        private static PowerPointCapitalization? ToCapitalization(A.TextCapsValues? value) {
            if (!value.HasValue) return null;
            if (value.Value == A.TextCapsValues.None) return PowerPointCapitalization.None;
            if (value.Value == A.TextCapsValues.Small) return PowerPointCapitalization.SmallCaps;
            if (value.Value == A.TextCapsValues.All) return PowerPointCapitalization.AllCaps;
            return null;
        }
    }
}
