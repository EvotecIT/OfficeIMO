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
            OpenXmlCompositeElement? masterTextStyle,
            IReadOnlyList<A.TableCellTextStyle>? tableTextStyles = null) {
            IReadOnlyList<A.TextCharacterPropertiesType> directSources = ResolveDirectSources(run, paragraph, listStyle);
            IReadOnlyList<A.TextCharacterPropertiesType> masterSources = FindDefaultRunProperties(
                masterTextStyle,
                paragraph.Paragraph.ParagraphProperties?.Level?.Value ?? 0).Cast<A.TextCharacterPropertiesType>().ToArray();
            IReadOnlyList<A.TableCellTextStyle> tableSources = tableTextStyles ?? Array.Empty<A.TableCellTextStyle>();
            bool? bold = ResolveBoolean(directSources, static source => source.Bold?.Value)
                ?? ResolveTableBoolean(tableSources, static source => source.Bold?.Value)
                ?? ResolveBoolean(masterSources, static source => source.Bold?.Value);
            bool? italic = ResolveBoolean(directSources, static source => source.Italic?.Value)
                ?? ResolveTableBoolean(tableSources, static source => source.Italic?.Value)
                ?? ResolveBoolean(masterSources, static source => source.Italic?.Value);
            A.TextUnderlineValues? underline = ResolveValue(directSources, static source => source.Underline?.Value)
                ?? ResolveValue(masterSources, static source => source.Underline?.Value);
            A.TextStrikeValues? strike = ResolveValue(directSources, static source => source.Strike?.Value)
                ?? ResolveValue(masterSources, static source => source.Strike?.Value);
            A.TextCapsValues? capitalization = ResolveValue(directSources, static source => source.Capital?.Value)
                ?? ResolveValue(masterSources, static source => source.Capital?.Value);
            int? baseline = ResolveValue(directSources, static source => source.Baseline?.Value)
                ?? ResolveValue(masterSources, static source => source.Baseline?.Value);
            int? fontSize = ResolveValue(directSources, static source => source.FontSize?.Value)
                ?? ResolveValue(masterSources, static source => source.FontSize?.Value);
            string? language = ResolveString(directSources, static source => source.Language?.Value)
                ?? ResolveString(masterSources, static source => source.Language?.Value);
            PowerPointFontScript fontScript = ResolveFontScript(run.Text, language);
            string? fontName = ResolveFontName(run, directSources, fontScript)
                ?? ResolveTableFontName(run, tableSources, fontScript)
                ?? ResolveFontName(run, masterSources, fontScript);
            string? color = ResolveColor(run, directSources)
                ?? ResolveTableColor(run, tableSources)
                ?? ResolveColor(run, masterSources);

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

        private static T? ResolveValue<T>(
            IEnumerable<A.TextCharacterPropertiesType> sources,
            Func<A.TextCharacterPropertiesType, T?> selector) where T : struct =>
            sources.Select(selector).FirstOrDefault(static value => value.HasValue);

        private static bool? ResolveBoolean(
            IEnumerable<A.TextCharacterPropertiesType> sources,
            Func<A.TextCharacterPropertiesType, bool?> selector) =>
            sources.Select(selector).FirstOrDefault(static value => value.HasValue);

        private static bool? ResolveTableBoolean(
            IEnumerable<A.TableCellTextStyle> sources,
            Func<A.TableCellTextStyle, A.BooleanStyleValues?> selector) {
            A.BooleanStyleValues? value = sources.Select(selector).FirstOrDefault(static item => item.HasValue);
            if (!value.HasValue) return null;
            if (value.Value == A.BooleanStyleValues.On) return true;
            if (value.Value == A.BooleanStyleValues.Off) return false;
            return null;
        }

        private static string? ResolveString(
            IEnumerable<A.TextCharacterPropertiesType> sources,
            Func<A.TextCharacterPropertiesType, string?> selector) =>
            sources.Select(selector).FirstOrDefault(static value => !string.IsNullOrWhiteSpace(value));

        private static string? ResolveFontName(
            PowerPointTextRun run,
            IReadOnlyList<A.TextCharacterPropertiesType> sources,
            PowerPointFontScript script) {
            string? typeface = sources.Select(source => ResolveTypeface(source, script))
                .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
            if (string.IsNullOrWhiteSpace(typeface)) return null;
            string token = typeface!.Trim().ToLowerInvariant();
            if (token is not ("+mj-lt" or "+mn-lt" or "+mj-ea" or "+mn-ea" or "+mj-cs" or "+mn-cs")) {
                return typeface;
            }

            foreach (A.FontScheme scheme in EnumerateFontSchemes(run.OwnerPart)) {
                string? resolved = token switch {
                    "+mj-lt" => scheme.MajorFont?.LatinFont?.Typeface?.Value,
                    "+mn-lt" => scheme.MinorFont?.LatinFont?.Typeface?.Value,
                    "+mj-ea" => scheme.MajorFont?.EastAsianFont?.Typeface?.Value,
                    "+mn-ea" => scheme.MinorFont?.EastAsianFont?.Typeface?.Value,
                    "+mj-cs" => scheme.MajorFont?.ComplexScriptFont?.Typeface?.Value,
                    "+mn-cs" => scheme.MinorFont?.ComplexScriptFont?.Typeface?.Value,
                    _ => null
                };
                if (!string.IsNullOrWhiteSpace(resolved)) return resolved;
            }

            return PowerPointTextDefaults.LegacyFallbackFontFamily;
        }

        private static string? ResolveTableFontName(
            PowerPointTextRun run,
            IEnumerable<A.TableCellTextStyle> sources,
            PowerPointFontScript script) {
            foreach (A.TableCellTextStyle source in sources) {
                string? typeface = ResolveTypeface(source.GetFirstChild<A.Fonts>(), script);
                if (!string.IsNullOrWhiteSpace(typeface)) return typeface;
                A.FontReference? reference = source.GetFirstChild<A.FontReference>();
                if (reference?.Index?.Value == A.FontCollectionIndexValues.Major) {
                    return ResolveThemeFontName(run.OwnerPart, major: true, script);
                }
                if (reference?.Index?.Value == A.FontCollectionIndexValues.Minor) {
                    return ResolveThemeFontName(run.OwnerPart, major: false, script);
                }
            }
            return null;
        }

        private static string? ResolveThemeFontName(
            OpenXmlPartContainer? ownerPart,
            bool major,
            PowerPointFontScript script) {
            foreach (A.FontScheme scheme in EnumerateFontSchemes(ownerPart)) {
                A.FontCollectionType? collection = major ? scheme.MajorFont : scheme.MinorFont;
                string? typeface = ResolveTypeface(collection, script);
                if (!string.IsNullOrWhiteSpace(typeface)) return typeface;
            }
            return PowerPointTextDefaults.LegacyFallbackFontFamily;
        }

        private static string? ResolveTypeface(OpenXmlCompositeElement? source, PowerPointFontScript script) {
            if (source == null) return null;
            string? preferred = script switch {
                PowerPointFontScript.EastAsian => source.GetFirstChild<A.EastAsianFont>()?.Typeface?.Value,
                PowerPointFontScript.ComplexScript => source.GetFirstChild<A.ComplexScriptFont>()?.Typeface?.Value,
                _ => source.GetFirstChild<A.LatinFont>()?.Typeface?.Value
            };
            return !string.IsNullOrWhiteSpace(preferred)
                ? preferred
                : source.GetFirstChild<A.LatinFont>()?.Typeface?.Value
                  ?? source.GetFirstChild<A.EastAsianFont>()?.Typeface?.Value
                  ?? source.GetFirstChild<A.ComplexScriptFont>()?.Typeface?.Value;
        }

        private static PowerPointFontScript ResolveFontScript(string? text, string? language) {
            if (!string.IsNullOrEmpty(text)) {
                for (int index = 0; index < text!.Length; index++) {
                    int scalar = text[index];
                    if (char.IsHighSurrogate(text[index]) && index + 1 < text.Length && char.IsLowSurrogate(text[index + 1])) {
                        scalar = char.ConvertToUtf32(text[index], text[++index]);
                    }
                    if (IsEastAsianScalar(scalar)) return PowerPointFontScript.EastAsian;
                }
                if (OfficeTextElements.ContainsRightToLeft(text)
                    || OfficeTextElements.ContainsShapingRequiredScript(text)) {
                    return PowerPointFontScript.ComplexScript;
                }
            }

            string primaryLanguage = (language ?? string.Empty).Split('-')[0].Trim().ToLowerInvariant();
            if (primaryLanguage is "zh" or "ja" or "ko") return PowerPointFontScript.EastAsian;
            if (primaryLanguage is "ar" or "he" or "fa" or "ur" or "ps" or "dv" or "syr" or "yi"
                or "hi" or "bn" or "pa" or "gu" or "or" or "ta" or "te" or "kn" or "ml" or "si"
                or "th" or "lo" or "my" or "km") return PowerPointFontScript.ComplexScript;
            return PowerPointFontScript.Latin;
        }

        private static bool IsEastAsianScalar(int scalar) =>
            scalar is >= 0x2E80 and <= 0xA4CF
            || scalar is >= 0xAC00 and <= 0xD7AF
            || scalar is >= 0xF900 and <= 0xFAFF
            || scalar is >= 0xFE30 and <= 0xFE4F
            || scalar is >= 0xFF00 and <= 0xFFEF
            || scalar is >= 0x20000 and <= 0x323AF;

        private enum PowerPointFontScript {
            Latin,
            EastAsian,
            ComplexScript
        }

        private static IEnumerable<A.FontScheme> EnumerateFontSchemes(OpenXmlPartContainer? ownerPart) {
            if (ownerPart is SlidePart slidePart) {
                if (slidePart.ThemeOverridePart?.ThemeOverride?.FontScheme is A.FontScheme slideScheme) yield return slideScheme;
                if (slidePart.SlideLayoutPart?.ThemeOverridePart?.ThemeOverride?.FontScheme is A.FontScheme layoutScheme) yield return layoutScheme;
                if (slidePart.SlideLayoutPart?.SlideMasterPart?.ThemePart?.Theme?.ThemeElements?.FontScheme is A.FontScheme masterScheme) yield return masterScheme;
                yield break;
            }
            if (ownerPart is SlideLayoutPart layoutPart) {
                if (layoutPart.ThemeOverridePart?.ThemeOverride?.FontScheme is A.FontScheme layoutScheme) yield return layoutScheme;
                if (layoutPart.SlideMasterPart?.ThemePart?.Theme?.ThemeElements?.FontScheme is A.FontScheme masterScheme) yield return masterScheme;
                yield break;
            }
            if (ownerPart is SlideMasterPart masterPart
                && masterPart.ThemePart?.Theme?.ThemeElements?.FontScheme is A.FontScheme scheme) {
                yield return scheme;
            }
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

        private static string? ResolveTableColor(
            PowerPointTextRun run,
            IEnumerable<A.TableCellTextStyle> sources) {
            A.ColorScheme? colorScheme = ResolveColorScheme(run.OwnerPart);
            OpenXmlCompositeElement? colorMap = ResolveColorMap(run.OwnerPart);
            foreach (A.TableCellTextStyle source in sources) {
                OpenXmlElement? color = source.ChildElements.FirstOrDefault(IsColorElement)
                    ?? source.GetFirstChild<A.FontReference>()?.ChildElements.FirstOrDefault(IsColorElement);
                if (color == null) continue;
                OpenXmlElement effectiveColor = ApplyColorMap(color, colorMap);
                OfficeColor? resolved = OfficeOpenXmlThemeColorResolver.ResolveColor(effectiveColor, colorScheme);
                if (resolved.HasValue) return resolved.Value.ToRgbHex();
            }
            return null;
        }

        private static bool IsColorElement(OpenXmlElement element) => element is
            A.RgbColorModelHex or A.RgbColorModelPercentage or A.HslColor or A.SystemColor or A.SchemeColor or A.PresetColor;

        private static OpenXmlElement ApplyColorMap(OpenXmlElement colorOwner, OpenXmlCompositeElement? colorMap) {
            A.SchemeColor? scheme = colorOwner as A.SchemeColor ?? colorOwner.GetFirstChild<A.SchemeColor>();
            string? schemeName = scheme?.GetAttribute("val", string.Empty).Value;
            string? mapped = MapSchemeColor(schemeName, colorMap);
            if (string.IsNullOrWhiteSpace(mapped)
                || string.Equals(mapped, schemeName, StringComparison.OrdinalIgnoreCase)) return colorOwner;
            OpenXmlElement clone = colorOwner.CloneNode(true);
            (clone as A.SchemeColor ?? clone.GetFirstChild<A.SchemeColor>())!
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

        private static IReadOnlyList<A.TextCharacterPropertiesType> ResolveDirectSources(
            PowerPointTextRun run,
            PowerPointParagraph paragraph,
            A.ListStyle? listStyle) {
            var sources = new List<A.TextCharacterPropertiesType>();
            if (run.RunProperties != null) sources.Add(run.RunProperties);
            A.DefaultRunProperties? paragraphDefaults = paragraph.Paragraph.ParagraphProperties?
                .GetFirstChild<A.DefaultRunProperties>();
            if (paragraphDefaults != null) sources.Add(paragraphDefaults);
            int level = paragraph.Paragraph.ParagraphProperties?.Level?.Value ?? 0;
            sources.AddRange(FindDefaultRunProperties(listStyle, level));
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
