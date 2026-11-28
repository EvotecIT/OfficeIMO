using System.Globalization;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;

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

    internal readonly struct PowerPointEffectiveTextSegment {
        internal PowerPointEffectiveTextSegment(string text, PowerPointEffectiveRunStyle style) {
            Text = text;
            Style = style;
        }

        internal string Text { get; }
        internal PowerPointEffectiveRunStyle Style { get; }
    }

    internal static partial class PowerPointEffectiveRunStyleResolver {
        private const int Drawing = OfficeXmlNamespaces.DrawingML;
        private const int Presentation = OfficeXmlNamespaces.PresentationML;

        /// <summary>The paragraph mark's style from the paragraph's own and inherited properties, without a list or master text style.</summary>
        internal static PowerPointEffectiveRunStyle ResolveParagraph(PowerPointParagraph paragraph) => ResolveParagraph(paragraph, default, default);

        /// <summary>A run's style from its own, paragraph and inherited properties, without a list or master text style.</summary>
        internal static PowerPointEffectiveRunStyle Resolve(PowerPointTextRun run, PowerPointParagraph paragraph) => Resolve(run, paragraph, default, default);

        // Blank paragraphs still have a formatted paragraph mark and inherited line metrics.
        internal static PowerPointEffectiveRunStyle ResolveParagraph(
            PowerPointParagraph paragraph, OfficeXmlNode listStyle, OfficeXmlNode masterTextStyle) {
            PowerPointTextRun? run = paragraph.InlineNodes.FirstOrDefault(node => node.Run != null)?.Run;
            return Resolve(run, paragraph, listStyle, masterTextStyle, null, null, includeParagraphMark: true);
        }

        internal static PowerPointEffectiveRunStyle Resolve(
            PowerPointTextRun run,
            PowerPointParagraph paragraph,
            OfficeXmlNode listStyle,
            OfficeXmlNode masterTextStyle,
            IReadOnlyList<OfficeXmlNode>? tableTextStyles = null) {
            return Resolve(run, paragraph, listStyle, masterTextStyle, tableTextStyles, null);
        }

        internal static IReadOnlyList<PowerPointEffectiveTextSegment> ResolveSegments(
            PowerPointTextRun run,
            PowerPointParagraph paragraph,
            OfficeXmlNode listStyle,
            OfficeXmlNode masterTextStyle,
            IReadOnlyList<OfficeXmlNode>? tableTextStyles = null) {
            string text = run.Text ?? string.Empty;
            if (text.Length == 0) return Array.Empty<PowerPointEffectiveTextSegment>();
            var raw = new List<(string Text, PowerPointFontScript? Script)>();
            TextElementEnumerator enumerator = StringInfo.GetTextElementEnumerator(text);
            while (enumerator.MoveNext()) {
                string element = enumerator.GetTextElement();
                PowerPointFontScript? script = ResolveStrongFontScript(element);
                if (raw.Count > 0 && script.HasValue && raw[raw.Count - 1].Script == script) {
                    (string previous, PowerPointFontScript? previousScript) = raw[raw.Count - 1];
                    raw[raw.Count - 1] = (previous + element, previousScript);
                } else if (raw.Count > 0 && !script.HasValue) {
                    (string previous, PowerPointFontScript? previousScript) = raw[raw.Count - 1];
                    raw[raw.Count - 1] = (previous + element, previousScript);
                } else {
                    raw.Add((element, script));
                }
            }
            if (raw.Count > 1 && !raw[0].Script.HasValue) {
                raw[1] = (raw[0].Text + raw[1].Text, raw[1].Script);
                raw.RemoveAt(0);
            }
            PowerPointFontScript fallback = ResolveFontScript(text, run.Language);
            return raw.Select(segment => new PowerPointEffectiveTextSegment(
                segment.Text,
                Resolve(run, paragraph, listStyle, masterTextStyle, tableTextStyles, segment.Script ?? fallback))).ToArray();
        }

        private static PowerPointEffectiveRunStyle Resolve(
            PowerPointTextRun? run,
            PowerPointParagraph paragraph,
            OfficeXmlNode listStyle,
            OfficeXmlNode masterTextStyle,
            IReadOnlyList<OfficeXmlNode>? tableTextStyles,
            PowerPointFontScript? fontScriptOverride, bool includeParagraphMark = false) {
            OfficeXmlNode paragraphProperties = paragraph.Element.Child(Drawing, "pPr");
            List<OfficeXmlNode> directSources = ResolveDirectSources(run, paragraph, listStyle, includeParagraphMark);
            OfficeXmlNode[] masterSources = FindDefaultRunProperties(masterTextStyle, ParagraphLevel(paragraphProperties)).ToArray();
            IReadOnlyList<OfficeXmlNode> tableSources = tableTextStyles ?? Array.Empty<OfficeXmlNode>();
            bool? bold = ResolveValue(directSources, static source => source.BooleanAttribute("b"))
                ?? ResolveTableBoolean(tableSources, "b")
                ?? ResolveValue(masterSources, static source => source.BooleanAttribute("b"));
            bool? italic = ResolveValue(directSources, static source => source.BooleanAttribute("i"))
                ?? ResolveTableBoolean(tableSources, "i")
                ?? ResolveValue(masterSources, static source => source.BooleanAttribute("i"));
            string? underline = ResolveAttribute(directSources, "u") ?? ResolveAttribute(masterSources, "u");
            string? strike = ResolveAttribute(directSources, "strike") ?? ResolveAttribute(masterSources, "strike");
            string? capitalization = ResolveAttribute(directSources, "cap") ?? ResolveAttribute(masterSources, "cap");
            int? baseline = ResolveValue(directSources, static source => source.Int32Attribute("baseline"))
                ?? ResolveValue(masterSources, static source => source.Int32Attribute("baseline"));
            int? fontSize = ResolveValue(directSources, static source => source.Int32Attribute("sz"))
                ?? ResolveValue(masterSources, static source => source.Int32Attribute("sz"));
            string? language = ResolveAttribute(directSources, "lang") ?? ResolveAttribute(masterSources, "lang");
            PowerPointFontScript fontScript = fontScriptOverride ?? ResolveFontScript(run?.Text, language);
            PowerPointPartRef owner = run?.OwnerPartRef ?? paragraph.OwnerPartRef;
            string? fontName = ResolveFontName(owner, directSources, fontScript)
                ?? ResolveTableFontName(owner, tableSources, fontScript)
                ?? ResolveFontName(owner, masterSources, fontScript);
            string? color = ResolveColor(owner, directSources)
                ?? ResolveTableColor(owner, tableSources)
                ?? ResolveColor(owner, masterSources);

            return new PowerPointEffectiveRunStyle(
                bold,
                italic,
                underline != null && PowerPointOpenXmlValueTypeExtensions.TryParseOpenXmlValue(underline, out PowerPointUnderlineStyle underlineStyle)
                    ? underlineStyle : null,
                ToStrikeStyle(strike),
                ToCapitalization(capitalization),
                baseline.HasValue ? baseline.Value / 1000D : (double?)null,
                fontSize.HasValue ? fontSize.Value / 100D : (double?)null,
                fontName,
                color,
                language);
        }

        private static PowerPointFontScript? ResolveStrongFontScript(string text) {
            for (int index = 0; index < text.Length; index++) {
                int scalar = text[index];
                if (char.IsHighSurrogate(text[index]) && index + 1 < text.Length && char.IsLowSurrogate(text[index + 1])) {
                    scalar = char.ConvertToUtf32(text[index], text[++index]);
                }
                if (IsEastAsianScalar(scalar)) return PowerPointFontScript.EastAsian;
            }
            if (OfficeTextElements.ContainsRightToLeft(text) || OfficeTextElements.ContainsShapingRequiredScript(text)) {
                return PowerPointFontScript.ComplexScript;
            }
            foreach (char character in text) {
                UnicodeCategory category = char.GetUnicodeCategory(character);
                if (category is UnicodeCategory.UppercaseLetter or UnicodeCategory.LowercaseLetter or UnicodeCategory.TitlecaseLetter
                    or UnicodeCategory.ModifierLetter or UnicodeCategory.OtherLetter or UnicodeCategory.DecimalDigitNumber
                    or UnicodeCategory.LetterNumber or UnicodeCategory.OtherNumber) return PowerPointFontScript.Latin;
            }
            return null;
        }

        private static T? ResolveValue<T>(IReadOnlyList<OfficeXmlNode> sources,
            Func<OfficeXmlNode, T?> selector) where T : struct {
            for (int index = 0; index < sources.Count; index++) {
                if (selector(sources[index]) is { } value) return value;
            }

            return null;
        }

        private static bool? ResolveTableBoolean(IReadOnlyList<OfficeXmlNode> sources, string attribute) {
            for (int index = 0; index < sources.Count; index++) {
                switch (sources[index].Attribute(attribute)) {
                    case "on": return true;
                    case "off": return false;
                    case "def": return null;
                }
            }

            return null;
        }

        private static string? ResolveAttribute(IReadOnlyList<OfficeXmlNode> sources, string attribute) {
            for (int index = 0; index < sources.Count; index++) {
                string? value = sources[index].Attribute(attribute);
                if (!string.IsNullOrWhiteSpace(value)) return value;
            }

            return null;
        }

        private static string? ResolveFontName(
            PowerPointPartRef owner, IReadOnlyList<OfficeXmlNode> sources, PowerPointFontScript script) {
            for (int index = 0; index < sources.Count; index++) {
                string? typeface = ResolveTypeface(sources[index], script);
                if (!string.IsNullOrWhiteSpace(typeface)) return ResolveThemeFontReference(typeface, owner);
            }

            return null;
        }

        /// <summary>Resolves a DrawingML theme reference at its relationship owner; literal family names pass through.</summary>
        internal static string? ResolveThemeFontReference(string? typeface, PowerPointPartRef ownerPart) {
            if (string.IsNullOrWhiteSpace(typeface)) return null;
            string token = typeface!.Trim().ToLowerInvariant();
            if (token is not ("+mj-lt" or "+mn-lt" or "+mj-ea" or "+mn-ea" or "+mj-cs" or "+mn-cs")) {
                return typeface;
            }

            string collection = token[2] == 'j' ? "majorFont" : "minorFont";
            string script = token.Substring(4) switch {
                "ea" => "ea",
                "cs" => "cs",
                _ => "latin"
            };
            foreach (OfficeXmlNode scheme in EnumerateFontSchemes(ownerPart)) {
                string? resolved = scheme.Child(Drawing, collection).Child(Drawing, script).Attribute("typeface");
                if (!string.IsNullOrWhiteSpace(resolved)) return resolved;
            }

            return PowerPointTextDefaults.LegacyFallbackFontFamily;
        }

        private static string? ResolveTableFontName(
            PowerPointPartRef owner, IReadOnlyList<OfficeXmlNode> sources, PowerPointFontScript script) {
            for (int index = 0; index < sources.Count; index++) {
                OfficeXmlNode source = sources[index];
                string? typeface = ResolveTypeface(source.Child(Drawing, "font"), script);
                if (!string.IsNullOrWhiteSpace(typeface)) return typeface;
                switch (source.Child(Drawing, "fontRef").Attribute("idx")) {
                    case "major": return ResolveThemeFontName(owner, major: true, script);
                    case "minor": return ResolveThemeFontName(owner, major: false, script);
                }
            }

            return null;
        }

        /// <summary>Resolves table and preview Latin defaults in the relationship owner's theme.</summary>
        internal static string? ResolveLatinThemeFontName(PowerPointPartRef ownerPart, bool major = false) =>
            ResolveThemeFontName(ownerPart, major, PowerPointFontScript.Latin);

        private static string? ResolveThemeFontName(PowerPointPartRef ownerPart, bool major, PowerPointFontScript script) {
            foreach (OfficeXmlNode scheme in EnumerateFontSchemes(ownerPart)) {
                string? typeface = ResolveTypeface(scheme.Child(Drawing, major ? "majorFont" : "minorFont"), script);
                if (!string.IsNullOrWhiteSpace(typeface)) return typeface;
            }

            return PowerPointTextDefaults.LegacyFallbackFontFamily;
        }

        private static string? ResolveTypeface(OfficeXmlNode source, PowerPointFontScript script) {
            if (source.IsNone) return null;
            string? preferred = source.Child(Drawing, script switch {
                PowerPointFontScript.EastAsian => "ea",
                PowerPointFontScript.ComplexScript => "cs",
                _ => "latin"
            }).Attribute("typeface");
            return !string.IsNullOrWhiteSpace(preferred)
                ? preferred
                : source.Child(Drawing, "latin").Attribute("typeface")
                  ?? source.Child(Drawing, "ea").Attribute("typeface")
                  ?? source.Child(Drawing, "cs").Attribute("typeface");
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

        private static IEnumerable<OfficeXmlNode> EnumerateFontSchemes(PowerPointPartRef part) {
            PowerPointPartKind kind = part.Kind;
            if (kind == PowerPointPartKind.Slide) {
                OfficeXmlNode scheme = part.RelatedPart(PowerPointPartRef.ThemeOverrideRelationshipType).Root.Child(Drawing, "fontScheme");
                if (scheme.Exists) yield return scheme;
                part = part.RelatedPart(PowerPointPartRef.SlideLayoutRelationshipType);
            }

            if (kind is PowerPointPartKind.Slide or PowerPointPartKind.SlideLayout) {
                OfficeXmlNode scheme = part.RelatedPart(PowerPointPartRef.ThemeOverrideRelationshipType).Root.Child(Drawing, "fontScheme");
                if (scheme.Exists) yield return scheme;
                part = part.RelatedPart(PowerPointPartRef.SlideMasterRelationshipType);
            }

            if (kind is PowerPointPartKind.Slide or PowerPointPartKind.SlideLayout or PowerPointPartKind.SlideMaster) {
                OfficeXmlNode scheme = part.RelatedPart(PowerPointPartRef.ThemeRelationshipType).Root
                    .Child(Drawing, "themeElements").Child(Drawing, "fontScheme");
                if (scheme.Exists) yield return scheme;
            }
        }

        private static string? ResolveColor(PowerPointPartRef owner, IReadOnlyList<OfficeXmlNode> sources) {
            for (int index = 0; index < sources.Count; index++) {
                OfficeXmlNode fill = sources[index].Child(Drawing, "solidFill");
                if (fill.Exists && ResolveRunColor(fill, owner) is { } color) return color.ToRgbHex();
            }

            return null;
        }

        private static string? ResolveTableColor(PowerPointPartRef owner, IReadOnlyList<OfficeXmlNode> sources) {
            for (int index = 0; index < sources.Count; index++) {
                OfficeXmlNode source = sources[index];
                OfficeXmlNode color = OfficeOpenXmlThemeColorResolver.FindColorElement(source);
                if (color.IsNone) color = OfficeOpenXmlThemeColorResolver.FindColorElement(source.Child(Drawing, "fontRef"));
                if (color.Exists && ResolveRunColor(color, owner) is { } resolved) return resolved.ToRgbHex();
            }

            return null;
        }

        private static OfficeColor? ResolveRunColor(OfficeXmlNode color, PowerPointPartRef owner) {
            PowerPointThemeScope scope = PowerPointThemeScope.For(owner);
            OfficeXmlNode choice = OfficeOpenXmlThemeColorResolver.FindColorElement(color);
            // Explicit dark/light slots bypass the role map for effective text styles.
            string? token = choice.Is(Drawing, "schemeClr") ? choice.Attribute("val")?.Trim().ToLowerInvariant() : null;
            OfficeXmlNode map = token is "tx1" or "tx2" or "text1" or "text2" or "bg1" or "bg2" or "background1" or "background2"
                or "accent1" or "accent2" or "accent3" or "accent4" or "accent5" or "accent6" or "hlink" or "folhlink" or "hyperlink" or "followedhyperlink"
                ? ResolveRunColorMap(owner, scope) : default;
            return OfficeOpenXmlThemeColorResolver.ResolveColor(choice, scope.ColorScheme, colorMap: map);
        }

        // Effective run styles historically use the nearest explicit mapping. Consuming-slide paint also
        // honors the slide's explicit request to use the master mapping.
        internal static OfficeXmlNode ResolveRunColorMap(PowerPointPartRef owner, PowerPointThemeScope scope) {
            OfficeXmlNode slideOverride = owner.Kind == PowerPointPartKind.Slide
                ? owner.Root.Child(Presentation, "clrMapOvr") : default;
            if (slideOverride.Child(Drawing, "masterClrMapping").IsNone) return scope.ColorMap;
            OfficeXmlNode layoutMap = owner.RelatedPart(PowerPointPartRef.SlideLayoutRelationshipType).Root
                .Child(Presentation, "clrMapOvr").Child(Drawing, "overrideClrMapping");
            return layoutMap.Exists ? layoutMap : scope.ColorMap;
        }

        private static List<OfficeXmlNode> ResolveDirectSources(
            PowerPointTextRun? run, PowerPointParagraph paragraph, OfficeXmlNode listStyle, bool includeParagraphMark) {
            OfficeXmlNode element = paragraph.Element;
            OfficeXmlNode paragraphProperties = element.Child(Drawing, "pPr");
            var sources = new List<OfficeXmlNode>(4);
            if (includeParagraphMark && element.Child(Drawing, "endParaRPr") is { Exists: true } paragraphMark)
                sources.Add(paragraphMark);
            if (run?.PropertiesElement is { Exists: true } runProperties) sources.Add(runProperties);
            if (paragraphProperties.Child(Drawing, "defRPr") is { Exists: true } defaults) sources.Add(defaults);
            int level = ParagraphLevel(paragraphProperties);
            sources.AddRange(FindDefaultRunProperties(listStyle, level));
            // Text paragraphs belong to a shape only at the text body's direct schema path.
            OfficeXmlNode body = element.Parent;
            OfficeXmlNode shape = body.Is(Presentation, "txBody") ? body.Parent : default;
            if (ShapePlaceholder(shape).Exists) {
                PowerPointPartRef owner = paragraph.OwnerPartRef;
                if (owner.IsNone) owner = run?.OwnerPartRef ?? default;
                if (owner.Kind == PowerPointPartKind.Slide) {
                    foreach (OfficeXmlNode inherited in EnumeratePlaceholderShapes(shape,
                                 owner.RelatedPart(PowerPointPartRef.SlideLayoutRelationshipType)).Skip(1)) {
                        AddPlaceholderDefaults(sources, inherited, level);
                    }
                }
            }

            return sources;
        }

        /// <summary>Finds the nearest authored highlight using the same character inheritance as fonts and emphasis.</summary>
        internal static OfficeXmlNode ResolveHighlightElement(PowerPointTextRun? run, PowerPointParagraph? paragraph,
            OfficeXmlNode listStyle, OfficeXmlNode masterTextStyle) {
            if (run == null) return default;
            if (paragraph == null) return run.PropertiesElement.Child(Drawing, "highlight");
            int level = ParagraphLevel(paragraph.Element.Child(Drawing, "pPr"));
            foreach (OfficeXmlNode properties in ResolveDirectSources(run, paragraph, listStyle, includeParagraphMark: false)
                         .Concat(FindDefaultRunProperties(masterTextStyle, level))) {
                OfficeXmlNode highlight = properties.Child(Drawing, "highlight");
                if (highlight.Exists) return highlight;
            }

            return default;
        }

        private static int ParagraphLevel(OfficeXmlNode properties) => properties.Int32Attribute("lvl") ?? 0;

        private static IEnumerable<OfficeXmlNode> FindDefaultRunProperties(OfficeXmlNode container, int level) {
            foreach (OfficeXmlNode properties in FindParagraphProperties(container, level)) {
                if (properties.Child(Drawing, "defRPr") is { Exists: true } defaults) yield return defaults;
            }
        }

        private static readonly string[] LevelProperties = {
            "lvl1pPr", "lvl2pPr", "lvl3pPr", "lvl4pPr", "lvl5pPr", "lvl6pPr", "lvl7pPr", "lvl8pPr", "lvl9pPr"
        };

        private static IEnumerable<OfficeXmlNode> FindParagraphProperties(OfficeXmlNode container, int level) {
            if (container.IsNone) yield break;
            if (level >= 0 && level < LevelProperties.Length
                && container.Child(Drawing, LevelProperties[level]) is { Exists: true } levelProperties) yield return levelProperties;
            if (container.Child(Drawing, "defPPr") is { Exists: true } fallback) yield return fallback;
        }

        private static PowerPointStrikeStyle? ToStrikeStyle(string? value) => value switch {
            "noStrike" => PowerPointStrikeStyle.None,
            "sngStrike" => PowerPointStrikeStyle.Single,
            "dblStrike" => PowerPointStrikeStyle.Double,
            _ => null
        };

        private static PowerPointCapitalization? ToCapitalization(string? value) => value switch {
            "none" => PowerPointCapitalization.None,
            "small" => PowerPointCapitalization.SmallCaps,
            "all" => PowerPointCapitalization.AllCaps,
            _ => null
        };
    }
}
