using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Word.Pdf {

    public static partial class WordPdfConverterExtensions {
        private const int MaxNativeParagraphTabStops = 1_024;
        private readonly record struct NativeParagraphStyleDefaults(
            double? FontSize,
            string? FontFamily,
            bool? Bold,
            bool? Italic,
            OfficeTextDecorationStyle? UnderlineStyle,
            OfficeTextDecorationStyle? StrikeStyle,
            bool? Hidden,
            bool? AllCaps,
            string? Baseline,
            string? ColorHex,
            string? Highlight,
            double? LineHeight,
            double? LineSpacingPoints,
            string? LineSpacingRule,
            double? SpacingBefore,
            double? SpacingAfter,
            double? LeftIndent,
            double? RightIndent,
            double? FirstLineIndent,
            string? Alignment,
            bool? PageBreakBefore,
            bool? KeepTogether,
            bool? KeepWithNext,
            bool? WidowControl,
            bool? ContextualSpacing,
            string? ShadingFillColorHex,
            NativeParagraphBorders Borders) {
            public bool? SmallCaps { get; init; }
            public bool? BiDi { get; init; }
            public NativeLatinFontFamilies FontFamilies { get; init; }
            public NativeComplexScriptDefaults ComplexScript { get; init; }
            public NativeTextSpacing TextSpacing { get; init; }
            public NativeLineSpacing LineSpacing { get; init; }
            public NativeParagraphSpacing ParagraphSpacing { get; init; }
            public static NativeParagraphStyleDefaults Empty { get; } = new(null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, NativeParagraphBorders.Empty);
        }

        private readonly record struct NativeParagraphBorderSide(string? Style, string? ColorHex, uint? Size, uint? Space) {
            public bool Shadow { get; init; }
            public bool Frame { get; init; }
            public bool IsEmpty => Style == null && string.IsNullOrWhiteSpace(ColorHex) && Size == null && Space == null;
        }

        private readonly record struct NativeParagraphBorders(
            NativeParagraphBorderSide Top,
            NativeParagraphBorderSide Right,
            NativeParagraphBorderSide Bottom,
            NativeParagraphBorderSide Left) {
            public NativeParagraphBorderSide Between { get; init; }
            public static NativeParagraphBorders Empty { get; } = new(default, default, default, default);
        }

        private readonly record struct NativeCharacterStyleDefaults(
            double? FontSize,
            string? FontFamily,
            bool? Bold,
            bool? Italic,
            OfficeTextDecorationStyle? UnderlineStyle,
            OfficeTextDecorationStyle? StrikeStyle,
            bool? Hidden,
            bool? AllCaps,
            string? Baseline,
            string? ColorHex,
            string? Highlight) {
            public bool? SmallCaps { get; init; }
            public NativeLatinFontFamilies FontFamilies { get; init; }
            public NativeComplexScriptDefaults ComplexScript { get; init; }
            public NativeTextSpacing TextSpacing { get; init; }
            public static NativeCharacterStyleDefaults Empty { get; } = new(null, null, null, null, null, null, null, null, null, null, null);
        }

        private static NativeParagraphStyleDefaults GetNativeParagraphStyleDefaults(WordParagraph paragraph) {
            NativeStyleLookupCache? cache = GetNativeStyleLookupCache(paragraph._document);
            string? resolvedStyleId = cache?.ResolveParagraphStyleId(paragraph.StyleId);
            if (cache != null && !string.IsNullOrWhiteSpace(resolvedStyleId) && cache.ParagraphDefaults.TryGetValue(resolvedStyleId!, out NativeParagraphStyleDefaults cachedDefaults)) {
                return cachedDefaults;
            }

            // The document default direction applies even without a paragraph style, as with a missing style.
            bool? documentBidi = ReadNativeOnOff(WordStyleCatalog.For(paragraph._document).DocumentDefaultsNode().WChild("pPrDefault").WChild("pPr").WChild("bidi"));
            IReadOnlyList<OfficeXmlNode> styleChain = GetNativeParagraphStyleChain(paragraph._document, paragraph.StyleId);
            if (styleChain.Count == 0) {
                return NativeParagraphStyleDefaults.Empty with { BiDi = documentBidi };
            }

            NativeDocumentDefaults documentDefaults = GetNativeDocumentDefaults(paragraph._document);
            NativeComplexScriptDefaults complexScript = default;
            NativeTextSpacing textSpacing = default;
            double? fontSize = null;
            string? fontFamily = null;
            NativeLatinFontFamilies fontFamilies = default;
            bool? bold = null;
            bool? italic = null;
            OfficeTextDecorationStyle? underlineStyle = null;
            OfficeTextDecorationStyle? strikeStyle = null;
            bool? hidden = null;
            bool? allCaps = null;
            bool? smallCaps = null;
            string? baseline = null;
            string? colorHex = null;
            string? highlight = null;
            double? lineHeight = null;
            double? lineSpacingPoints = null;
            string? lineSpacingRule = null;
            NativeLineSpacing authoredLineSpacing = default;
            double? spacingBefore = null;
            double? spacingAfter = null;
            double? leftIndent = null;
            double? rightIndent = null;
            double? firstLineIndent = null;
            string? alignment = null;
            bool? pageBreakBefore = null;
            bool? keepTogether = null;
            bool? keepWithNext = null;
            bool? widowControl = null;
            bool? contextualSpacing = null;
            bool? bidi = documentBidi;
            NativeParagraphSpacing paragraphSpacing = default;
            string? shadingFillColorHex = null;
            NativeParagraphBorders borders = NativeParagraphBorders.Empty;

            foreach (OfficeXmlNode style in styleChain) {
                OfficeXmlNode runProperties = style.WChild("rPr");
                complexScript = complexScript.Merge(runProperties);
                textSpacing = textSpacing.Merge(runProperties);
                fontSize = GetNativeStyleFontSize(runProperties) ?? fontSize;
                fontFamilies = GetNativeRunFontFamilies(paragraph._document, runProperties.WChild("rFonts")).Inherit(fontFamilies);
                fontFamily = fontFamilies.Primary;
                bold = ReadNativeOnOff(runProperties.WChild("b")) ?? bold;
                italic = ReadNativeOnOff(runProperties.WChild("i")) ?? italic;
                underlineStyle = MapNativeUnderlineStyle(runProperties.WChild("u")) ?? underlineStyle;
                strikeStyle = MapNativeStrikeStyle(runProperties) ?? strikeStyle;
                hidden = ReadNativeOnOff(runProperties.WChild("vanish")) ?? hidden;
                allCaps = ReadNativeOnOff(runProperties.WChild("caps")) ?? allCaps;
                smallCaps = ReadNativeOnOff(runProperties.WChild("smallCaps")) ?? smallCaps;
                baseline = runProperties.WChild("vertAlign").WVal() ?? baseline;
                colorHex = WordConverterNodes.ResolveColorHex(runProperties.WChild("color"), paragraph._document) ?? colorHex;
                highlight = runProperties.WChild("highlight").WVal() ?? highlight;

                OfficeXmlNode paragraphProperties = style.WChild("pPr");
                paragraphSpacing = paragraphSpacing.Merge(paragraphProperties);
                if (paragraphProperties.Exists) {
                    OfficeXmlNode spacing = paragraphProperties.WChild("spacing");
                    authoredLineSpacing = ReadNativeLineSpacing(spacing, paragraphProperties.WChild("textAlignment")).Inherit(authoredLineSpacing);
                    if (spacing.Exists) {
                        double? styleLineHeight = GetNativeStyleParagraphLineHeight(
                            spacing,
                            fontFamily,
                            documentDefaults.FontFamily);
                        double? styleLineSpacingPoints = GetNativeStyleParagraphLineSpacingPoints(spacing);
                        if (styleLineHeight.HasValue || styleLineSpacingPoints.HasValue) {
                            lineHeight = styleLineHeight;
                            lineSpacingPoints = styleLineSpacingPoints;
                            lineSpacingRule = spacing.WAttr("lineRule");
                        }

                        double effectiveFontSize = fontSize ?? documentDefaults.FontSize;
                        double effectiveLineHeight = styleLineSpacingPoints.HasValue && effectiveFontSize > 0D
                            ? styleLineSpacingPoints.Value / effectiveFontSize
                            : styleLineHeight ?? lineHeight ?? NativeDocumentDefaults.WordDefault.ParagraphLineHeight;
                        spacingBefore = GetNativeSpacingBeforePoints(spacing, effectiveFontSize, effectiveLineHeight) ?? spacingBefore;
                        spacingAfter = GetNativeSpacingAfterPoints(spacing, effectiveFontSize, effectiveLineHeight) ?? spacingAfter;
                    }

                    OfficeXmlNode indentation = paragraphProperties.WChild("ind");
                    if (indentation.Exists) {
                        leftIndent = ConvertNativeTwipsToPoints(indentation.WAttr("left")) ?? leftIndent;
                        rightIndent = ConvertNativeTwipsToPoints(indentation.WAttr("right")) ?? rightIndent;

                        double? firstLine = ConvertNativeTwipsToPoints(indentation.WAttr("firstLine"));
                        double? hanging = ConvertNativeTwipsToPoints(indentation.WAttr("hanging"));
                        if (hanging.HasValue) {
                            firstLineIndent = -hanging.Value;
                        } else if (firstLine.HasValue) {
                            firstLineIndent = firstLine.Value;
                        }
                    }

                    alignment = paragraphProperties.WChild("jc").WVal() ?? alignment;
                    pageBreakBefore = ReadNativeOnOff(paragraphProperties.WChild("pageBreakBefore")) ?? pageBreakBefore;
                    keepTogether = ReadNativeOnOff(paragraphProperties.WChild("keepLines")) ?? keepTogether;
                    keepWithNext = ReadNativeOnOff(paragraphProperties.WChild("keepNext")) ?? keepWithNext;
                    widowControl = ReadNativeOnOff(paragraphProperties.WChild("widowControl")) ?? widowControl;
                    contextualSpacing = ReadNativeOnOff(paragraphProperties.WChild("contextualSpacing")) ?? contextualSpacing;
                    bidi = ReadNativeOnOff(paragraphProperties.WChild("bidi")) ?? bidi;
                    if (paragraphProperties.WChild("shd") is { Exists: true } shading) {
                        // A derived style's automatic fill clears the inherited fill.
                        shadingFillColorHex = NormalizeNativeShadingFill(shading.WAttr("fill"));
                    }
                    borders = MergeNativeParagraphBorders(borders, paragraphProperties.WChild("pBdr"));
                }
            }

            var result = new NativeParagraphStyleDefaults(
                fontSize,
                fontFamily,
                bold,
                italic,
                underlineStyle,
                strikeStyle,
                hidden,
                allCaps,
                baseline,
                colorHex,
                highlight,
                lineHeight,
                lineSpacingPoints,
                lineSpacingRule,
                spacingBefore,
                spacingAfter,
                leftIndent,
                rightIndent,
                firstLineIndent,
                alignment,
                pageBreakBefore,
                keepTogether,
                keepWithNext,
                widowControl,
                contextualSpacing,
                shadingFillColorHex,
                borders) { SmallCaps = smallCaps, BiDi = bidi, FontFamilies = fontFamilies, ComplexScript = complexScript, TextSpacing = textSpacing, LineSpacing = authoredLineSpacing, ParagraphSpacing = paragraphSpacing };
            if (cache != null && !string.IsNullOrWhiteSpace(resolvedStyleId)) {
                cache.ParagraphDefaults[resolvedStyleId!] = result;
            }

            return result;
        }

        private static IReadOnlyList<OfficeXmlNode> GetNativeParagraphStyleChain(WordDocument? document, string? styleId) {
            NativeStyleLookupCache? cache = GetNativeStyleLookupCache(document);
            string? resolvedStyleId = cache?.ResolveParagraphStyleId(styleId);
            if (cache == null || string.IsNullOrWhiteSpace(resolvedStyleId)) {
                return Array.Empty<OfficeXmlNode>();
            }

            if (cache.ParagraphChains.TryGetValue(resolvedStyleId!, out IReadOnlyList<OfficeXmlNode>? cachedChain)) {
                return cachedChain;
            }

            var chain = new List<OfficeXmlNode>();
            var visited = new HashSet<string>(StringComparer.Ordinal);
            string? currentStyleId = resolvedStyleId;
            while (!string.IsNullOrWhiteSpace(currentStyleId) && visited.Add(currentStyleId!) && cache.ParagraphStyles.TryGetValue(currentStyleId!, out OfficeXmlNode style)) {
                cache.RecordStyleChainReference(chain.Count + 1);
                chain.Add(style);
                currentStyleId = style.ChildVal("basedOn");
            }

            chain.Reverse();
            IReadOnlyList<OfficeXmlNode> result = chain.ToArray();
            cache.ParagraphChains[resolvedStyleId!] = result;
            return result;
        }

        private static NativeCharacterStyleDefaults GetNativeCharacterStyleDefaults(WordDocument? document, OfficeXmlNode runProperties) {
            string? styleId = runProperties.WChild("rStyle").WVal();
            NativeStyleLookupCache? cache = GetNativeStyleLookupCache(document);
            if (cache != null && !string.IsNullOrWhiteSpace(styleId) && cache.CharacterDefaults.TryGetValue(styleId!, out NativeCharacterStyleDefaults cachedDefaults)) {
                return cachedDefaults;
            }

            IReadOnlyList<OfficeXmlNode> styleChain = GetNativeCharacterStyleChain(document, styleId);
            if (styleChain.Count == 0) {
                return NativeCharacterStyleDefaults.Empty;
            }

            NativeComplexScriptDefaults complexScript = default;
            NativeTextSpacing textSpacing = default;
            double? fontSize = null;
            string? fontFamily = null;
            NativeLatinFontFamilies fontFamilies = default;
            bool? bold = null;
            bool? italic = null;
            OfficeTextDecorationStyle? underlineStyle = null;
            OfficeTextDecorationStyle? strikeStyle = null;
            bool? hidden = null;
            bool? allCaps = null;
            bool? smallCaps = null;
            string? baseline = null;
            string? colorHex = null;
            string? highlight = null;

            foreach (OfficeXmlNode style in styleChain) {
                OfficeXmlNode styleRunProperties = style.WChild("rPr");
                complexScript = complexScript.Merge(styleRunProperties);
                textSpacing = textSpacing.Merge(styleRunProperties);
                fontSize = GetNativeStyleFontSize(styleRunProperties) ?? fontSize;
                fontFamilies = GetNativeRunFontFamilies(document, styleRunProperties.WChild("rFonts")).Inherit(fontFamilies);
                fontFamily = fontFamilies.Primary;
                bold = ReadNativeOnOff(styleRunProperties.WChild("b")) ?? bold;
                italic = ReadNativeOnOff(styleRunProperties.WChild("i")) ?? italic;
                underlineStyle = MapNativeUnderlineStyle(styleRunProperties.WChild("u")) ?? underlineStyle;
                strikeStyle = MapNativeStrikeStyle(styleRunProperties) ?? strikeStyle;
                hidden = ReadNativeOnOff(styleRunProperties.WChild("vanish")) ?? hidden;
                allCaps = ReadNativeOnOff(styleRunProperties.WChild("caps")) ?? allCaps;
                smallCaps = ReadNativeOnOff(styleRunProperties.WChild("smallCaps")) ?? smallCaps;
                baseline = styleRunProperties.WChild("vertAlign").WVal() ?? baseline;
                colorHex = WordConverterNodes.ResolveColorHex(styleRunProperties.WChild("color"), document) ?? colorHex;
                highlight = styleRunProperties.WChild("highlight").WVal() ?? highlight;
            }

            var result = new NativeCharacterStyleDefaults(
                fontSize,
                fontFamily,
                bold,
                italic,
                underlineStyle,
                strikeStyle,
                hidden,
                allCaps,
                baseline,
                colorHex,
                highlight) { SmallCaps = smallCaps, FontFamilies = fontFamilies, ComplexScript = complexScript, TextSpacing = textSpacing };
            if (cache != null && !string.IsNullOrWhiteSpace(styleId)) {
                cache.CharacterDefaults[styleId!] = result;
            }

            return result;
        }

        private static IReadOnlyList<OfficeXmlNode> GetNativeCharacterStyleChain(WordDocument? document, string? styleId) {
            NativeStyleLookupCache? cache = GetNativeStyleLookupCache(document);
            if (cache == null || string.IsNullOrWhiteSpace(styleId)) {
                return Array.Empty<OfficeXmlNode>();
            }

            if (cache.CharacterChains.TryGetValue(styleId!, out IReadOnlyList<OfficeXmlNode>? cachedChain)) {
                return cachedChain;
            }

            var chain = new List<OfficeXmlNode>();
            var visited = new HashSet<string>(StringComparer.Ordinal);
            string? currentStyleId = styleId;
            while (!string.IsNullOrWhiteSpace(currentStyleId) && visited.Add(currentStyleId!) && cache.CharacterStyles.TryGetValue(currentStyleId!, out OfficeXmlNode style)) {
                cache.RecordStyleChainReference(chain.Count + 1);
                chain.Add(style);
                currentStyleId = style.ChildVal("basedOn");
            }

            chain.Reverse();
            IReadOnlyList<OfficeXmlNode> result = chain.ToArray();
            cache.CharacterChains[styleId!] = result;
            return result;
        }

        private static IReadOnlyList<WordTabStop> GetNativeParagraphEffectiveTabStops(WordParagraph paragraph,
            OfficeXmlNode numberingTabs = default) {
            if (numberingTabs.IsNone && WordDocumentTraversal.GetListInfo(paragraph) is { } listInfo) {
                numberingTabs = listInfo.LevelTabStopsNode();
            }
            var tabStops = new Dictionary<OfficeLength, WordTabStop>();
            int inspectedTabStops = 0;
            // Word tab lists compose across the hierarchy. Visit the most
            // specific declaration first, including clears that mask ancestors.
            AddTabs(paragraph.ParagraphPropertiesNode().WChild("tabs"));
            OfficeXmlNode directNumbering = paragraph.ParagraphPropertiesNode().WChild("numPr");
            bool hasDirectNumbering = directNumbering.Exists
                && (WordXml.ParseInt32(directNumbering.ChildVal("numId")) is > 0
                    || WordXml.ParseInt32(directNumbering.ChildVal("ilvl")) is >= 0);
            // A directly applied numbering definition precedes paragraph styles;
            // style-inherited numbering remains below the style declarations.
            if (hasDirectNumbering) AddTabs(numberingTabs);
            foreach (OfficeXmlNode style in GetNativeParagraphStyleChain(paragraph._document, paragraph.StyleId).Reverse()) {
                if (inspectedTabStops >= MaxNativeParagraphTabStops) {
                    break;
                }

                AddTabs(style.WChild("pPr").WChild("tabs"));
            }

            if (!hasDirectNumbering) AddTabs(numberingTabs);
            if (paragraph._document != null) {
                AddTabs(WordStyleCatalog.For(paragraph._document).DocumentDefaultsNode()
                    .WChild("pPrDefault").WChild("pPr").WChild("tabs"));
            }
            return tabStops.Values.OrderBy(tabStop => tabStop.Position).ToArray();

            void AddTabs(OfficeXmlNode tabs) {
                if (tabs.IsNone) {
                    return;
                }

                foreach (OfficeXmlNode tabStop in tabs.WElements("tab")) {
                    if (inspectedTabStops >= MaxNativeParagraphTabStops) {
                        break;
                    }

                    inspectedTabStops++;
                    // A stop without a readable position has no place on the line.
                    if (tabStop.WInt32("pos") == null) {
                        continue;
                    }

                    WordTabStop wordTabStop = WordConverterNodes.DetachedTabStop(paragraph, tabStop);
                    if (!tabStops.ContainsKey(wordTabStop.Position)) {
                        tabStops.Add(wordTabStop.Position, wordTabStop);
                    }
                }
            }
        }

        private static NativeParagraphBorders MergeNativeParagraphBorders(NativeParagraphBorders current, OfficeXmlNode borders) {
            if (borders.IsNone) {
                return current;
            }

            return current with {
                Top = ReadNativeParagraphBorderSide(borders.WChild("top")) ?? current.Top,
                Right = ReadNativeParagraphBorderSide(borders.WChild("right")) ?? current.Right,
                Bottom = ReadNativeParagraphBorderSide(borders.WChild("bottom")) ?? current.Bottom,
                Left = ReadNativeParagraphBorderSide(borders.WChild("left")) ?? current.Left,
                Between = ReadNativeParagraphBorderSide(borders.WChild("between")) ?? current.Between
            };
        }

        private static NativeParagraphBorderSide? ReadNativeParagraphBorderSide(OfficeXmlNode border) {
            if (border.IsNone) {
                return null;
            }

            return new NativeParagraphBorderSide(
                border.WVal(),
                NormalizeNativeBorderColor(border.WAttr("color")),
                border.WUInt32("sz"),
                border.WUInt32("space")) {
                Shadow = WordXml.ParseOnOff(border.WAttr("shadow")) == true,
                Frame = WordXml.ParseOnOff(border.WAttr("frame")) == true
            };
        }

        private static double? GetNativeStyleParagraphLineHeight(
            OfficeXmlNode spacing,
            params string?[] fontFamilies) {
            if (ReadNativeWordLineSpacing(spacing)?.MultipleValue is not double multiple || multiple <= 0D) return null;
            return Math.Max(0.01D, ResolveNativeWordSingleLineHeight(fontFamilies) * multiple);
        }

        private static double? GetNativeStyleParagraphLineSpacingPoints(OfficeXmlNode spacing) {
            return ReadNativeWordLineSpacing(spacing)?.Height?.ToPoints();
        }

        /// <summary>A <c>w:spacing</c> line value and rule as written; an unknown rule reads as no spacing.</summary>
        private static WordLineSpacing? ReadNativeWordLineSpacing(OfficeXmlNode spacing) =>
            spacing.IsNone ? null : WordLineSpacing.Read(spacing.WAttr("line"), spacing.WAttr("lineRule"));

        private static double? GetNativeStyleFontSize(OfficeXmlNode runProperties) {
            string? value = runProperties.WChild("sz").WVal();
            if (!OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double halfPoints) ||
                halfPoints <= 0D ||
                double.IsNaN(halfPoints) ||
                double.IsInfinity(halfPoints)) {
                return null;
            }

            return halfPoints / 2D;
        }

        private static bool? ReadNativeOnOff(OfficeXmlNode value) => value.OnOff();

        private static bool? ReadNativeDirectParagraphOnOff(WordParagraph paragraph, string localName) =>
            ReadNativeOnOff(paragraph.Node().WChild("pPr").WChild(localName));

        private static bool HasNativePageBreakBefore(WordParagraph paragraph) =>
            ReadNativeDirectParagraphOnOff(paragraph, "pageBreakBefore") ??
            GetNativeParagraphStyleDefaults(paragraph).PageBreakBefore ?? false;
    }
}
