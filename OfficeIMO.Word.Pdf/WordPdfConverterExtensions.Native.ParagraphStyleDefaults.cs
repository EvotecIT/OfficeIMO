using System.Collections.Generic;
using System.Globalization;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private const int MaxNativeParagraphTabStops = 1_024;
        private readonly record struct NativeParagraphStyleDefaults(
            double? FontSize,
            string? FontFamily,
            bool? Bold,
            bool? Italic,
            bool? Underline,
            bool? Strike,
            bool? Hidden,
            bool? AllCaps,
            W.VerticalPositionValues? Baseline,
            string? ColorHex,
            W.HighlightColorValues? Highlight,
            double? LineHeight,
            double? LineSpacingPoints,
            W.LineSpacingRuleValues? LineSpacingRule,
            double? SpacingBefore,
            double? SpacingAfter,
            double? LeftIndent,
            double? RightIndent,
            double? FirstLineIndent,
            W.JustificationValues? Alignment,
            bool? PageBreakBefore,
            bool? KeepTogether,
            bool? KeepWithNext,
            bool? WidowControl,
            bool? ContextualSpacing,
            string? ShadingFillColorHex,
            NativeParagraphBorders Borders) {
            public static NativeParagraphStyleDefaults Empty { get; } = new(null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, NativeParagraphBorders.Empty);
        }

        private readonly record struct NativeParagraphBorderSide(W.BorderValues? Style, string? ColorHex, uint? Size, uint? Space) {
            public bool IsEmpty => Style == null && string.IsNullOrWhiteSpace(ColorHex) && Size == null && Space == null;
        }

        private readonly record struct NativeParagraphBorders(
            NativeParagraphBorderSide Top,
            NativeParagraphBorderSide Right,
            NativeParagraphBorderSide Bottom,
            NativeParagraphBorderSide Left) {
            public static NativeParagraphBorders Empty { get; } = new(default, default, default, default);
        }

        private readonly record struct NativeCharacterStyleDefaults(
            double? FontSize,
            string? FontFamily,
            bool? Bold,
            bool? Italic,
            bool? Underline,
            bool? Strike,
            bool? Hidden,
            bool? AllCaps,
            W.VerticalPositionValues? Baseline,
            string? ColorHex,
            W.HighlightColorValues? Highlight) {
            public static NativeCharacterStyleDefaults Empty { get; } = new(null, null, null, null, null, null, null, null, null, null, null);
        }

        private static NativeParagraphStyleDefaults GetNativeParagraphStyleDefaults(WordParagraph paragraph) {
            IReadOnlyList<W.Style> styleChain = GetNativeParagraphStyleChain(paragraph._document, paragraph.StyleId);
            if (styleChain.Count == 0) {
                return NativeParagraphStyleDefaults.Empty;
            }

            double? fontSize = null;
            string? fontFamily = null;
            bool? bold = null;
            bool? italic = null;
            bool? underline = null;
            bool? strike = null;
            bool? hidden = null;
            bool? allCaps = null;
            W.VerticalPositionValues? baseline = null;
            string? colorHex = null;
            W.HighlightColorValues? highlight = null;
            double? lineHeight = null;
            double? lineSpacingPoints = null;
            W.LineSpacingRuleValues? lineSpacingRule = null;
            double? spacingBefore = null;
            double? spacingAfter = null;
            double? leftIndent = null;
            double? rightIndent = null;
            double? firstLineIndent = null;
            W.JustificationValues? alignment = null;
            bool? pageBreakBefore = null;
            bool? keepTogether = null;
            bool? keepWithNext = null;
            bool? widowControl = null;
            bool? contextualSpacing = null;
            string? shadingFillColorHex = null;
            NativeParagraphBorders borders = NativeParagraphBorders.Empty;

            foreach (W.Style style in styleChain) {
                W.StyleRunProperties? runProperties = style.GetFirstChild<W.StyleRunProperties>();
                fontSize = GetNativeStyleFontSize(runProperties) ?? fontSize;
                fontFamily = ResolveNativeRunFontsFamily(paragraph._document, runProperties?.GetFirstChild<W.RunFonts>()) ?? fontFamily;
                bold = ReadNativeOnOff(runProperties?.GetFirstChild<W.Bold>()) ?? bold;
                italic = ReadNativeOnOff(runProperties?.GetFirstChild<W.Italic>()) ?? italic;
                underline = ReadNativeUnderline(runProperties?.GetFirstChild<W.Underline>()) ?? underline;
                strike = ReadNativeOnOff(runProperties?.GetFirstChild<W.Strike>()) ?? ReadNativeOnOff(runProperties?.GetFirstChild<W.DoubleStrike>()) ?? strike;
                hidden = ReadNativeOnOff(runProperties?.GetFirstChild<W.Vanish>()) ?? hidden;
                allCaps = ReadNativeOnOff(runProperties?.GetFirstChild<W.Caps>()) ?? ReadNativeOnOff(runProperties?.GetFirstChild<W.SmallCaps>()) ?? allCaps;
                baseline = runProperties?.GetFirstChild<W.VerticalTextAlignment>()?.Val?.Value ?? baseline;
                colorHex = runProperties?.GetFirstChild<W.Color>()?.Val?.Value ?? colorHex;
                highlight = runProperties?.GetFirstChild<W.Highlight>()?.Val?.Value ?? highlight;

                W.StyleParagraphProperties? paragraphProperties = style.GetFirstChild<W.StyleParagraphProperties>();
                if (paragraphProperties != null) {
                    W.SpacingBetweenLines? spacing = paragraphProperties.GetFirstChild<W.SpacingBetweenLines>();
                    if (spacing != null) {
                        double? styleLineHeight = GetNativeStyleParagraphLineHeight(spacing);
                        double? styleLineSpacingPoints = GetNativeStyleParagraphLineSpacingPoints(spacing);
                        if (styleLineHeight.HasValue || styleLineSpacingPoints.HasValue) {
                            lineHeight = styleLineHeight;
                            lineSpacingPoints = styleLineSpacingPoints;
                            lineSpacingRule = spacing.LineRule?.Value;
                        }

                        double effectiveFontSize = fontSize ?? NativeDocumentDefaults.WordDefault.FontSize;
                        double effectiveLineHeight = styleLineSpacingPoints.HasValue && effectiveFontSize > 0D
                            ? styleLineSpacingPoints.Value / effectiveFontSize
                            : styleLineHeight ?? lineHeight ?? NativeDocumentDefaults.WordDefault.ParagraphLineHeight;
                        spacingBefore = GetNativeSpacingBeforePoints(spacing, effectiveFontSize, effectiveLineHeight) ?? spacingBefore;
                        spacingAfter = GetNativeSpacingAfterPoints(spacing, effectiveFontSize, effectiveLineHeight) ?? spacingAfter;
                    }

                    W.Indentation? indentation = paragraphProperties.GetFirstChild<W.Indentation>();
                    if (indentation != null) {
                        leftIndent = ConvertNativeTwipsToPoints(indentation.Left?.Value) ?? leftIndent;
                        rightIndent = ConvertNativeTwipsToPoints(indentation.Right?.Value) ?? rightIndent;

                        double? firstLine = ConvertNativeTwipsToPoints(indentation.FirstLine?.Value);
                        double? hanging = ConvertNativeTwipsToPoints(indentation.Hanging?.Value);
                        if (hanging.HasValue) {
                            firstLineIndent = -hanging.Value;
                        } else if (firstLine.HasValue) {
                            firstLineIndent = firstLine.Value;
                        }
                    }

                    alignment = paragraphProperties.GetFirstChild<W.Justification>()?.Val?.Value ?? alignment;
                    pageBreakBefore = ReadNativeOnOff(paragraphProperties.GetFirstChild<W.PageBreakBefore>()) ?? pageBreakBefore;
                    keepTogether = ReadNativeOnOff(paragraphProperties.GetFirstChild<W.KeepLines>()) ?? keepTogether;
                    keepWithNext = ReadNativeOnOff(paragraphProperties.GetFirstChild<W.KeepNext>()) ?? keepWithNext;
                    widowControl = ReadNativeOnOff(paragraphProperties.GetFirstChild<W.WidowControl>()) ?? widowControl;
                    contextualSpacing = ReadNativeOnOff(paragraphProperties.GetFirstChild<W.ContextualSpacing>()) ?? contextualSpacing;
                    shadingFillColorHex = NormalizeNativeShadingFill(paragraphProperties.GetFirstChild<W.Shading>()?.Fill?.Value) ?? shadingFillColorHex;
                    borders = MergeNativeParagraphBorders(borders, paragraphProperties.GetFirstChild<W.ParagraphBorders>());
                }
            }

            return new NativeParagraphStyleDefaults(
                fontSize,
                fontFamily,
                bold,
                italic,
                underline,
                strike,
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
                borders);
        }

        private static IReadOnlyList<W.Style> GetNativeParagraphStyleChain(WordDocument? document, string? styleId) {
            W.Styles? styles = document?._wordprocessingDocument?.MainDocumentPart?.StyleDefinitionsPart?.Styles;
            if (styles == null) {
                return Array.Empty<W.Style>();
            }

            Dictionary<string, W.Style> paragraphStyles = styles
                .Elements<W.Style>()
                .Where(style => IsNativeParagraphStyle(style) && !string.IsNullOrEmpty(style.StyleId?.Value))
                .ToDictionary(style => style.StyleId!.Value!, style => style, StringComparer.Ordinal);

            if (string.IsNullOrWhiteSpace(styleId)) {
                styleId = paragraphStyles.Values.FirstOrDefault(style => style.Default?.Value == true)?.StyleId?.Value;
            }

            if (string.IsNullOrWhiteSpace(styleId)) {
                return Array.Empty<W.Style>();
            }

            var chain = new List<W.Style>();
            var visited = new HashSet<string>(StringComparer.Ordinal);
            string? currentStyleId = styleId;
            while (!string.IsNullOrWhiteSpace(currentStyleId) && visited.Add(currentStyleId!) && paragraphStyles.TryGetValue(currentStyleId!, out W.Style? style)) {
                chain.Add(style);
                currentStyleId = style.BasedOn?.Val?.Value;
            }

            chain.Reverse();
            return chain;
        }

        private static NativeCharacterStyleDefaults GetNativeCharacterStyleDefaults(WordDocument? document, W.RunProperties? runProperties) {
            string? styleId = runProperties?.RunStyle?.Val?.Value;
            IReadOnlyList<W.Style> styleChain = GetNativeCharacterStyleChain(document, styleId);
            if (styleChain.Count == 0) {
                return NativeCharacterStyleDefaults.Empty;
            }

            double? fontSize = null;
            string? fontFamily = null;
            bool? bold = null;
            bool? italic = null;
            bool? underline = null;
            bool? strike = null;
            bool? hidden = null;
            bool? allCaps = null;
            W.VerticalPositionValues? baseline = null;
            string? colorHex = null;
            W.HighlightColorValues? highlight = null;

            foreach (W.Style style in styleChain) {
                W.StyleRunProperties? styleRunProperties = style.GetFirstChild<W.StyleRunProperties>();
                fontSize = GetNativeStyleFontSize(styleRunProperties) ?? fontSize;
                fontFamily = ResolveNativeRunFontsFamily(document, styleRunProperties?.GetFirstChild<W.RunFonts>()) ?? fontFamily;
                bold = ReadNativeOnOff(styleRunProperties?.GetFirstChild<W.Bold>()) ?? bold;
                italic = ReadNativeOnOff(styleRunProperties?.GetFirstChild<W.Italic>()) ?? italic;
                underline = ReadNativeUnderline(styleRunProperties?.GetFirstChild<W.Underline>()) ?? underline;
                strike = ReadNativeOnOff(styleRunProperties?.GetFirstChild<W.Strike>()) ?? ReadNativeOnOff(styleRunProperties?.GetFirstChild<W.DoubleStrike>()) ?? strike;
                hidden = ReadNativeOnOff(styleRunProperties?.GetFirstChild<W.Vanish>()) ?? hidden;
                allCaps = ReadNativeOnOff(styleRunProperties?.GetFirstChild<W.Caps>()) ?? ReadNativeOnOff(styleRunProperties?.GetFirstChild<W.SmallCaps>()) ?? allCaps;
                baseline = styleRunProperties?.GetFirstChild<W.VerticalTextAlignment>()?.Val?.Value ?? baseline;
                colorHex = styleRunProperties?.GetFirstChild<W.Color>()?.Val?.Value ?? colorHex;
                highlight = styleRunProperties?.GetFirstChild<W.Highlight>()?.Val?.Value ?? highlight;
            }

            return new NativeCharacterStyleDefaults(
                fontSize,
                fontFamily,
                bold,
                italic,
                underline,
                strike,
                hidden,
                allCaps,
                baseline,
                colorHex,
                highlight);
        }

        private static IReadOnlyList<W.Style> GetNativeCharacterStyleChain(WordDocument? document, string? styleId) {
            W.Styles? styles = document?._wordprocessingDocument?.MainDocumentPart?.StyleDefinitionsPart?.Styles;
            if (styles == null || string.IsNullOrWhiteSpace(styleId)) {
                return Array.Empty<W.Style>();
            }

            Dictionary<string, W.Style> characterStyles = styles
                .Elements<W.Style>()
                .Where(style => IsNativeCharacterStyle(style) && !string.IsNullOrEmpty(style.StyleId?.Value))
                .ToDictionary(style => style.StyleId!.Value!, style => style, StringComparer.Ordinal);

            var chain = new List<W.Style>();
            var visited = new HashSet<string>(StringComparer.Ordinal);
            string? currentStyleId = styleId;
            while (!string.IsNullOrWhiteSpace(currentStyleId) && visited.Add(currentStyleId!) && characterStyles.TryGetValue(currentStyleId!, out W.Style? style)) {
                chain.Add(style);
                currentStyleId = style.BasedOn?.Val?.Value;
            }

            chain.Reverse();
            return chain;
        }

        private static IReadOnlyList<WordTabStop> GetNativeParagraphEffectiveTabStops(WordParagraph paragraph) {
            W.Tabs? directTabs = paragraph._paragraphProperties?.Tabs;
            if (directTabs != null) {
                WordTabStop[] directTabStops = directTabs.Elements<W.TabStop>()
                    .Take(MaxNativeParagraphTabStops)
                    .Select(tabStop => new WordTabStop(paragraph, tabStop))
                    .ToArray();
                if (directTabStops.Length > 0) {
                    return directTabStops;
                }
            }

            var styleTabStops = new Dictionary<int, WordTabStop>();
            int inspectedStyleTabStops = 0;
            foreach (W.Style style in GetNativeParagraphStyleChain(paragraph._document, paragraph.StyleId).Reverse()) {
                W.Tabs? tabs = style.GetFirstChild<W.StyleParagraphProperties>()?.GetFirstChild<W.Tabs>();
                if (tabs == null) {
                    continue;
                }

                foreach (W.TabStop tabStop in tabs.Elements<W.TabStop>()) {
                    if (inspectedStyleTabStops >= MaxNativeParagraphTabStops) {
                        break;
                    }

                    inspectedStyleTabStops++;
                    WordTabStop wordTabStop = new WordTabStop(paragraph, (W.TabStop)tabStop.CloneNode(true));
                    if (wordTabStop.Position <= 0 || !IsNativeRenderableTextTabStop(wordTabStop.Alignment)) {
                        continue;
                    }

                    if (!styleTabStops.ContainsKey(wordTabStop.Position)) {
                        styleTabStops.Add(wordTabStop.Position, wordTabStop);
                    }
                    if (styleTabStops.Count >= MaxNativeParagraphTabStops) {
                        break;
                    }
                }

                if (inspectedStyleTabStops >= MaxNativeParagraphTabStops) {
                    break;
                }
            }

            if (styleTabStops.Count == 0) {
                return Array.Empty<WordTabStop>();
            }

            return styleTabStops.Values.OrderBy(tabStop => tabStop.Position).ToArray();
        }

        private static bool IsNativeParagraphStyle(W.Style style) {
            if (style.Type == null) {
                return false;
            }

            string? type = string.IsNullOrWhiteSpace(style.Type.InnerText)
                ? style.Type.Value.ToString()
                : style.Type.InnerText;
            return string.Equals(type, "paragraph", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsNativeCharacterStyle(W.Style style) {
            if (style.Type == null) {
                return false;
            }

            string? type = string.IsNullOrWhiteSpace(style.Type.InnerText)
                ? style.Type.Value.ToString()
                : style.Type.InnerText;
            return string.Equals(type, "character", StringComparison.OrdinalIgnoreCase);
        }

        private static NativeParagraphBorders MergeNativeParagraphBorders(NativeParagraphBorders current, W.ParagraphBorders? borders) {
            if (borders == null) {
                return current;
            }

            return current with {
                Top = ReadNativeParagraphBorderSide(borders.TopBorder) ?? current.Top,
                Right = ReadNativeParagraphBorderSide(borders.RightBorder) ?? current.Right,
                Bottom = ReadNativeParagraphBorderSide(borders.BottomBorder) ?? current.Bottom,
                Left = ReadNativeParagraphBorderSide(borders.LeftBorder) ?? current.Left
            };
        }

        private static NativeParagraphBorderSide? ReadNativeParagraphBorderSide(W.BorderType? border) {
            if (border == null) {
                return null;
            }

            return new NativeParagraphBorderSide(
                border.Val?.Value,
                NormalizeNativeBorderColor(border.Color?.Value),
                border.Size?.Value,
                border.Space?.Value);
        }

        private static double? GetNativeStyleParagraphLineHeight(W.SpacingBetweenLines spacing) {
            if (spacing.LineRule?.Value != W.LineSpacingRuleValues.Auto) {
                return null;
            }

            if (!double.TryParse(spacing.Line?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double line) ||
                line <= 0D ||
                double.IsNaN(line) ||
                double.IsInfinity(line)) {
                return null;
            }

            return Math.Max(0.01D, NativeWordAutoLineSpacingHeight * (line / 240D));
        }

        private static double? GetNativeStyleParagraphLineSpacingPoints(W.SpacingBetweenLines spacing) {
            if (spacing.LineRule?.Value == W.LineSpacingRuleValues.Auto) {
                return null;
            }

            return ConvertNativeTwipsToPoints(spacing.Line?.Value);
        }

        private static double? GetNativeStyleFontSize(W.StyleRunProperties? runProperties) {
            string? value = runProperties?.FontSize?.Val?.Value;
            if (!double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double halfPoints) ||
                halfPoints <= 0D ||
                double.IsNaN(halfPoints) ||
                double.IsInfinity(halfPoints)) {
                return null;
            }

            return halfPoints / 2D;
        }

        private static bool? ReadNativeOnOff(W.OnOffType? value) {
            if (value == null) {
                return null;
            }

            return value.Val?.Value != false;
        }

        private static bool? ReadNativeUnderline(W.Underline? value) {
            if (value == null) {
                return null;
            }

            return value.Val?.Value != W.UnderlineValues.None;
        }

        private static bool? ReadNativeDirectParagraphOnOff<T>(WordParagraph paragraph) where T : W.OnOffType =>
            ReadNativeOnOff(paragraph._paragraph?.ParagraphProperties?.GetFirstChild<T>());

        private static bool HasNativePageBreakBefore(WordParagraph paragraph) =>
            paragraph.PageBreakBefore ||
            GetNativeParagraphStyleDefaults(paragraph).PageBreakBefore == true;
    }
}
