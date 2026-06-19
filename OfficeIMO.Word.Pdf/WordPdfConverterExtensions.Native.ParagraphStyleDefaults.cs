using System.Collections.Generic;
using System.Globalization;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private readonly record struct NativeParagraphStyleDefaults(
            double? FontSize,
            double? LineHeight,
            double? SpacingBefore,
            double? SpacingAfter,
            bool? KeepTogether,
            bool? KeepWithNext,
            bool? WidowControl) {
            public static NativeParagraphStyleDefaults Empty { get; } = new(null, null, null, null, null, null, null);
        }

        private static NativeParagraphStyleDefaults GetNativeParagraphStyleDefaults(WordParagraph paragraph) {
            IReadOnlyList<W.Style> styleChain = GetNativeParagraphStyleChain(paragraph._document, paragraph.StyleId);
            if (styleChain.Count == 0) {
                return NativeParagraphStyleDefaults.Empty;
            }

            double? fontSize = null;
            double? lineHeight = null;
            double? spacingBefore = null;
            double? spacingAfter = null;
            bool? keepTogether = null;
            bool? keepWithNext = null;
            bool? widowControl = null;

            foreach (W.Style style in styleChain) {
                W.StyleParagraphProperties? paragraphProperties = style.GetFirstChild<W.StyleParagraphProperties>();
                if (paragraphProperties != null) {
                    W.SpacingBetweenLines? spacing = paragraphProperties.GetFirstChild<W.SpacingBetweenLines>();
                    if (spacing != null) {
                        lineHeight = GetNativeStyleParagraphLineHeight(spacing) ?? lineHeight;
                        spacingBefore = ConvertNativeTwipsToPoints(spacing.Before?.Value) ?? spacingBefore;
                        spacingAfter = ConvertNativeTwipsToPoints(spacing.After?.Value) ?? spacingAfter;
                    }

                    keepTogether = ReadNativeOnOff(paragraphProperties.GetFirstChild<W.KeepLines>()) ?? keepTogether;
                    keepWithNext = ReadNativeOnOff(paragraphProperties.GetFirstChild<W.KeepNext>()) ?? keepWithNext;
                    widowControl = ReadNativeOnOff(paragraphProperties.GetFirstChild<W.WidowControl>()) ?? widowControl;
                }

                fontSize = GetNativeStyleFontSize(style.GetFirstChild<W.StyleRunProperties>()) ?? fontSize;
            }

            return new NativeParagraphStyleDefaults(
                fontSize,
                lineHeight,
                spacingBefore,
                spacingAfter,
                keepTogether,
                keepWithNext,
                widowControl);
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

        private static bool IsNativeParagraphStyle(W.Style style) {
            if (style.Type == null) {
                return false;
            }

            string? type = string.IsNullOrWhiteSpace(style.Type.InnerText)
                ? style.Type.Value.ToString()
                : style.Type.InnerText;
            return string.Equals(type, "paragraph", StringComparison.OrdinalIgnoreCase);
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

        private static bool? ReadNativeDirectParagraphOnOff<T>(WordParagraph paragraph) where T : W.OnOffType =>
            ReadNativeOnOff(paragraph._paragraph?.ParagraphProperties?.GetFirstChild<T>());
    }
}
