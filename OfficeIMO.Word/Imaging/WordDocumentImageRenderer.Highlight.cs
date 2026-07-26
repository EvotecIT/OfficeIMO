using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Word {
    internal static partial class WordDocumentImageRenderer {
        private static bool HasRunBackground(WordParagraph paragraph, A.ColorScheme? colorScheme) =>
            ResolveRunBackgroundColor(paragraph, colorScheme).HasValue;

        private static OfficeColor? ResolveRunBackgroundColor(
            WordParagraph paragraph,
            A.ColorScheme? colorScheme) {
            OfficeColor? highlight = ResolveRunHighlightColor(ResolveRunHighlight(paragraph));
            if (highlight.HasValue) {
                return highlight;
            }

            return TryResolveShadingFillColor(ResolveRunShading(paragraph), colorScheme, out OfficeColor shading)
                ? shading
                : null;
        }

        internal static string? ResolveRunShadingFillColorHex(WordParagraph paragraph) {
            A.ColorScheme? colorScheme = GetDocumentColorScheme(paragraph._document);
            return TryResolveShadingFillColor(ResolveRunShading(paragraph), colorScheme, out OfficeColor shading)
                ? shading.ToRgbHex()
                : null;
        }

        internal static HighlightColorValues? ResolveRunHighlight(WordParagraph paragraph) {
            RunProperties? directProperties = GetDirectRunProperties(paragraph);
            if (directProperties?.Highlight != null) {
                return directProperties.Highlight.Val?.Value;
            }

            foreach (StyleRunProperties properties in EnumerateRunStyleProperties(paragraph)) {
                Highlight? highlight = properties.GetFirstChild<Highlight>();
                if (highlight != null) {
                    return highlight.Val?.Value;
                }
            }

            return null;
        }

        private static Shading? ResolveRunShading(WordParagraph paragraph) {
            RunProperties? directProperties = GetDirectRunProperties(paragraph);
            if (directProperties?.Shading != null) {
                return directProperties.Shading;
            }

            foreach (StyleRunProperties properties in EnumerateRunStyleProperties(paragraph)) {
                Shading? shading = properties.GetFirstChild<Shading>();
                if (shading != null) {
                    return shading;
                }
            }

            return null;
        }

        private static RunProperties? GetDirectRunProperties(WordParagraph paragraph) =>
            paragraph.IsHyperLink
                ? paragraph.Hyperlink?._runProperties
                : paragraph._runProperties;

        private static OfficeColor? ResolveRunHighlightColor(HighlightColorValues? highlight) {
            if (!highlight.HasValue || highlight.Value == HighlightColorValues.None) {
                return null;
            }

            if (highlight.Value == HighlightColorValues.Black) return OfficeColor.Black;
            if (highlight.Value == HighlightColorValues.Blue) return OfficeColor.Blue;
            if (highlight.Value == HighlightColorValues.Cyan) return OfficeColor.Cyan;
            if (highlight.Value == HighlightColorValues.Green) return OfficeColor.Lime;
            if (highlight.Value == HighlightColorValues.Magenta) return OfficeColor.Magenta;
            if (highlight.Value == HighlightColorValues.Red) return OfficeColor.Red;
            if (highlight.Value == HighlightColorValues.Yellow) return OfficeColor.Yellow;
            if (highlight.Value == HighlightColorValues.White) return OfficeColor.White;
            if (highlight.Value == HighlightColorValues.DarkBlue) return OfficeColor.FromRgb(0, 0, 139);
            if (highlight.Value == HighlightColorValues.DarkCyan) return OfficeColor.FromRgb(0, 139, 139);
            if (highlight.Value == HighlightColorValues.DarkGreen) return OfficeColor.FromRgb(0, 100, 0);
            if (highlight.Value == HighlightColorValues.DarkMagenta) return OfficeColor.FromRgb(139, 0, 139);
            if (highlight.Value == HighlightColorValues.DarkRed) return OfficeColor.FromRgb(139, 0, 0);
            if (highlight.Value == HighlightColorValues.DarkYellow) return OfficeColor.FromRgb(184, 134, 11);
            if (highlight.Value == HighlightColorValues.LightGray) return OfficeColor.LightGray;
            if (highlight.Value == HighlightColorValues.DarkGray) return OfficeColor.DarkGray;
            return null;
        }
    }
}
