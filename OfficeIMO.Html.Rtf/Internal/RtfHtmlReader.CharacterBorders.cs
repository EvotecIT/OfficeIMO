namespace OfficeIMO.Html.Rtf;

internal static partial class RtfHtmlReader {
    private sealed partial class ReadContext {
        private HtmlBorderDeclaration? ResolveCharacterBorder() {
            foreach (HtmlStyleScope scope in _styles) {
                if (!scope.Style.HasBorderFormatting) {
                    continue;
                }

                return TryGetEquivalentCharacterBorder(scope.Style, out HtmlBorderDeclaration? border) ? border : null;
            }

            return null;
        }

        private void ApplyCharacterBorder(RtfCharacterBorder target, HtmlBorderDeclaration source) {
            target.Style = MapParagraphBorderStyle(source.Style);
            target.Width = source.Width;
            target.ColorIndex = source.Color == null ? null : GetOrAddColorIndex(source.Color);
        }

        private static bool TryGetEquivalentCharacterBorder(HtmlStyleDeclaration style, out HtmlBorderDeclaration? border) {
            border = null;
            if (style.TopBorder == null ||
                style.LeftBorder == null ||
                style.BottomBorder == null ||
                style.RightBorder == null) {
                return false;
            }

            if (!BorderEquals(style.TopBorder, style.LeftBorder) ||
                !BorderEquals(style.TopBorder, style.BottomBorder) ||
                !BorderEquals(style.TopBorder, style.RightBorder)) {
                return false;
            }

            border = style.TopBorder;
            return true;
        }

        private static bool BorderEquals(HtmlBorderDeclaration left, HtmlBorderDeclaration right) =>
            left.Style == right.Style &&
            left.Width == right.Width &&
            ColorEquals(left.Color, right.Color);

        private static bool ColorEquals(RtfColor? left, RtfColor? right) {
            if (left == null || right == null) {
                return left == null && right == null;
            }

            return left.Red == right.Red &&
                   left.Green == right.Green &&
                   left.Blue == right.Blue;
        }
    }
}
