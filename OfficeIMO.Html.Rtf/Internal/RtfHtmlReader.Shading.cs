namespace OfficeIMO.Html.Rtf;

internal static partial class RtfHtmlReader {
    private sealed partial class ReadContext {
        private void ApplyParagraphShading(HtmlStyleDeclaration style) {
            if (_paragraph == null) {
                return;
            }

            if (style.ShadingForegroundColor != null) {
                _paragraph.ShadingForegroundColorIndex = GetOrAddColorIndex(style.ShadingForegroundColor);
            }

            if (style.ShadingPatternPercent.HasValue) {
                _paragraph.ShadingPatternPercent = style.ShadingPatternPercent.Value;
            }

            if (style.ShadingPattern.HasValue) {
                _paragraph.ShadingPattern = style.ShadingPattern.Value;
            }
        }

        private void ApplyCharacterShading(RtfRun run) {
            RtfColor? foreground = ResolveStyleColor(style => style.ShadingForegroundColor);
            if (foreground != null) {
                run.CharacterShadingForegroundColorIndex = GetOrAddColorIndex(foreground);
            }

            int? percent = ResolveStyleInteger(style => style.ShadingPatternPercent);
            if (percent.HasValue) {
                run.CharacterShadingPatternPercent = percent.Value;
            }

            RtfShadingPattern? pattern = ResolveShadingPattern();
            if (pattern.HasValue) {
                run.CharacterShadingPattern = pattern.Value;
            }
        }

        private int? ResolveStyleInteger(Func<HtmlStyleDeclaration, int?> selector) {
            foreach (HtmlStyleScope scope in _styles) {
                int? value = selector(scope.Style);
                if (value.HasValue) {
                    return value.Value;
                }
            }

            return null;
        }

        private RtfShadingPattern? ResolveShadingPattern() {
            foreach (HtmlStyleScope scope in _styles) {
                if (scope.Style.ShadingPattern.HasValue) {
                    return scope.Style.ShadingPattern.Value;
                }
            }

            return null;
        }
    }
}
