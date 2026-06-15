namespace OfficeIMO.Rtf.Html;

internal static partial class RtfHtmlReader {
    private sealed partial class ReadContext {
        private void ApplyRichUnderline(RtfRun run, bool underline) {
            if (!underline) {
                return;
            }

            RtfUnderlineStyle? underlineStyle = ResolveUnderlineStyle();
            if (underlineStyle.HasValue) {
                run.UnderlineStyle = underlineStyle.Value;
            }

            RtfColor? underlineColor = ResolveStyleColor(style => style.UnderlineColor);
            if (underlineColor != null && run.Underline) {
                run.UnderlineColorIndex = GetOrAddColorIndex(underlineColor);
            }
        }

        private RtfUnderlineStyle? ResolveUnderlineStyle() {
            foreach (HtmlStyleScope scope in _styles) {
                if (scope.Style.UnderlineStyle.HasValue) {
                    return scope.Style.UnderlineStyle.Value;
                }
            }

            return null;
        }
    }
}
