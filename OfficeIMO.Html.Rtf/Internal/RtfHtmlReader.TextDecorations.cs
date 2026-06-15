namespace OfficeIMO.Html.Rtf;

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

        private void ApplyRichStrike(RtfRun run, bool strike) {
            if (!strike) {
                return;
            }

            bool? doubleStrike = ResolveDoubleStrike();
            if (doubleStrike.GetValueOrDefault()) {
                run.Strike = false;
                run.DoubleStrike = true;
            }
        }

        private bool? ResolveDoubleStrike() {
            foreach (HtmlStyleScope scope in _styles) {
                if (scope.Style.DoubleStrike.HasValue) {
                    return scope.Style.DoubleStrike.Value;
                }
            }

            return null;
        }
    }
}
