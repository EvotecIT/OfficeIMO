namespace OfficeIMO.Rtf.Html;

internal static partial class RtfHtmlReader {
    private sealed partial class ReadContext {
        private RtfCapsStyle ResolveCapsStyle() {
            foreach (HtmlStyleScope scope in _styles) {
                if (scope.Style.CapsStyle.HasValue) {
                    return scope.Style.CapsStyle.Value;
                }
            }

            return RtfCapsStyle.None;
        }
    }
}
