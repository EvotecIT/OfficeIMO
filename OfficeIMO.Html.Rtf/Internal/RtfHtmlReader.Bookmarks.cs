namespace OfficeIMO.Html;

internal static partial class RtfHtmlReader {
    private sealed partial class ReadContext {
        private void StartAnchor(IElement token) {
            string? href = GetAttribute(token, "href");
            Uri? uri = ReadUriValue(href);
            if (uri != null) {
                _hyperlink = uri;
                return;
            }

            if (!string.IsNullOrWhiteSpace(href) && !IsFragmentHref(href)) {
                _options.AddDiagnostic(
                    "HtmlRtfHyperlinkRejected",
                    "HTML hyperlink was omitted because it was rejected by the configured URL policy.",
                    "href",
                    action: RtfConversionAction.Blocked);
            }

            string? kind = GetAttribute(token, "data-officeimo-rtf-bookmark");
            if (string.Equals(kind, "end", StringComparison.OrdinalIgnoreCase)) {
                string? endName = ReadBookmarkName(token);
                if (!string.IsNullOrWhiteSpace(endName)) {
                    EnsureInlineParagraph().AddBookmarkEnd(endName!);
                }

                return;
            }

            string? name = ReadBookmarkName(token);
            if (!string.IsNullOrWhiteSpace(name)) {
                EnsureInlineParagraph().AddBookmarkStart(name!);
            }
        }

        private static string? ReadBookmarkName(IElement token) =>
            GetAttribute(token, "data-officeimo-rtf-bookmark-name") ??
            GetAttribute(token, "id") ??
            GetAttribute(token, "name");
    }
}
