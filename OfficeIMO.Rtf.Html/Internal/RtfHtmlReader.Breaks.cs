namespace OfficeIMO.Rtf.Html;

internal static partial class RtfHtmlReader {
    private sealed partial class ReadContext {
        private void AddBreak(HtmlToken token) {
            EnsureParagraph().AddBreak(ReadBreakKind(token));
        }

        private static RtfBreakKind ReadBreakKind(HtmlToken token) {
            string? value = GetAttribute(token, "data-officeimo-rtf-break");
            switch (value?.Trim().ToLowerInvariant()) {
                case "page":
                    return RtfBreakKind.Page;
                case "column":
                    return RtfBreakKind.Column;
                case "line":
                    return RtfBreakKind.Line;
            }

            HtmlStyleDeclaration style = HtmlStyleDeclarationParser.Parse(GetAttribute(token, "style"));
            return style.PageBreakBefore || style.PageBreakAfter
                ? RtfBreakKind.Page
                : RtfBreakKind.Line;
        }
    }
}
