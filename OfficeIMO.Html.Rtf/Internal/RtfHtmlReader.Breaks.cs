namespace OfficeIMO.Html;

internal static partial class RtfHtmlReader {
    private sealed partial class ReadContext {
        private void AddBreak(IElement token) {
            EnsureInlineParagraph().AddBreak(ReadBreakKind(token));
        }

        private static RtfBreakKind ReadBreakKind(IElement token) {
            string? value = GetAttribute(token, "data-officeimo-rtf-break");
            switch (value?.Trim().ToLowerInvariant()) {
                case "soft-line":
                case "softline":
                    return RtfBreakKind.SoftLine;
                case "page":
                    return RtfBreakKind.Page;
                case "soft-page":
                case "softpage":
                    return RtfBreakKind.SoftPage;
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
