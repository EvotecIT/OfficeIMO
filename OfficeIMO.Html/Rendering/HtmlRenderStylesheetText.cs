namespace OfficeIMO.Html;

internal static class HtmlRenderStylesheetText {
    internal static bool TryDecode(byte[] bytes, string? contentType, out string css) =>
        HtmlTextEncodingResolver.Default.TryDecodeCss(bytes, contentType, out css);
}
