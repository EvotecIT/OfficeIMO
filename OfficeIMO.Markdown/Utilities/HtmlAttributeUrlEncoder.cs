namespace OfficeIMO.Markdown;

internal static class HtmlAttributeUrlEncoder {
    internal static string Encode(string? url) {
        if (string.IsNullOrEmpty(url)) {
            return string.Empty;
        }

        var value = url!;
        return System.Net.WebUtility.HtmlEncode(value.Replace(" ", "%20"));
    }
}
