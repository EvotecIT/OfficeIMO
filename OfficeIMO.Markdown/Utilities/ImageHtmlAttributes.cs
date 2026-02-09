namespace OfficeIMO.Markdown;

internal static class ImageHtmlAttributes {
    internal static string BuildImageAttributes(HtmlOptions? o, string? src) {
        if (o == null) return string.Empty;

        var sb = new System.Text.StringBuilder();

        if (o.ImagesLoadingLazy) sb.Append(" loading=\"lazy\"");
        if (o.ImagesDecodingAsync) sb.Append(" decoding=\"async\"");

        var rp = (o.ImagesReferrerPolicy ?? string.Empty).Trim();
        if (rp.Length > 0) sb.Append(" referrerpolicy=\"").Append(System.Net.WebUtility.HtmlEncode(rp)).Append("\"");

        return sb.ToString();
    }

    internal static string BuildBlockedPlaceholder(string? alt) {
        // Keep markup minimal and safe.
        alt ??= string.Empty;
        var trimmed = alt.Trim();
        var text = trimmed.Length == 0 ? "image blocked" : ("image blocked: " + trimmed);
        return "<span class=\"omd-image-blocked\">" + System.Net.WebUtility.HtmlEncode(text) + "</span>";
    }
}
