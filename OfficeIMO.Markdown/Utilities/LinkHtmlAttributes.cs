namespace OfficeIMO.Markdown;

internal static class LinkHtmlAttributes {
    internal static string BuildExternalLinkAttributes(HtmlOptions? o, string? url) {
        if (o == null) return string.Empty;
        var u = (url ?? string.Empty).Trim();
        if (u.Length == 0) return string.Empty;

        // Only apply to absolute HTTP(S) and protocol-relative links. (Not fragments, not relative paths, not mailto/data.)
        bool externalHttp =
            u.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
            u.StartsWith("https://", StringComparison.OrdinalIgnoreCase) ||
            u.StartsWith("//", StringComparison.OrdinalIgnoreCase);
        if (!externalHttp) return string.Empty;

        var sb = new System.Text.StringBuilder();

        if (o.ExternalLinksTargetBlank) {
            sb.Append(" target=\"_blank\"");
        }

        var rel = (o.ExternalLinksRel ?? string.Empty).Trim();
        if (o.ExternalLinksTargetBlank) {
            // If you open a new tab/window, always prevent tabnabbing even if the caller forgot to set rel.
            if (rel.Length == 0) rel = "noopener noreferrer";
            else {
                var relLower = rel.ToLowerInvariant();
                if (!relLower.Contains("noopener")) rel += " noopener";
                if (!relLower.Contains("noreferrer")) rel += " noreferrer";
            }
        }
        if (rel.Length > 0) sb.Append(" rel=\"").Append(System.Net.WebUtility.HtmlEncode(rel)).Append("\"");

        var rp = (o.ExternalLinksReferrerPolicy ?? string.Empty).Trim();
        if (rp.Length > 0) sb.Append(" referrerpolicy=\"").Append(System.Net.WebUtility.HtmlEncode(rp)).Append("\"");

        return sb.ToString();
    }
}
