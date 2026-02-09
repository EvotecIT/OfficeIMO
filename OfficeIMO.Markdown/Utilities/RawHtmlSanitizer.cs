namespace OfficeIMO.Markdown;

/// <summary>
/// Dependency-free allowlist sanitizer for raw HTML blocks. This is intentionally small and conservative.
/// </summary>
internal static class RawHtmlSanitizer {
    internal static string Sanitize(string html) {
        if (string.IsNullOrEmpty(html)) return string.Empty;

        var sb = new StringBuilder(html.Length + 32);
        int i = 0;
        while (i < html.Length) {
            int lt = html.IndexOf('<', i);
            if (lt < 0) {
                sb.Append(System.Net.WebUtility.HtmlEncode(html.Substring(i)));
                break;
            }
            if (lt > i) sb.Append(System.Net.WebUtility.HtmlEncode(html.Substring(i, lt - i)));

            int gt = html.IndexOf('>', lt + 1);
            if (gt < 0) {
                sb.Append(System.Net.WebUtility.HtmlEncode(html.Substring(lt)));
                break;
            }

            string tag = html.Substring(lt, gt - lt + 1);
            if (TrySanitizeAllowedTag(tag, out var sanitized)) {
                sb.Append(sanitized);
            } else {
                sb.Append(System.Net.WebUtility.HtmlEncode(tag));
            }
            i = gt + 1;
        }

        return sb.ToString();
    }

    private static bool TrySanitizeAllowedTag(string tag, out string sanitized) {
        sanitized = string.Empty;
        if (string.IsNullOrEmpty(tag)) return false;
        if (tag.Length < 3) return false;
        if (tag[0] != '<' || tag[tag.Length - 1] != '>') return false;

        // Comments are not allowed; escape them.
        if (tag.StartsWith("<!--", StringComparison.Ordinal)) return false;

        // Extract name and whether it's a closing tag.
        int p = 1;
        while (p < tag.Length && char.IsWhiteSpace(tag[p])) p++;
        bool closing = p < tag.Length && tag[p] == '/';
        if (closing) p++;
        while (p < tag.Length && char.IsWhiteSpace(tag[p])) p++;

        int nameStart = p;
        while (p < tag.Length && (char.IsLetterOrDigit(tag[p]) || tag[p] == '-')) p++;
        if (p <= nameStart) return false;

        string name = tag.Substring(nameStart, p - nameStart).ToLowerInvariant();

        switch (name) {
            case "br": {
                if (closing) return false;
                sanitized = "<br/>";
                return true;
            }
            case "u": {
                sanitized = closing ? "</u>" : "<u>";
                return true;
            }
            case "summary": {
                sanitized = closing ? "</summary>" : "<summary>";
                return true;
            }
            case "details": {
                if (closing) {
                    sanitized = "</details>";
                    return true;
                }

                // Allow only "open" (boolean attribute).
                bool open = tag.IndexOf(" open", StringComparison.OrdinalIgnoreCase) >= 0
                            || tag.IndexOf(" open>", StringComparison.OrdinalIgnoreCase) >= 0
                            || tag.IndexOf(" open/", StringComparison.OrdinalIgnoreCase) >= 0;
                sanitized = open ? "<details open>" : "<details>";
                return true;
            }
            default:
                return false;
        }
    }
}

