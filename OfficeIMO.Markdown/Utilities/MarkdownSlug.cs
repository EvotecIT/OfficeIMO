namespace OfficeIMO.Markdown;

/// <summary>
/// Slug utilities for generating anchor ids compatible with GitHub-like platforms.
/// </summary>
internal static class MarkdownSlug {
    public static string GitHub(string text) {
        if (string.IsNullOrEmpty(text)) return string.Empty;
        var sb = new StringBuilder(text.Length);
        bool prevHyphen = false;
        foreach (char ch in text.ToLowerInvariant()) {
            if ((ch >= 'a' && ch <= 'z') || (ch >= '0' && ch <= '9')) { sb.Append(ch); prevHyphen = false; } else if (ch == ' ' || ch == '-' || ch == '_') { if (!prevHyphen) { sb.Append('-'); prevHyphen = true; } } else {
                // skip punctuation
            }
        }
        // trim trailing hyphen
        var result = sb.ToString().Trim('-');
        return result;
    }
}

