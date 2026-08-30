namespace OfficeIMO.Html;

/// <summary>
/// Shared HTML text escaping helpers for OfficeIMO adapters.
/// </summary>
public static class OfficeHtmlText {
    /// <summary>Escapes text for HTML element content.</summary>
    public static string Escape(string? value) {
        return WebUtility.HtmlEncode(value ?? string.Empty);
    }

    /// <summary>Escapes text for HTML attribute values.</summary>
    public static string EscapeAttribute(string? value) {
        return WebUtility.HtmlEncode(value ?? string.Empty).Replace("\"", "&quot;");
    }

    /// <summary>Quotes a value as a single-quoted CSS string literal.</summary>
    public static string QuoteCssString(string? value) {
        var builder = new System.Text.StringBuilder((value?.Length ?? 0) + 2);
        builder.Append('\'');
        foreach (char character in value ?? string.Empty) {
            switch (character) {
                case '\\':
                    builder.Append("\\\\");
                    break;
                case '\'':
                    builder.Append("\\'");
                    break;
                case '\n':
                    builder.Append("\\A ");
                    break;
                case '\r':
                    builder.Append("\\D ");
                    break;
                case '\f':
                    builder.Append("\\C ");
                    break;
                case '\0':
                    builder.Append("\\FFFD ");
                    break;
                default:
                    if (char.IsControl(character)) {
                        builder.Append('\\').Append(((int)character).ToString("X", System.Globalization.CultureInfo.InvariantCulture)).Append(' ');
                    } else {
                        builder.Append(character);
                    }
                    break;
            }
        }
        return builder.Append('\'').ToString();
    }
}
