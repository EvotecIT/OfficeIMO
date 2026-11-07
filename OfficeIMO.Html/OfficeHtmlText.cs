namespace OfficeIMO.Html;

/// <summary>
/// Shared HTML text escaping helpers for OfficeIMO adapters.
/// </summary>
public static class OfficeHtmlText {
    /// <summary>Keeps a document family first, followed by its shared installed-font substitutes and a CSS generic family.</summary>
    internal static string CssFontFamily(string family) {
        string source = family.Trim();
        string lower = source.ToLowerInvariant();
        if (lower is "serif" or "sans-serif" or "monospace" or "cursive" or "fantasy" or "system-ui") return lower;
        var families = new System.Collections.Generic.List<string> { QuoteCssString(source) };
        foreach (string substitute in OfficeIMO.Drawing.OfficeDocumentFontSubstitutions.Enumerate(source)) {
            families.Add(QuoteCssString(substitute));
        }
        families.Add(OfficeIMO.Drawing.OfficeDocumentFontSubstitutions.ClassifyGenericFamily(source));
        return string.Join(",", families);
    }

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
