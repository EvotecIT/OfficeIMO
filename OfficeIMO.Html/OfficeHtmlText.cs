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
}
