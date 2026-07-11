namespace OfficeIMO.Html;

/// <summary>
/// Source document metadata retained by the shared renderer for output adapters.
/// </summary>
public sealed class HtmlRenderMetadata {
    internal HtmlRenderMetadata(string? title, string? language) {
        Title = Normalize(title, 1024);
        Language = Normalize(language, 128);
    }

    /// <summary>HTML document title, when present.</summary>
    public string? Title { get; }

    /// <summary>HTML document language from <c>lang</c> or <c>xml:lang</c>, when present.</summary>
    public string? Language { get; }

    private static string? Normalize(string? value, int maximumLength) {
        if (string.IsNullOrWhiteSpace(value)) return null;
        string normalized = value!.Trim();
        if (normalized.Length > maximumLength || normalized.Any(char.IsControl)) return null;
        return normalized;
    }
}
