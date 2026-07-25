namespace OfficeIMO.Html;

/// <summary>
/// Source document metadata retained by the shared renderer for output adapters.
/// </summary>
public sealed class HtmlRenderMetadata {
    internal HtmlRenderMetadata(
        string? title,
        string? language,
        HtmlRenderTextDirection direction = HtmlRenderTextDirection.LeftToRight,
        string? author = null,
        string? subject = null,
        string? keywords = null,
        string? creator = null) {
        Title = Normalize(title, 1024);
        Language = Normalize(language, 128);
        Direction = direction;
        Author = Normalize(author, 1024);
        Subject = Normalize(subject, 4096);
        Keywords = Normalize(keywords, 4096);
        Creator = Normalize(creator, 1024);
    }

    /// <summary>HTML document title, when present.</summary>
    public string? Title { get; }

    /// <summary>HTML document language from <c>lang</c> or <c>xml:lang</c>, when present.</summary>
    public string? Language { get; }

    /// <summary>Document-level direction retained for navigation-capable output adapters.</summary>
    public HtmlRenderTextDirection Direction { get; }

    /// <summary>Document author from supported HTML metadata names.</summary>
    public string? Author { get; }

    /// <summary>Document subject or description from supported HTML metadata names.</summary>
    public string? Subject { get; }

    /// <summary>Document keywords from supported HTML metadata names.</summary>
    public string? Keywords { get; }

    /// <summary>Source creator or generator from supported HTML metadata names.</summary>
    public string? Creator { get; }

    internal static HtmlRenderMetadata FromDocument(
        AngleSharp.Html.Dom.IHtmlDocument document,
        HtmlRenderTextDirection direction) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (AngleSharp.Dom.IElement meta in document.QuerySelectorAll("meta[name][content]")) {
            string? name = meta.GetAttribute("name")?.Trim();
            string? content = meta.GetAttribute("content")?.Trim();
            if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(content) || values.ContainsKey(name!)) continue;
            values[name!] = content!;
        }

        string? author = First(values, "author", "dc.creator", "dcterms.creator");
        string? creator = First(values, "creator", "generator", "application-name");
        return new HtmlRenderMetadata(
            document.Title,
            ResolveLanguage(document),
            direction,
            author,
            First(values, "description", "subject", "dc.description", "dcterms.description"),
            First(values, "keywords"),
            creator);
    }

    private static string? ResolveLanguage(AngleSharp.Html.Dom.IHtmlDocument document) {
        AngleSharp.Dom.IElement? root = document.DocumentElement;
        return root?.GetAttribute("lang") ?? root?.GetAttribute("xml:lang");
    }

    private static string? First(IReadOnlyDictionary<string, string> values, params string[] names) {
        foreach (string name in names) {
            if (values.TryGetValue(name, out string? value)) return value;
        }
        return null;
    }

    private static string? Normalize(string? value, int maximumLength) {
        if (string.IsNullOrWhiteSpace(value)) return null;
        string normalized = value!.Trim();
        if (normalized.Length > maximumLength || normalized.Any(char.IsControl)) return null;
        return normalized;
    }
}
