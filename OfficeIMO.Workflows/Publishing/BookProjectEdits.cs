namespace OfficeIMO.Workflows;

/// <summary>A staged editor batch. Null values retain the existing package values.</summary>
public sealed class BookProjectEdits {
    /// <summary>Replacement primary title.</summary>
    public string? Title { get; init; }
    /// <summary>Replacement primary language tag.</summary>
    public string? Language { get; init; }
    /// <summary>Replacement primary creator. A blank value retains the current value.</summary>
    public string? Creator { get; init; }
    /// <summary>Replacement project typography stylesheet, retaining imported styles.</summary>
    public string? Stylesheet { get; init; }
    /// <summary>Replacement XHTML body elements indexed by existing manifest identifier.</summary>
    public IReadOnlyDictionary<string, string>? ChapterBodies { get; init; }
    /// <summary>Replacement document and primary navigation titles indexed by existing XHTML spine identifier.</summary>
    public IReadOnlyDictionary<string, string>? ChapterTitles { get; init; }
}
