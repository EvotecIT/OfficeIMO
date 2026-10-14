namespace OfficeIMO.DjVu;

/// <summary>An immutable entry in the document's stored outline. Targets never trigger file or network access.</summary>
public sealed class DjVuBookmark {
    internal DjVuBookmark(string title, string target, int? pageNumber, List<DjVuBookmark> children) {
        Title = title; Target = target; PageNumber = pageNumber; Children = children.AsReadOnly();
    }
    /// <summary>Exact stored UTF-8 title decoded to a .NET string.</summary>
    public string Title { get; }
    /// <summary>Exact stored target, including local fragments or external URLs.</summary>
    public string Target { get; }
    /// <summary>Resolved one-based local source page when the target identifies a page number or component identity.</summary>
    public int? PageNumber { get; }
    /// <summary>Immediate child entries in source order.</summary>
    public IReadOnlyList<DjVuBookmark> Children { get; }
}
