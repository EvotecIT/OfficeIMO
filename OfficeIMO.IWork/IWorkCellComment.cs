namespace OfficeIMO.IWork;

/// <summary>A qualified root table-cell comment, without replies or collaboration state.</summary>
public sealed class IWorkCellComment {
    internal IWorkCellComment(string text, string author, DateTime creationDateUtc,
        IWorkObjectIdentity sourceIdentity, IWorkObjectIdentity sourceAuthorIdentity) {
        Text = text;
        Author = author;
        CreationDateUtc = creationDateUtc;
        SourceIdentity = sourceIdentity;
        SourceAuthorIdentity = sourceAuthorIdentity;
    }

    /// <summary>Gets the exact plain comment text.</summary>
    public string Text { get; }
    /// <summary>Gets the source author's display name.</summary>
    public string Author { get; }
    /// <summary>Gets the native creation timestamp converted from the Apple epoch to UTC.</summary>
    public DateTime CreationDateUtc { get; }
    /// <summary>Gets the selected native comment record identity.</summary>
    public IWorkObjectIdentity SourceIdentity { get; }
    /// <summary>Gets the selected native author record identity. Display colors and collaboration identifiers are not projected.</summary>
    public IWorkObjectIdentity SourceAuthorIdentity { get; }
}
