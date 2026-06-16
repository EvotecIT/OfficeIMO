namespace OfficeIMO.Rtf;

/// <summary>
/// Author entry from the RTF revision table.
/// </summary>
public sealed class RtfRevisionAuthor {
    /// <summary>Creates a revision author entry.</summary>
    public RtfRevisionAuthor(string name) {
        Name = name ?? string.Empty;
    }

    /// <summary>Author display name.</summary>
    public string Name { get; set; }
}
