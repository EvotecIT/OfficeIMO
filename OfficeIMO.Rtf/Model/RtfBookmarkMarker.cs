namespace OfficeIMO.Rtf;

/// <summary>
/// Start or end marker for a named RTF bookmark.
/// </summary>
public sealed class RtfBookmarkMarker : IRtfInline {
    /// <summary>Creates a bookmark marker.</summary>
    public RtfBookmarkMarker(RtfBookmarkMarkerKind kind, string name) {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Bookmark name cannot be empty.", nameof(name));
        Kind = kind;
        Name = name;
    }

    /// <summary>Marker kind.</summary>
    public RtfBookmarkMarkerKind Kind { get; }

    /// <summary>Bookmark name.</summary>
    public string Name { get; }
}
