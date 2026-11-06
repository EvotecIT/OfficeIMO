namespace OfficeIMO.Rtf;

/// <summary>
/// Start or end marker for a named RTF bookmark.
/// </summary>
public sealed partial class RtfBookmarkMarker : IRtfInline {
    /// <summary>Creates a bookmark marker.</summary>
    public RtfBookmarkMarker(RtfBookmarkMarkerKind kind, string name) {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        Kind = kind;
        Name = name;
    }

    /// <summary>Marker kind.</summary>
    public RtfBookmarkMarkerKind Kind { get; }

    /// <summary>Bookmark name.</summary>
    public string Name { get; }
}
