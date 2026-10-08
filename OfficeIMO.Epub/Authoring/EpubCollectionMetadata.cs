namespace OfficeIMO.Epub;

/// <summary>The standard EPUB 3 forms of publication collection.</summary>
public enum EpubCollectionKind {
    /// <summary>An open-ended sequence of individually issued related works.</summary>
    Series,
    /// <summary>A finite group of works forming one intellectual unit.</summary>
    Set
}

/// <summary>Membership in a named EPUB 3 series or set, independent of spine order.</summary>
public sealed class EpubCollectionMetadata {
    /// <summary>Collection display name.</summary>
    public string Name { get; set; } = string.Empty;
    /// <summary>Whether the collection is a series or set.</summary>
    public EpubCollectionKind Kind { get; set; } = EpubCollectionKind.Series;
    /// <summary>Optional sorting name for the collection.</summary>
    public string? FileAs { get; set; }
    /// <summary>Optional BCP 47 language of the collection name.</summary>
    public string? Language { get; set; }
    /// <summary>
    /// Optional hierarchical position, serialized with dot separators: [2, 1] becomes 2.1.
    /// Empty leaves the position unspecified. This is an ordered identifier, not a decimal fraction.
    /// </summary>
    public IReadOnlyList<uint> Position { get; set; } = Array.Empty<uint>();
}
