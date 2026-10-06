namespace OfficeIMO.Workflows;

/// <summary>Collection provenance from ONIX list 148.</summary>
public enum BookOnixCollectionType {
    /// <summary>Publisher-defined bibliographic series or set (10).</summary>
    Publisher,
    /// <summary>Publisher marketing collection, or collection éditoriale (11).</summary>
    Editorial,
    /// <summary>Collection defined by another party in the supply chain (20).</summary>
    Ascribed
}

/// <summary>Identifier schemes supported for a collection.</summary>
public enum BookOnixCollectionIdentifierType {
    /// <summary>Named proprietary collection identifier (01).</summary>
    Proprietary,
    /// <summary>ISSN (02), validated and serialized without its hyphen.</summary>
    Issn,
    /// <summary>ISBN-13 (15). Use only when the collection is available as a single product.</summary>
    Isbn13
}

/// <summary>An explicit collection identifier; SchemeName is required only for proprietary identifiers.</summary>
public sealed record BookOnixCollectionIdentifier(BookOnixCollectionIdentifierType Type, string Value, string? SchemeName = null);

/// <summary>Collection ordering semantics from ONIX list 197.</summary>
public enum BookOnixCollectionSequenceType {
    /// <summary>Publisher-defined ordering with an explicit name (01).</summary>
    Proprietary,
    /// <summary>Order specified by the title (02).</summary>
    Title,
    /// <summary>Publication order (03).</summary>
    Publication,
    /// <summary>Temporal or narrative order (04).</summary>
    Narrative,
    /// <summary>Original publication order of a republished collection (05).</summary>
    OriginalPublication,
    /// <summary>Suggested reading order (06).</summary>
    SuggestedReading,
    /// <summary>Suggested display order (07).</summary>
    SuggestedDisplay
}

/// <summary>A product position expressed as dot-separated nonnegative ASCII integers or hyphens, such as 2.1 or 3.-.8.</summary>
/// <param name="Type">Ordering semantics.</param>
/// <param name="Number">Position, preserved as text; not a decimal fraction.</param>
/// <param name="Name">Required only for proprietary sequences.</param>
public sealed record BookOnixCollectionSequence(BookOnixCollectionSequenceType Type, string Number, string? Name = null);

/// <summary>Explicit membership in one named collection. Does not infer ONIX semantics from EPUB series metadata.</summary>
public sealed record BookOnixCollection {
    /// <summary>Who defines the collection.</summary>
    public required BookOnixCollectionType Type { get; init; }
    /// <summary>Top-level collection title, separate from the product title.</summary>
    public required string Title { get; init; }
    /// <summary>Optional subtitle of the collection.</summary>
    public string? Subtitle { get; init; }
    /// <summary>Optional ONIX list 74 language for the collection title and subtitle.</summary>
    public string? LanguageCode { get; init; }
    /// <summary>Source of an ascribed collection; required for Ascribed and optional otherwise.</summary>
    public string? SourceName { get; init; }
    /// <summary>Up to 16 identifiers; at most one value per type and proprietary scheme name.</summary>
    public IReadOnlyList<BookOnixCollectionIdentifier> Identifiers { get; init; } = [];
    /// <summary>Up to 16 positions; at most one per type and proprietary sequence name.</summary>
    public IReadOnlyList<BookOnixCollectionSequence> Sequences { get; init; } = [];
}
