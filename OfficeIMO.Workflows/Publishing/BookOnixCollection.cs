namespace OfficeIMO.Workflows;

/// <summary>Collection provenance from ONIX list 148.</summary>
public enum BookOnixCollectionType {
    /// <summary>Publisher-defined collection (10).</summary>
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
    Isbn13,
    /// <summary>German National Bibliography series identifier (03), supplied by the publisher.</summary>
    GermanNationalBibliography,
    /// <summary>German Books in Print (VLB) series identifier (04), supplied by the publisher.</summary>
    GermanBooksInPrint,
    /// <summary>Electre series identifier (05), supplied by the publisher.</summary>
    Electre,
    /// <summary>Bare Digital Object Identifier (06), beginning with 10. and without a resolver URL.</summary>
    Doi,
    /// <summary>Full Uniform Resource Name (22). Prefer a specific scheme when available.</summary>
    Urn,
    /// <summary>Five-digit Japanese magazine identifier (27), without an issue extension.</summary>
    JapaneseMagazine,
    /// <summary>French National Bibliography series control number (29), supplied by the publisher.</summary>
    BnfControlNumber,
    /// <summary>Archival Resource Key (35), including its HTTP or HTTPS resolver URL.</summary>
    Ark,
    /// <summary>Linking ISSN (38), validated and serialized without its hyphen. Use when distinct from the serial ISSN.</summary>
    IssnL
}

/// <summary>An explicit collection identifier; SchemeName is required only for proprietary identifiers.</summary>
public sealed record BookOnixCollectionIdentifier(BookOnixCollectionIdentifierType Type, string Value, string? SchemeName = null) {
    /// <summary>Optional title element level identified by this value; must occur in the collection title.</summary>
    public BookOnixCollectionLevel? Level { get; init; }
}

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

/// <summary>Explicit membership in one collection or named grouping. Does not infer ONIX semantics from EPUB series metadata.</summary>
public sealed record BookOnixCollection {
    /// <summary>Who defines the collection.</summary>
    public required BookOnixCollectionType Type { get; init; }
    /// <summary>Simple top-level collection title, separate from the product title. Mutually exclusive with TitleElements.</summary>
    public string? Title { get; init; }
    /// <summary>Optional explicit sorting prefix or no-prefix assertion; omission retains unsplit TitleText.</summary>
    public BookOnixTitleSorting? TitleSorting { get; init; }
    /// <summary>Optional subtitle of the collection.</summary>
    public string? Subtitle { get; init; }
    /// <summary>Optional ONIX list 74 language for the collection title and subtitle.</summary>
    public string? LanguageCode { get; init; }
    /// <summary>One to five title elements in display order, with one per level. Mutually exclusive with Title, Subtitle, LanguageCode and TitleSorting.</summary>
    public IReadOnlyList<BookOnixCollectionTitleElement> TitleElements { get; init; } = [];
    /// <summary>Optional explicit frequency of publication of successive products in the collection.</summary>
    public BookOnixCollectionFrequency? Frequency { get; init; }
    /// <summary>Source of an ascribed collection; required for Ascribed and optional otherwise.</summary>
    public string? SourceName { get; init; }
    /// <summary>Up to 16 identifiers; at most one value per level, type and proprietary scheme name. Unscoped and scoped values for the same scheme cannot coexist.</summary>
    public IReadOnlyList<BookOnixCollectionIdentifier> Identifiers { get; init; } = [];
    /// <summary>Up to 16 positions; at most one per type and proprietary sequence name.</summary>
    public IReadOnlyList<BookOnixCollectionSequence> Sequences { get; init; } = [];
    /// <summary>Up to 100 ordered collection credits, independent of product credits. Empty makes no assertion.</summary>
    public IReadOnlyList<BookOnixContributor> Contributors { get; init; } = [];
    /// <summary>Explicit assertion of no collection contributors; mutually exclusive with Contributors.</summary>
    public bool NoContributors { get; init; }
}
