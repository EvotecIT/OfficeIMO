namespace OfficeIMO.Workflows;

/// <summary>Supported ONIX list 153 collateral-text roles.</summary>
public enum BookOnixTextType {
    /// <summary>Product description, at most 350 Unicode scalar values (02).</summary>
    ShortDescription = 2,
    /// <summary>Product description (03).</summary>
    Description = 3,
    /// <summary>Table of contents (04).</summary>
    TableOfContents = 4,
    /// <summary>Primary cover copy (05).</summary>
    CoverCopy = 5,
    /// <summary>Review of this product or work (06).</summary>
    ReviewQuote = 6,
    /// <summary>Review of a previous edition (07).</summary>
    PreviousEditionReview = 7,
    /// <summary>Review of a previous work (08).</summary>
    PreviousWorkReview = 8,
    /// <summary>Endorsement, not a review (09).</summary>
    Endorsement = 9,
    /// <summary>Promotional headline (10).</summary>
    PromotionalHeadline = 10,
    /// <summary>Product feature (11).</summary>
    Feature = 11,
    /// <summary>Note about all contributors, not a single contributor (12).</summary>
    BiographicalNote = 12,
    /// <summary>Publisher's notice (13).</summary>
    PublisherNotice = 13,
    /// <summary>Excerpt from the work (14).</summary>
    Excerpt = 14,
    /// <summary>Index as a single text (15).</summary>
    Index = 15,
    /// <summary>Collection description, at most 350 Unicode scalar values (16).</summary>
    CollectionShortDescription = 16,
    /// <summary>Collection description (17).</summary>
    CollectionDescription = 17,
    /// <summary>New feature of this edition (18).</summary>
    NewFeature = 18,
    /// <summary>Version history (19).</summary>
    VersionHistory = 19
}

/// <summary>Recipients of collateral, distinct from the book's readership (ONIX list 154).</summary>
public enum BookOnixContentAudience {
    /// <summary>Any audience; cannot accompany another recipient code (00).</summary>
    Unrestricted = 0,
    /// <summary>Distribution by agreement between parties (01); metadata does not enforce access control.</summary>
    Restricted = 1,
    /// <summary>Book trade (02).</summary>
    BookTrade = 2,
    /// <summary>Potential purchasers (03).</summary>
    EndCustomers = 3,
    /// <summary>Librarians (04).</summary>
    Librarians = 4,
    /// <summary>Teachers and educators (05).</summary>
    Teachers = 5,
    /// <summary>Students (06).</summary>
    Students = 6,
    /// <summary>Press and traditional media (07).</summary>
    Press = 7,
    /// <summary>Shopping comparison services (08).</summary>
    ShoppingComparison = 8,
    /// <summary>Search indexing, not display (09).</summary>
    SearchIndex = 9,
    /// <summary>Social media (10).</summary>
    SocialMedia = 10,
    /// <summary>Children (11).</summary>
    Children = 11,
    /// <summary>Teens (12).</summary>
    Teens = 12
}

/// <summary>Representation of a collateral text variant.</summary>
public enum BookOnixCollateralTextFormat {
    /// <summary>Literal plain text (ONIX 06).</summary>
    PlainText,
    /// <summary>Well-formed XHTML fragment in the supported semantic authoring profile (ONIX 05).</summary>
    Xhtml
}

/// <summary>Text or an explicit XHTML fragment, with optional ONIX list 74 language.</summary>
public sealed record BookOnixCollateralTextValue(string Text, string? LanguageCode = null) {
    /// <summary>Plain text by default. XHTML is supported for Texts, not SourceTitles.</summary>
    public BookOnixCollateralTextFormat Format { get; init; }
}

/// <summary>Publisher-supplied supporting text, attribution and usage assertions. No content is fetched or inferred.</summary>
public sealed record BookOnixCollateralText {
    /// <summary>Purpose of this collateral item.</summary>
    public required BookOnixTextType Type { get; init; }
    /// <summary>One to 13 distinct recipient codes; must be explicitly supplied.</summary>
    public required IReadOnlyList<BookOnixContentAudience> Audiences { get; init; }
    /// <summary>One to 16 language variants; repeated variants require distinct explicit languages; at most 65,536 UTF-16 code units each, with a 524,288-unit aggregate per export including source titles, rating units, license names, license expression links, supporting-resource notes, filenames and links.</summary>
    public required IReadOnlyList<BookOnixCollateralTextValue> Texts { get; init; }
    /// <summary>Optional publisher-supplied score for review text types only; no scale or units are inferred.</summary>
    public BookOnixReviewRating? ReviewRating { get; init; }
    /// <summary>Optional territory where the collateral may be used, independent of product sales rights.</summary>
    public BookOnixTerritory? Territory { get; init; }
    /// <summary>At most 16 author names for this text, distinct from book contributors.</summary>
    public IReadOnlyList<string> Authors { get; init; } = [];
    /// <summary>Optional corporate source of this text.</summary>
    public string? SourceCorporate { get; init; }
    /// <summary>At most 16 language variants of the publication/source title; repeated variants require distinct explicit languages; at most 4096 UTF-16 code units each.</summary>
    public IReadOnlyList<BookOnixCollateralTextValue> SourceTitles { get; init; } = [];
    /// <summary>At most 16 distinct absolute HTTP(S) source links without credentials. Export does not fetch them.</summary>
    public IReadOnlyList<string> SourceLinks { get; init; } = [];
    /// <summary>Up to 32 explicit usage constraints. No permissions are inferred or enforced.</summary>
    public IReadOnlyList<BookOnixUsageConstraint> UsageConstraints { get; init; } = [];
    /// <summary>Up to 16 explicit collateral licenses. Names and expression links share the aggregate text budget.</summary>
    public IReadOnlyList<BookOnixLicense> Licenses { get; init; } = [];
    /// <summary>Publication date of the collateral, independent of the book's publication date.</summary>
    public DateOnly? PublishedOn { get; init; }
    /// <summary>First permitted use date (embargo). Serialized, not enforced as an access-control rule.</summary>
    public DateOnly? UsableFrom { get; init; }
    /// <summary>Last permitted use date.</summary>
    public DateOnly? UsableUntil { get; init; }
    /// <summary>Date the collateral was last updated.</summary>
    public DateOnly? UpdatedOn { get; init; }
}
