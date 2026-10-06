namespace OfficeIMO.Workflows;

/// <summary>Supported ONIX list 27 subject schemes.</summary>
public enum BookOnixSubjectScheme {
    /// <summary>Dewey Decimal Classification (01).</summary>
    Dewey,
    /// <summary>Library of Congress classification (03).</summary>
    LibraryOfCongressClassification,
    /// <summary>Library of Congress subject heading (04).</summary>
    LibraryOfCongressHeading,
    /// <summary>BISAC subject heading (10).</summary>
    Bisac,
    /// <summary>Indexing/search keywords in heading text (20).</summary>
    Keywords,
    /// <summary>Named proprietary scheme (24).</summary>
    Proprietary,
    /// <summary>Thema subject category (93).</summary>
    Thema,
    /// <summary>Thema geographical qualifier (94).</summary>
    ThemaGeographical,
    /// <summary>Thema language qualifier (95).</summary>
    ThemaLanguage,
    /// <summary>Thema time-period qualifier (96).</summary>
    ThemaTimePeriod,
    /// <summary>Thema educational-purpose qualifier (97).</summary>
    ThemaEducationalPurpose,
    /// <summary>Thema interest-age or special-interest qualifier (98).</summary>
    ThemaInterest,
    /// <summary>Thema style qualifier (99).</summary>
    ThemaStyle
}

/// <summary>A subject heading with an optional three-letter ONIX list 74 language code.</summary>
public sealed record BookOnixSubjectHeading(string Text, string? LanguageCode = null);

/// <summary>Explicit subject classification. Code membership and suitability are not inferred or verified.</summary>
public sealed record BookOnixSubject {
    /// <summary>Subject scheme.</summary>
    public required BookOnixSubjectScheme Scheme { get; init; }
    /// <summary>Required only for a proprietary scheme; forbidden for standard schemes.</summary>
    public string? SchemeName { get; init; }
    /// <summary>Optional scheme version, retained verbatim.</summary>
    public string? SchemeVersion { get; init; }
    /// <summary>Optional subject code. At least a code or heading is required; keywords use headings only.</summary>
    public string? Code { get; init; }
    /// <summary>At most 16 headings, with at most one for each language and one with unspecified language.</summary>
    public IReadOnlyList<BookOnixSubjectHeading> Headings { get; init; } = [];
    /// <summary>Marks the main subject for this scheme. Not allowed for keywords or Thema qualifiers.</summary>
    public bool IsMain { get; init; }
}
