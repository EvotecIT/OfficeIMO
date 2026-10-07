namespace OfficeIMO.Workflows;

/// <summary>Explicit book-title classifications from ONIX list 15. Serial-only types are excluded.</summary>
public enum BookOnixAlternativeTitleType {
    /// <summary>Original-language title of a translated book (03).</summary>
    OriginalLanguage,
    /// <summary>Abbreviated form of the distinctive title (05).</summary>
    Abbreviated,
    /// <summary>Parallel title in another language (06).</summary>
    OtherLanguage,
    /// <summary>Title under which the book was previously published (08).</summary>
    Former,
    /// <summary>Title used in a distributor's catalog (10).</summary>
    Distributor,
    /// <summary>Alternative title appearing on the cover (11).</summary>
    Cover,
    /// <summary>Alternative title appearing on the back cover (12).</summary>
    BackCover,
    /// <summary>Expanded title containing additional identifying context (13).</summary>
    Expanded,
    /// <summary>Another title by which the book is known, including a former working title (14).</summary>
    Alternative,
    /// <summary>Alternative title appearing on the spine (15).</summary>
    Spine,
    /// <summary>Title in an intermediate language through which the book was translated (16).</summary>
    TranslatedFrom
}

/// <summary>
/// Publisher-supplied alternative product-level title. Does not replace the selected EPUB title or
/// infer translation history, language, market applicability or metadata from the EPUB.
/// </summary>
public sealed record BookOnixAlternativeTitle {
    /// <summary>Explicit title classification.</summary>
    public required BookOnixAlternativeTitleType Type { get; init; }
    /// <summary>Full title text, preserving punctuation and Unicode.</summary>
    public required string Title { get; init; }
    /// <summary>Optional subtitle associated with this alternative.</summary>
    public string? Subtitle { get; init; }
    /// <summary>Optional ONIX list 74 code for this title and subtitle; omission leaves language unspecified.</summary>
    public string? LanguageCode { get; init; }
    /// <summary>Optional sorting prefix or explicit no-prefix assertion for this title.</summary>
    public BookOnixTitleSorting? TitleSorting { get; init; }
}
