namespace OfficeIMO.Workflows;

/// <summary>Non-deprecated ONIX list 32 reading and listening complexity schemes.</summary>
public enum BookOnixComplexityScheme {
    /// <summary>Fry readability score, integer 1 through 15 (03).</summary>
    FryReadability,
    /// <summary>Institute of Education Book Band, supplied by the publisher (04).</summary>
    IoeBookBand,
    /// <summary>Fountas and Pinnell text level, A through Z or Z+ (05).</summary>
    FountasAndPinnell,
    /// <summary>Combined English-text Lexile measure (06).</summary>
    Lexile,
    /// <summary>ATOS for Books, decimal score 0 through 17 (07).</summary>
    Atos,
    /// <summary>Flesch-Kincaid grade level; negative scores are permitted (08).</summary>
    FleschKincaid,
    /// <summary>Publisher or third-party Guided Reading level (09).</summary>
    GuidedReading,
    /// <summary>Reading Recovery level, integer 1 through 20 (10).</summary>
    ReadingRecovery,
    /// <summary>Scandinavian LIX readability index (11).</summary>
    Lix,
    /// <summary>Lexile listening-comprehension measure (12).</summary>
    LexileAudio,
    /// <summary>Spanish-text Lexile measure (13).</summary>
    LexileSpanish
}

/// <summary>A publisher-supplied complexity value, at most 20 characters, without surrounding whitespace.
/// Numeric schemes use invariant decimal notation. External code assignment and suitability are not verified.</summary>
public sealed record BookOnixComplexity(BookOnixComplexityScheme Scheme, string Value);
