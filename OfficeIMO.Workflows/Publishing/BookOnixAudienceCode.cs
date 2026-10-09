namespace OfficeIMO.Workflows;

/// <summary>Supported additional readership and education schemes from ONIX list 29.</summary>
public enum BookOnixAudienceScheme {
    /// <summary>A named publisher or recipient scheme (02).</summary>
    Proprietary,
    /// <summary>French Canadian BTLF readership codes (06).</summary>
    Btlf,
    /// <summary>French Electre readership codes (07).</summary>
    Electre,
    /// <summary>Spanish ANELE educational audience and material type (08).</summary>
    Anele,
    /// <summary>AVI reading levels used in Flanders (09).</summary>
    Avi,
    /// <summary>Flemish AWS readership codes (11).</summary>
    Aws,
    /// <summary>Finnish school or college level (15).</summary>
    FinnishSchoolLevel,
    /// <summary>UK Children's Book Group age guidance (16).</summary>
    CbgAgeGuidance,
    /// <summary>NielsenIQ BookData readership codes (17).</summary>
    BookData,
    /// <summary>Revised AVI reading levels used in the Netherlands (18).</summary>
    AviRevised,
    /// <summary>Japanese children's readership, expressed as two ASCII digits (21).</summary>
    JapaneseChildren,
    /// <summary>CEFR language learning levels A1 through C2 (23).</summary>
    Cefr,
    /// <summary>Intended audience language, expressed as a list 74 language code (27).</summary>
    IntendedLanguage,
    /// <summary>Swedish higher secondary curriculum (29).</summary>
    SwedishCurriculum,
    /// <summary>ISCED 2011 education classification (30).</summary>
    Isced2011
}

/// <summary>
/// An explicit code, heading, or both from an additional audience scheme. External code membership and reader suitability
/// are publisher assertions, not independently verified classifications.
/// </summary>
public sealed record BookOnixAudienceCode(BookOnixAudienceScheme Scheme, string? Value = null) {
    /// <summary>Up to 16 plain-text headings. At least one is required when Value is null; repeated headings require distinct explicit languages.</summary>
    public IReadOnlyList<BookOnixAudienceHeading> Headings { get; init; } = [];
    /// <summary>Required for proprietary schemes and unavailable for other schemes. Identifies the agreed code vocabulary.</summary>
    public string? SchemeName { get; init; }
    /// <summary>Marks the main audience. At most one declaration per list 29 code type may be main, including proprietary schemes.</summary>
    public bool IsMain { get; init; }
}
