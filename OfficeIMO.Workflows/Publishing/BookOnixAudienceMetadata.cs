namespace OfficeIMO.Workflows;

/// <summary>ONIX list 28 audience categories.</summary>
public enum BookOnixAudienceType {
    /// <summary>General adult audience (01).</summary>
    GeneralAdult,
    /// <summary>Children (02).</summary>
    Children,
    /// <summary>Teenage audience (03).</summary>
    Teenage,
    /// <summary>Primary and secondary education (04).</summary>
    PrimaryAndSecondaryEducation,
    /// <summary>Tertiary education (05).</summary>
    TertiaryEducation,
    /// <summary>Professional and scholarly (06).</summary>
    ProfessionalAndScholarly,
    /// <summary>English as an additional language (07).</summary>
    EnglishLanguageTeaching,
    /// <summary>Adult education (08).</summary>
    AdultEducation,
    /// <summary>Second or additional language teaching other than English (09).</summary>
    AdditionalLanguageTeaching,
    /// <summary>Pre-primary education (11).</summary>
    PrePrimaryEducation,
    /// <summary>Primary education (12).</summary>
    PrimaryEducation,
    /// <summary>Lower secondary education (13).</summary>
    LowerSecondaryEducation,
    /// <summary>Upper secondary education (14).</summary>
    UpperSecondaryEducation
}

/// <summary>One explicit list 28 audience category. At most one can be main.</summary>
public sealed record BookOnixAudience(BookOnixAudienceType Type, bool IsMain = false) {
    /// <summary>Up to 16 plain-text equivalents of this category; repeated headings require distinct explicit languages.</summary>
    public IReadOnlyList<BookOnixAudienceHeading> Headings { get; init; } = [];
}

/// <summary>Supported age semantics from ONIX list 30.</summary>
public enum BookOnixAgeRangeType {
    /// <summary>Interest age in months (16); first value at most 36, second value at most 42.</summary>
    InterestMonths,
    /// <summary>Interest age in years (17).</summary>
    InterestYears,
    /// <summary>Reading age in years (18).</summary>
    ReadingYears
}

/// <summary>Nonnegative integer age bounds. At least one bound is required; equal bounds express an exact age.</summary>
public sealed record BookOnixAgeRange(BookOnixAgeRangeType Type, int? Minimum = null, int? Maximum = null);

/// <summary>A plain-text audience description, with optional ONIX list 74 language.</summary>
public sealed record BookOnixAudienceDescription(string Text, string? LanguageCode = null);

/// <summary>Explicit audience assertions; not an automated suitability or reading-level assessment.</summary>
public sealed record BookOnixAudienceMetadata {
    /// <summary>Up to 13 distinct ONIX audience categories, with at most one marked main.</summary>
    public IReadOnlyList<BookOnixAudience> Categories { get; init; } = [];
    /// <summary>Up to 64 distinct additional code or heading assertions; at most one main per audience code type.</summary>
    public IReadOnlyList<BookOnixAudienceCode> Codes { get; init; } = [];
    /// <summary>Up to 14 distinct adult ratings. Requires GeneralAdult; unrated and unrestricted adult assertions must stand alone.</summary>
    public IReadOnlyList<BookOnixAdultAudience> AdultRatings { get; init; } = [];
    /// <summary>At most one range per age type. Interest months and years are mutually exclusive.</summary>
    public IReadOnlyList<BookOnixAgeRange> AgeRanges { get; init; } = [];
    /// <summary>At most one ordered school/college grade range per supported grading system.</summary>
    public IReadOnlyList<BookOnixGradeRange> GradeRanges { get; init; } = [];
    /// <summary>Up to 64 distinct publisher-supplied complexity assertions; no score calculation is performed.</summary>
    public IReadOnlyList<BookOnixComplexity> Complexities { get; init; } = [];
    /// <summary>At most 16 plain-text descriptions. Repeated descriptions require distinct explicit languages.</summary>
    public IReadOnlyList<BookOnixAudienceDescription> Descriptions { get; init; } = [];
}
