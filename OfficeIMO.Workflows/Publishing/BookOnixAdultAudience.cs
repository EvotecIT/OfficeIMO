namespace OfficeIMO.Workflows;

/// <summary>Publisher-supplied ONIX list 203 adult-audience ratings.</summary>
public enum BookOnixAdultAudienceRating {
    /// <summary>Publisher explicitly identifies the product as unrated (00).</summary>
    Unrated,
    /// <summary>Publisher considers the content suitable for any adult audience (01).</summary>
    AnyAdultAudience,
    /// <summary>General advice about potentially distressing or offensive content (02).</summary>
    ContentAdvice,
    /// <summary>Explicit sexual content (03).</summary>
    SexualContent,
    /// <summary>Extreme violence (04).</summary>
    Violence,
    /// <summary>Severe drug or alcohol misuse (05).</summary>
    DrugsOrAlcohol,
    /// <summary>Extreme or offensive language (06).</summary>
    OffensiveLanguage,
    /// <summary>Severe intolerance or abuse targeting social groups (07).</summary>
    Intolerance,
    /// <summary>Sexual or extreme domestic abuse (08).</summary>
    Abuse,
    /// <summary>Severe self-harm, including serious eating disorders (09).</summary>
    SelfHarm,
    /// <summary>Extreme cruelty to animals (10).</summary>
    AnimalCruelty,
    /// <summary>Serious physical or mental illness (11).</summary>
    Illness,
    /// <summary>Severe distress associated with death or grief (12).</summary>
    DeathAndGrief,
    /// <summary>Content relating to suicide (13).</summary>
    Suicide
}

/// <summary>An explicit publisher rating; no content analysis or statutory classification is performed.</summary>
public sealed record BookOnixAdultAudience(BookOnixAdultAudienceRating Rating, bool IsMain = false) {
    /// <summary>Up to 16 optional plain-text equivalents. Repeated headings require distinct explicit languages.</summary>
    public IReadOnlyList<BookOnixAudienceHeading> Headings { get; init; } = [];
}
