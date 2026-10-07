namespace OfficeIMO.Epub;

/// <summary>A named person or organization and EPUB 3 contributor refinements.</summary>
public sealed class EpubContributorMetadata {
    /// <summary>Display name, in its original script and order.</summary>
    public string Name { get; set; } = string.Empty;
    /// <summary>Optional sorting form, such as “Carroll, Lewis”.</summary>
    public string? FileAs { get; set; }
    /// <summary>Optional BCP 47 language of the display name.</summary>
    public string? Language { get; set; }
    /// <summary>
    /// Ordered MARC relator codes, such as aut, edt, ill, trl or nrt. The first is most important.
    /// Codes must contain three lowercase ASCII letters; registry membership is checked by independent validation.
    /// An empty list leaves roles unspecified. Repeated codes are collapsed in first-occurrence order.
    /// </summary>
    public IReadOnlyList<string> MarcRoles { get; set; } = Array.Empty<string>();
}
