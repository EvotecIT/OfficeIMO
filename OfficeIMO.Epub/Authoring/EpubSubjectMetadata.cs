namespace OfficeIMO.Epub;

/// <summary>A human-readable subject with optional classification authority and code.</summary>
public sealed class EpubSubjectMetadata {
    /// <summary>Human-readable heading.</summary>
    public string Text { get; set; } = string.Empty;
    /// <summary>Optional authority, such as BISAC or Thema. Vocabulary membership is not inferred.</summary>
    public string? Authority { get; set; }
    /// <summary>Optional subject code. A code requires an authority.</summary>
    public string? Code { get; set; }
    /// <summary>Optional BCP 47 language of the heading.</summary>
    public string? Language { get; set; }
}
