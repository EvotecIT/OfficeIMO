namespace OfficeIMO.Epub;

/// <summary>Publisher-supplied discovery metadata. Values describe reviewed content; they do not certify accessibility.</summary>
public sealed class EpubAccessibilityMetadata {
    /// <summary>Sensory modes required by the content, for example textual or visual.</summary>
    public IReadOnlyList<string> AccessModes { get; set; } = Array.Empty<string>();
    /// <summary>Alternative sets of sufficient modes; each inner list is one combination, not a separate claim per mode.</summary>
    public IReadOnlyList<IReadOnlyList<string>> SufficientAccessModes { get; set; } = Array.Empty<IReadOnlyList<string>>();
    /// <summary>Reviewed accessibility features using the schema.org discovery vocabulary.</summary>
    public IReadOnlyList<string> Features { get; set; } = Array.Empty<string>();
    /// <summary>Reviewed hazards using the schema.org discovery vocabulary. No absence-of-hazards claim is inferred.</summary>
    public IReadOnlyList<string> Hazards { get; set; } = Array.Empty<string>();
    /// <summary>Human-readable account of accessibility features and known limitations; null omits the summary.</summary>
    public string? Summary { get; set; }
}
