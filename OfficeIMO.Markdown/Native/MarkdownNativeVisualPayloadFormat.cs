namespace OfficeIMO.Markdown;

/// <summary>
/// Lightweight visual payload format classification.
/// </summary>
public enum MarkdownNativeVisualPayloadFormat {
    /// <summary>Payload format is unknown or empty.</summary>
    Unknown,
    /// <summary>Plain text payload.</summary>
    Text,
    /// <summary>JSON object payload.</summary>
    JsonObject,
    /// <summary>JSON array payload.</summary>
    JsonArray,
    /// <summary>Mermaid diagram payload.</summary>
    Mermaid
}
