namespace OfficeIMO.Markdown;

/// <summary>
/// Controls how document front matter is emitted during Markdown serialization.
/// </summary>
public enum MarkdownFrontMatterRenderingMode {
    /// <summary>Preserve YAML front matter fences and entries.</summary>
    Preserve,

    /// <summary>Omit front matter from the serialized Markdown document.</summary>
    Omit
}
