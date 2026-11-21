namespace OfficeIMO.Markdown;

/// <summary>The language of a document's front matter, which its fence lines select.</summary>
public enum MarkdownFrontMatterFormat {
    /// <summary>YAML between <c>---</c> lines, as Jekyll, Hugo and most static site generators read it.</summary>
    Yaml,

    /// <summary>TOML between <c>+++</c> lines, as Hugo and Zola read it.</summary>
    Toml
}
