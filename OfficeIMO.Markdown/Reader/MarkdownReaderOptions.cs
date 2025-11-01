namespace OfficeIMO.Markdown;

/// <summary>
/// Options for the Markdown reader. Feature toggles mirror the blocks and inlines
/// that OfficeIMO.Markdown produces, so generated Markdown can be parsed back predictably.
/// </summary>
public sealed class MarkdownReaderOptions {
    /// <summary>Enable YAML front matter parsing at the very top of the file.</summary>
    public bool FrontMatter { get; set; } = true;
    /// <summary>Enable recognition of Docs-style callouts ("> [!KIND] Title" blocks).</summary>
    public bool Callouts { get; set; } = true;
    /// <summary>Enable ATX headings (#, ##, ...).</summary>
    public bool Headings { get; set; } = true;
    /// <summary>Enable fenced code blocks (```lang ... ```), including caption on the following _line_ if present.</summary>
    public bool FencedCode { get; set; } = true;
    /// <summary>Enable images (standalone lines) with optional caption on the next _italic_ line.</summary>
    public bool Images { get; set; } = true;
    /// <summary>Enable unordered lists and task lists.</summary>
    public bool UnorderedLists { get; set; } = true;
    /// <summary>Enable ordered (numbered) lists.</summary>
    public bool OrderedLists { get; set; } = true;
    /// <summary>Enable pipe tables with optional header + alignment row.</summary>
    public bool Tables { get; set; } = true;
    /// <summary>Enable definition lists (Term: Definition lines).</summary>
    public bool DefinitionLists { get; set; } = true;
    /// <summary>
    /// Enable raw HTML blocks. When set to <c>false</c>, block-level HTML is preserved as plain text.
    /// </summary>
    public bool HtmlBlocks { get; set; } = true;
    /// <summary>Enable paragraph parsing and basic inlines.</summary>
    public bool Paragraphs { get; set; } = true;
    /// <summary>
    /// Enable inline HTML interpretations (e.g. &lt;br&gt;, &lt;u&gt;...&lt;/u&gt;). When disabled, HTML tags remain literal text.
    /// </summary>
    public bool InlineHtml { get; set; } = true;
}
