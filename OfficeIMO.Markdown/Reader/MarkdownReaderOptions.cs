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
    /// <summary>Enable indented code blocks (lines indented by 4 spaces).</summary>
    public bool IndentedCodeBlocks { get; set; } = true;
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
    /// Enable raw HTML blocks. When set to <c>false</c>, block-level HTML is preserved as plain text so that readers can postprocess
    /// or render it verbatim.
    /// </summary>
    public bool HtmlBlocks { get; set; } = true;
    /// <summary>Enable paragraph parsing and basic inlines.</summary>
    public bool Paragraphs { get; set; } = true;

    /// <summary>
    /// When <c>true</c>, auto-detects plain <c>http(s)://...</c> URLs in text and turns them into links.
    /// Default: <c>true</c>.
    /// </summary>
    public bool AutolinkUrls { get; set; } = true;

    /// <summary>
    /// When <c>true</c>, auto-detects plain <c>www.example.com</c> URLs in text and turns them into links.
    /// Default: <c>true</c>.
    /// </summary>
    public bool AutolinkWwwUrls { get; set; } = true;

    /// <summary>
    /// Scheme prefix to use for <see cref="AutolinkWwwUrls"/> (for example <c>https://</c>).
    /// Default: <c>https://</c>.
    /// </summary>
    public string AutolinkWwwScheme { get; set; } = "https://";

    /// <summary>
    /// When <c>true</c>, auto-detects plain emails in text (for example <c>user@example.com</c>) and turns them into <c>mailto:</c> links.
    /// Default: <c>true</c>.
    /// </summary>
    public bool AutolinkEmails { get; set; } = true;

    /// <summary>
    /// When <c>true</c>, a trailing backslash at the end of a paragraph line is treated as a hard line break (like GitHub/CommonMark).
    /// This is in addition to the "two trailing spaces" hard break form.
    /// Default: <c>true</c>.
    /// </summary>
    public bool BackslashHardBreaks { get; set; } = true;
    /// <summary>
    /// Enable inline HTML interpretations (e.g. &lt;br&gt;, &lt;u&gt;...&lt;/u&gt;). When disabled, HTML tags remain literal text and no HTML
    /// decoding is performed.
    /// </summary>
    /// <example>
    /// <code>
    /// var options = new MarkdownReaderOptions {
    ///     HtmlBlocks = false,
    ///     InlineHtml = false,
    /// };
    /// // MarkdownReader.Read("<div>hello<br/></div>", options) keeps the HTML tokens inside text runs.
    /// </code>
    /// </example>
    public bool InlineHtml { get; set; } = true;

    /// <summary>
    /// Optional base URI used to resolve relative links/images. When set, relative URLs (not starting with http/https,//,#,mailto:,data:)
    /// are converted to absolute using this base during parsing.
    /// </summary>
    public string? BaseUri { get; set; }

    /// <summary>
    /// When <c>true</c>, blocks scriptable URL schemes (e.g. <c>javascript:</c>, <c>vbscript:</c>) during parsing.
    /// If a link/image uses a blocked scheme, it is treated as plain text instead of producing a clickable/linkable node.
    /// Default: <c>true</c>.
    /// </summary>
    public bool DisallowScriptUrls { get; set; } = true;

    /// <summary>
    /// When <c>true</c>, blocks <c>file:</c> URLs (and Windows drive-like <c>C:\</c> paths) during parsing.
    /// Default: <c>false</c> to preserve legacy behavior for local/offline documents.
    /// </summary>
    public bool DisallowFileUrls { get; set; } = false;

    /// <summary>
    /// When <c>false</c>, <c>mailto:</c> links are treated as plain text. Default: <c>true</c>.
    /// </summary>
    public bool AllowMailtoUrls { get; set; } = true;

    /// <summary>
    /// When <c>false</c>, <c>data:</c> URLs are treated as plain text. Default: <c>true</c>.
    /// </summary>
    public bool AllowDataUrls { get; set; } = true;

    /// <summary>
    /// When <c>false</c>, protocol-relative URLs (<c>//example.com</c>) are treated as plain text. Default: <c>true</c>.
    /// </summary>
    public bool AllowProtocolRelativeUrls { get; set; } = true;
}
