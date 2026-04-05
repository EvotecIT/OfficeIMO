namespace OfficeIMO.Markdown;

/// <summary>
/// Options for the Markdown reader. Profiles and feature toggles shape the generic markdown core,
/// while <see cref="BlockParserExtensions"/> and <see cref="InlineParserExtensions"/> control opt-in
/// syntax such as OfficeIMO callouts, TOC placeholders, footnotes, or custom inline tokens.
/// </summary>
public sealed class MarkdownReaderOptions {
    /// <summary>
    /// Creates a new OfficeIMO-flavored reader configuration with the built-in block syntax extensions registered.
    /// </summary>
    public MarkdownReaderOptions() : this(seedBuiltInBlockParserExtensions: true) {
    }

    private MarkdownReaderOptions(bool seedBuiltInBlockParserExtensions) {
        if (seedBuiltInBlockParserExtensions) {
            MarkdownReaderBuiltInExtensions.RegisterOfficeIMODefaults(this);
        }
    }

    /// <summary>Named reader profiles for common markdown compatibility targets.</summary>
    public enum MarkdownDialectProfile {
        /// <summary>OfficeIMO defaults including host-oriented extensions.</summary>
        OfficeIMO,
        /// <summary>CommonMark-style core markdown without host-specific extensions.</summary>
        CommonMark,
        /// <summary>GitHub Flavored Markdown-style extensions without OfficeIMO host syntax.</summary>
        GitHubFlavoredMarkdown,
        /// <summary>Portable OfficeIMO subset for stricter or parity-sensitive hosts.</summary>
        Portable
    }

    /// <summary>Creates a reader configuration for the requested dialect/profile.</summary>
    public static MarkdownReaderOptions CreateProfile(MarkdownDialectProfile profile) =>
        profile switch {
            MarkdownDialectProfile.OfficeIMO => CreateOfficeIMOProfile(),
            MarkdownDialectProfile.CommonMark => CreateCommonMarkProfile(),
            MarkdownDialectProfile.GitHubFlavoredMarkdown => CreateGitHubFlavoredMarkdownProfile(),
            MarkdownDialectProfile.Portable => CreatePortableProfile(),
            _ => throw new ArgumentOutOfRangeException(nameof(profile), profile, "Unknown markdown reader profile.")
        };

    /// <summary>
    /// Creates the explicit OfficeIMO profile. This mirrors the library defaults and keeps host-oriented
    /// extensions such as callouts, TOC placeholders, and footnotes enabled.
    /// </summary>
    public static MarkdownReaderOptions CreateOfficeIMOProfile() => new MarkdownReaderOptions();

    /// <summary>
    /// Creates a CommonMark-style core profile. This disables OfficeIMO-only and GFM-style extensions such as
    /// front matter, task lists, tables, definition lists, TOC placeholders, and footnotes.
    /// </summary>
    public static MarkdownReaderOptions CreateCommonMarkProfile() {
        return new MarkdownReaderOptions(seedBuiltInBlockParserExtensions: false) {
            FrontMatter = false,
            Callouts = false,
            TaskLists = false,
            Tables = false,
            DefinitionLists = false,
            TocPlaceholders = false,
            Footnotes = false,
            StandaloneImageBlocks = false,
            StrictListIndentation = true,
            AutolinkUrls = false,
            AutolinkWwwUrls = false,
            AutolinkEmails = false
        };
    }

    /// <summary>
    /// Creates a GitHub Flavored Markdown-style profile. This keeps CommonMark core behavior plus
    /// tables, task lists, and footnotes, while disabling OfficeIMO-only callouts and TOC placeholders.
    /// </summary>
    public static MarkdownReaderOptions CreateGitHubFlavoredMarkdownProfile() {
        var options = new MarkdownReaderOptions(seedBuiltInBlockParserExtensions: false) {
            FrontMatter = false,
            Callouts = false,
            TaskLists = true,
            Tables = true,
            DefinitionLists = false,
            TocPlaceholders = false,
            Footnotes = true,
            StandaloneImageBlocks = false,
            StrictListIndentation = true,
            SingleTildeStrikethrough = true,
            AutolinkUrls = true,
            AutolinkWwwUrls = true,
            AutolinkWwwScheme = "http://",
            AutolinkEmails = true
        };

        MarkdownReaderBuiltInExtensions.AddFootnotes(options);
        return options;
    }

    /// <summary>
    /// Creates a reader configuration for portable Markdown behavior across stricter hosts.
    /// Bare <c>http(s)://...</c>, <c>www.*</c>, and plain email tokens remain literal text, and
    /// OfficeIMO-specific extensions such as Docs-style callouts, TOC placeholders, and footnotes are disabled.
    /// Explicit Markdown links, angle-bracket autolinks, and plain unordered lists continue to work.
    /// </summary>
    public static MarkdownReaderOptions CreatePortableProfile() {
        return new MarkdownReaderOptions(seedBuiltInBlockParserExtensions: false) {
            Callouts = false,
            TaskLists = false,
            TocPlaceholders = false,
            Footnotes = false,
            StandaloneImageBlocks = false,
            StrictListIndentation = true,
            AutolinkUrls = false,
            AutolinkWwwUrls = false,
            AutolinkEmails = false
        };
    }

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
    /// <summary>
    /// When <c>true</c>, a line containing only markdown image syntax (optionally with a following caption line)
    /// is promoted into a typed <see cref="ImageBlock"/> instead of remaining a paragraph with inline image content.
    /// Default: <c>true</c> for OfficeIMO-oriented parsing.
    /// </summary>
    public bool StandaloneImageBlocks { get; set; } = true;
    /// <summary>Enable unordered lists and task lists.</summary>
    public bool UnorderedLists { get; set; } = true;
    /// <summary>Enable task list checkbox parsing inside unordered and ordered list items.</summary>
    public bool TaskLists { get; set; } = true;
    /// <summary>Enable ordered (numbered) lists.</summary>
    public bool OrderedLists { get; set; } = true;
    /// <summary>
    /// When <c>true</c>, nested list levels are derived from continuation-indent rules instead of the library's
    /// more permissive legacy indentation heuristic. Enabled by stricter compatibility profiles.
    /// </summary>
    public bool StrictListIndentation { get; set; } = false;
    /// <summary>Enable pipe tables with optional header + alignment row.</summary>
    public bool Tables { get; set; } = true;
    /// <summary>Enable definition lists (Term: Definition lines).</summary>
    public bool DefinitionLists { get; set; } = true;
    /// <summary>Enable placeholder TOC markers such as <c>[TOC]</c> and <c>{:toc}</c>.</summary>
    public bool TocPlaceholders { get; set; } = true;
    /// <summary>Enable footnote references and footnote definition blocks.</summary>
    public bool Footnotes { get; set; } = true;
    /// <summary>
    /// When <c>true</c>, GitHub Flavored Markdown-style single-tilde strikethrough (<c>~text~</c>) is enabled.
    /// Default: <c>false</c>.
    /// </summary>
    public bool SingleTildeStrikethrough { get; set; } = false;
    /// <summary>
    /// When <c>true</c>, isolated single-line <c>Term: Definition</c> patterns stay as narrative paragraphs.
    /// Consecutive definition-like lines still parse as a definition list.
    /// </summary>
    public bool PreferNarrativeSingleLineDefinitions { get; set; } = false;
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

    /// <summary>
    /// When <c>true</c>, only URL schemes listed in <see cref="AllowedUrlSchemes"/> are allowed.
    /// Relative URLs and fragments are still allowed.
    /// Default: <c>false</c> to preserve legacy behavior.
    /// </summary>
    public bool RestrictUrlSchemes { get; set; } = false;

    /// <summary>
    /// List of allowed URL schemes when <see cref="RestrictUrlSchemes"/> is enabled.
    /// Values are compared case-insensitively and should not include the trailing colon.
    /// Default: <c>http</c>, <c>https</c>, <c>mailto</c>.
    /// </summary>
    public string[] AllowedUrlSchemes { get; set; } = new[] { "http", "https", "mailto" };

    /// <summary>
    /// Optional markdown input normalization before parsing. Defaults are conservative (no transformations).
    /// </summary>
    public MarkdownInputNormalizationOptions InputNormalization { get; set; } = new MarkdownInputNormalizationOptions();

    /// <summary>
    /// Optional maximum input length, in characters, accepted by <see cref="MarkdownReader"/>.
    /// When set and exceeded, parsing fails fast with an <see cref="ArgumentOutOfRangeException"/>.
    /// </summary>
    public int? MaxInputCharacters { get; set; }

    /// <summary>
    /// Optional language-based fenced block factories that can produce specialized AST nodes
    /// instead of plain <see cref="CodeBlock"/> instances.
    /// Later registrations win when languages overlap.
    /// </summary>
    public List<MarkdownFencedBlockExtension> FencedBlockExtensions { get; } = new();

    /// <summary>
    /// Optional block parser extensions layered into the default reader pipeline at named placement anchors.
    /// Profiles use this to opt into OfficeIMO/GFM-style non-core block syntax such as callouts, TOC placeholders, and footnotes.
    /// </summary>
    public List<MarkdownBlockParserExtension> BlockParserExtensions { get; } = new();

    /// <summary>
    /// Optional ordered inline parser extensions that get a chance to recognize custom inline tokens before
    /// the built-in inline parser handles the current position.
    /// </summary>
    public List<MarkdownInlineParserExtension> InlineParserExtensions { get; } = new();

    /// <summary>
    /// Optional ordered post-parse document transforms.
    /// Use these for AST-level upgrades and host-specific semantic rewrites after markdown has been parsed.
    /// </summary>
    /// <example>
    /// <code>
    /// var options = MarkdownReaderOptions.CreatePortableProfile();
    /// options.DocumentTransforms.Add(
    ///     new MarkdownJsonVisualCodeBlockTransform(MarkdownVisualFenceLanguageMode.GenericSemanticFence));
    ///
    /// var document = MarkdownReader.Parse(markdown, options);
    /// </code>
    /// </example>
    public List<IMarkdownDocumentTransform> DocumentTransforms { get; } = new();
}
