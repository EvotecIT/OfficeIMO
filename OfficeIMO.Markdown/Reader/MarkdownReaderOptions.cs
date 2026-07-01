namespace OfficeIMO.Markdown;

/// <summary>
/// Controls how OfficeIMO callout headers handle text after the <c>[!KIND]</c> marker.
/// </summary>
public enum MarkdownCalloutTitleMode {
    /// <summary>Parse trailing header text as an OfficeIMO callout title.</summary>
    OfficeIMO,
    /// <summary>Match Markdig's alert extension boundary: only no-title alert headers are parsed as callouts.</summary>
    MarkdigCompatible
}

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
            Abbreviations = false,
            GenericAttributes = false,
            StandaloneImageBlocks = false,
            StrictListIndentation = true,
            PreserveHtmlBlockBlankLineContent = false,
            Subscript = false,
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
            AllowHeaderlessTables = false,
            ParseTableCellBlocks = false,
            DefinitionLists = false,
            TocPlaceholders = false,
            Footnotes = true,
            Abbreviations = false,
            GenericAttributes = false,
            StandaloneImageBlocks = false,
            StrictListIndentation = true,
            PreserveHtmlBlockBlankLineContent = false,
            SingleTildeStrikethrough = true,
            Subscript = false,
            CjkFriendlyEmphasis = false,
            AutolinkUrls = true,
            AutolinkAllowDomainWithoutPeriod = false,
            AutolinkAllowQueryAndFragmentSpecialCharacters = true,
            AutolinkAllowBalancedParenthesesWithTrailingPunctuation = true,
            AutolinkRequireLowercaseWwwPrefix = true,
            AutolinkRequireLowercaseBareSchemePrefix = true,
            AutolinkBareSchemeUrls = true,
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
            Abbreviations = false,
            GenericAttributes = false,
            StandaloneImageBlocks = false,
            StrictListIndentation = true,
            Subscript = false,
            AutolinkUrls = false,
            AutolinkWwwUrls = false,
            AutolinkEmails = false
        };
    }

    /// <summary>Enable YAML front matter parsing at the very top of the file.</summary>
    public bool FrontMatter { get; set; } = true;
    /// <summary>Enable recognition of Docs-style callouts ("> [!KIND] Title" blocks).</summary>
    public bool Callouts { get; set; } = true;
    /// <summary>
    /// Controls whether trailing callout header text is parsed as an OfficeIMO title or left as ordinary blockquote text for Markdig-compatible alert parsing.
    /// </summary>
    public MarkdownCalloutTitleMode CalloutTitleMode { get; set; } = MarkdownCalloutTitleMode.OfficeIMO;
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
    /// Enable Markdig-style extra ordered-list markers such as alphabetic <c>a.</c>/<c>A.</c>
    /// and lower/upper roman markers up to <c>xxxix</c>.
    /// Default: <c>false</c>; opt in when mirroring Markdig's <c>UseListExtras</c> extension.
    /// </summary>
    public bool ListExtras { get; set; } = false;
    /// <summary>
    /// When <c>true</c>, nested list levels are derived from continuation-indent rules instead of the library's
    /// more permissive legacy indentation heuristic. Enabled by stricter compatibility profiles.
    /// </summary>
    public bool StrictListIndentation { get; set; } = false;
    /// <summary>Enable pipe tables with optional header + alignment row.</summary>
    public bool Tables { get; set; } = true;
    /// <summary>
    /// When <c>true</c>, pipe rows without a GFM delimiter/alignment row can be parsed as OfficeIMO headerless tables.
    /// Disable this for stricter GitHub Flavored Markdown compatibility, where a delimiter row is required.
    /// </summary>
    public bool AllowHeaderlessTables { get; set; } = true;
    /// <summary>
    /// When <c>true</c>, table cells that look like nested markdown can be upgraded to structured block content.
    /// Disable this for GitHub Flavored Markdown compatibility, where table cells contain inline content only.
    /// </summary>
    public bool ParseTableCellBlocks { get; set; } = true;
    /// <summary>Enable definition lists (Term: Definition lines).</summary>
    public bool DefinitionLists { get; set; } = true;
    /// <summary>Enable placeholder TOC markers such as <c>[TOC]</c> and <c>{:toc}</c>.</summary>
    public bool TocPlaceholders { get; set; } = true;
    /// <summary>Enable footnote references and footnote definition blocks.</summary>
    public bool Footnotes { get; set; } = true;
    /// <summary>
    /// Enable Markdig-style abbreviation definitions and inline expansion, e.g.
    /// <c>*[HTML]: Hyper Text Markup Language</c> followed by <c>HTML</c>.
    /// Default: <c>false</c>; opt in when mirroring Markdig's <c>UseAbbreviations</c> extension.
    /// </summary>
    public bool Abbreviations { get; set; } = false;
    /// <summary>
    /// Enable Markdig-style generic trailing attribute blocks on supported Markdown elements,
    /// for example <c># Heading {#id .wide key="value"}</c>.
    /// Default: <c>false</c>; opt in when mirroring Markdig's <c>UseGenericAttributes</c> extension.
    /// </summary>
    public bool GenericAttributes { get; set; } = false;
    /// <summary>
    /// Enable Markdig-style colon-fenced custom containers such as <c>::: note</c>.
    /// Default: <c>false</c>; opt in when mirroring Markdig's <c>UseCustomContainers</c> extension.
    /// </summary>
    public bool CustomContainers { get; set; } = false;
    /// <summary>
    /// When <c>true</c>, GitHub Flavored Markdown-style single-tilde strikethrough (<c>~text~</c>) is enabled.
    /// Default: <c>false</c>.
    /// </summary>
    public bool SingleTildeStrikethrough { get; set; } = false;
    /// <summary>
    /// When <c>true</c>, Markdig emphasis-extra subscript (<c>~text~</c>) is enabled.
    /// This is disabled by strict CommonMark and GFM profiles because GFM can use the same
    /// delimiter for single-tilde strikethrough.
    /// </summary>
    public bool Subscript { get; set; } = true;

    /// <summary>
    /// When <c>true</c>, star emphasis uses Markdig's CJK-friendly delimiter behavior so emphasis can
    /// open or close next to CJK text and CJK punctuation without relaxing underscore intraword rules.
    /// This mirrors Markdig's <c>UseCjkFriendlyEmphasis</c> pipeline option.
    /// </summary>
    public bool CjkFriendlyEmphasis { get; set; } = false;

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
    /// <summary>
    /// When <c>true</c>, selected OfficeIMO-friendly raw HTML containers such as <c>table</c> and <c>details</c>
    /// can continue across blank lines until their matching closing tag. Strict CommonMark/GFM profiles disable
    /// this because CommonMark HTML block types 6 and 7 end at the next blank line.
    /// </summary>
    public bool PreserveHtmlBlockBlankLineContent { get; set; } = true;
    /// <summary>Enable paragraph parsing and basic inlines.</summary>
    public bool Paragraphs { get; set; } = true;

    /// <summary>
    /// When <c>true</c>, auto-detects plain <c>http(s)://...</c> URLs in text and turns them into links.
    /// Default: <c>true</c>.
    /// </summary>
    public bool AutolinkUrls { get; set; } = true;

    /// <summary>
    /// When <c>true</c>, bare URL autolinks may target hosts without a period, such as
    /// <c>https://localhost</c> or <c>www.local</c>. OfficeIMO defaults to <c>true</c> for
    /// compatibility with existing documents; stricter Markdig/GFM-style profiles can set
    /// this to <c>false</c>.
    /// </summary>
    public bool AutolinkAllowDomainWithoutPeriod { get; set; } = true;

    /// <summary>
    /// When <c>true</c>, bare URL autolinks may keep balanced parentheses and ampersands inside query
    /// strings or fragments. OfficeIMO defaults to <c>false</c> for compatibility with existing conservative
    /// parsing; the GitHub Flavored Markdown profile enables this to match Markdig/GFM autolink behavior.
    /// </summary>
    public bool AutolinkAllowQueryAndFragmentSpecialCharacters { get; set; } = false;

    /// <summary>
    /// When <c>true</c>, bare URL autolinks may keep a balanced parenthesized segment inside
    /// the link while leaving an extra closing parenthesis or trailing punctuation outside the
    /// link. OfficeIMO defaults to <c>false</c> for compatibility with its older conservative
    /// parser; the GitHub Flavored Markdown profile enables this for Markdig/GFM-style autolinks.
    /// </summary>
    public bool AutolinkAllowBalancedParenthesesWithTrailingPunctuation { get; set; } = false;

    /// <summary>
    /// When <c>true</c>, bare URL autolinks may keep final punctuation such as <c>.</c>, <c>,</c>,
    /// <c>;</c>, <c>!</c>, or <c>?</c> inside the link when the next source character is a closing
    /// parenthesis outside the URL. Markdig <c>UseAutoLinks</c> keeps that punctuation; cmark-gfm
    /// trims at least the period in the comparable GFM case, so this remains opt-in.
    /// </summary>
    public bool AutolinkAllowTrailingPunctuationBeforeClosingParenthesis { get; set; } = false;

    /// <summary>
    /// When <c>true</c>, bare URL autolinks trim at most one final punctuation character or
    /// underscore from the parsed target. Markdig <c>UseAutoLinks</c> keeps earlier repeated
    /// punctuation inside the link; OfficeIMO's legacy behavior trims the full trailing run.
    /// </summary>
    public bool AutolinkTrimSingleTrailingPunctuationOrUnderscore { get; set; } = false;

    /// <summary>
    /// When <c>true</c>, bare URL autolinks keep trailing semicolons inside the consumed
    /// URL target. Markdig <c>UseAutoLinks</c> keeps semicolons while still trimming
    /// other single trailing punctuation such as periods and commas.
    /// </summary>
    public bool AutolinkKeepTrailingSemicolonPunctuation { get; set; } = false;

    /// <summary>
    /// When <c>true</c>, bare <c>www.</c> autolinks require the prefix itself to be lowercase.
    /// The host portion after the prefix may still use mixed case.
    /// </summary>
    public bool AutolinkRequireLowercaseWwwPrefix { get; set; } = false;

    /// <summary>
    /// When <c>true</c>, bare <c>www.</c> autolinks reject host labels containing underscores.
    /// This mirrors Markdig <c>UseAutoLinks</c> while leaving OfficeIMO's older permissive
    /// behavior available for existing consumers.
    /// </summary>
    public bool AutolinkRejectUnderscoreInWwwHost { get; set; } = false;

    /// <summary>
    /// When <c>true</c>, bare URL autolinks with an authority such as <c>https://</c> or
    /// <c>ftp://</c> reject host labels containing underscores. Markdig <c>UseAutoLinks</c>
    /// leaves those URL-shaped tokens as literal text; OfficeIMO's legacy behavior can still
    /// link them.
    /// </summary>
    public bool AutolinkRejectUnderscoreInUrlHost { get; set; } = false;

    /// <summary>
    /// When <c>true</c>, bare URL autolinks reject authority/user-info forms containing <c>@</c>,
    /// such as <c>https://user@example.com/path</c>. Markdig <c>UseAutoLinks</c> leaves those
    /// tokens as literal text; OfficeIMO's legacy behavior can still link them.
    /// </summary>
    public bool AutolinkRejectUserInfoAuthority { get; set; } = false;

    /// <summary>
    /// When <c>true</c>, bare URL autolinks may include a closing square bracket <c>]</c>
    /// in the consumed URL target. Markdig <c>UseAutoLinks</c> keeps this character inside
    /// the link; OfficeIMO's legacy behavior stops before it to avoid crossing bracketed text.
    /// </summary>
    public bool AutolinkAllowClosingBracketInUrl { get; set; } = false;

    /// <summary>
    /// When <c>true</c>, bare URL autolinks keep trailing single or double quote characters
    /// inside the consumed URL target. Markdig <c>UseAutoLinks</c> keeps these quote
    /// characters when no matching opening quote prevents the autolink.
    /// </summary>
    public bool AutolinkKeepTrailingQuotePunctuation { get; set; } = false;

    /// <summary>
    /// When <c>true</c>, bare scheme autolinks such as <c>mailto:</c>, <c>ftp://</c>, and
    /// <c>tel:</c> require the scheme prefix itself to be lowercase.
    /// </summary>
    public bool AutolinkRequireLowercaseBareSchemePrefix { get; set; } = false;

    /// <summary>
    /// When <c>true</c>, a bare <c>mailto:user@example.com</c> autolink displays only
    /// <c>user@example.com</c> while keeping the link target as <c>mailto:user@example.com</c>.
    /// Markdig <c>UseAutoLinks</c> uses the address-only display; cmark-gfm keeps the full
    /// <c>mailto:</c> source text as the display label.
    /// </summary>
    public bool AutolinkBareMailtoDisplayAddressOnly { get; set; } = false;

    /// <summary>
    /// When <c>true</c>, bare <c>mailto:</c> autolinks use Markdig-style semicolon handling:
    /// address-only tokens followed by a semicolon remain literal, while path, query, or
    /// fragment targets keep trailing semicolons inside the link.
    /// </summary>
    public bool AutolinkBareMailtoMarkdigSemicolonHandling { get; set; } = false;

    /// <summary>
    /// Optional previous-character allow-list for bare URL/email autolinks. When set, a bare
    /// autolink may start at the beginning of text, after whitespace, or after one of these
    /// characters. When <c>null</c>, OfficeIMO's legacy boundary heuristic is used.
    /// </summary>
    public string? AutolinkValidPreviousCharacters { get; set; }

    /// <summary>
    /// When <c>true</c>, auto-detects selected bare URI schemes such as <c>mailto:</c>, <c>ftp://</c>,
    /// <c>tel:</c>, and <c>xmpp:</c>. Use <see cref="AutolinkBareSchemePrefixes"/> to narrow the
    /// scheme set for a compatibility profile.
    /// Default: <c>false</c>; enabled by <see cref="CreateGitHubFlavoredMarkdownProfile"/>.
    /// </summary>
    public bool AutolinkBareSchemeUrls { get; set; } = false;

    /// <summary>
    /// Optional bare-scheme prefix allow-list used when <see cref="AutolinkBareSchemeUrls"/> is enabled.
    /// Prefixes should include their punctuation, for example <c>mailto:</c>, <c>ftp://</c>, or <c>tel:</c>.
    /// When <c>null</c>, OfficeIMO's built-in selected scheme set is used.
    /// </summary>
    public string[]? AutolinkBareSchemePrefixes { get; set; }

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
    /// When <c>true</c>, ordinary soft line breaks inside paragraphs are parsed as hard line breaks.
    /// This mirrors Markdig's <c>UseSoftlineBreakAsHardlineBreak</c> pipeline option without changing
    /// the default CommonMark/GFM soft-break behavior.
    /// </summary>
    public bool SoftLineBreaksAsHardLineBreaks { get; set; } = false;

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
    /// When <c>true</c>, syntax-backed parse results retain the raw markdown input beside the normalized source
    /// text used for source spans. This is groundwork for lossless trivia/roundtrip support; semantic
    /// <see cref="MarkdownDoc.ToMarkdown(MarkdownWriteOptions?)"/> output remains normalized markdown generation.
    /// </summary>
    public bool PreserveTrivia { get; set; } = false;

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
    /// Optional ordered post-parse inline AST transforms.
    /// These run after built-in inline parsing and input normalization, before document-level transforms.
    /// </summary>
    public List<MarkdownInlineTransformExtension> InlineTransformExtensions { get; } = new();

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
