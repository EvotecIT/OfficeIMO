namespace OfficeIMO.Markdown;

/// <summary>
/// Options controlling markdown serialization from the typed document model.
/// </summary>
public sealed class MarkdownSaveOptions {
    /// <summary>Accepts native-save omissions explicitly selected by the writer options.</summary>
    public bool AllowLoss { get; set; }
    /// <summary>Output encoding, or null for UTF-8 without a byte-order mark.</summary>
    public Encoding? Encoding { get; set; }

    private string? _outputLineEnding;
    private char _unorderedListMarker = '-';

    /// <summary>Creates default OfficeIMO-flavored markdown writer options.</summary>
    public static MarkdownSaveOptions CreateOfficeIMOProfile() => new MarkdownSaveOptions();

    /// <summary>Creates writer options for the requested output profile.</summary>
    public static MarkdownSaveOptions CreateProfile(MarkdownOutputProfile profile) =>
        profile switch {
            MarkdownOutputProfile.OfficeIMO => CreateOfficeIMOProfile(),
            MarkdownOutputProfile.CommonMark => CreateCommonMarkProfile(),
            MarkdownOutputProfile.GitHubFlavoredMarkdown => CreateGitHubFlavoredMarkdownProfile(),
            MarkdownOutputProfile.Portable => CreatePortableProfile(),
            _ => throw new ArgumentOutOfRangeException(nameof(profile), profile, "Unknown markdown output profile.")
        };

    /// <summary>
    /// Creates a CommonMark-oriented writer profile. GitHub-only constructs such as pipe tables are
    /// emitted using portable raw HTML fallbacks when a plain CommonMark representation does not exist.
    /// </summary>
    public static MarkdownSaveOptions CreateCommonMarkProfile() {
        var options = CreatePortableProfile();
        options.OutputLineEnding = "\n";
        options.FrontMatterRendering = MarkdownFrontMatterRenderingMode.Omit;
        options.AbbreviationDefinitionRendering = MarkdownAbbreviationDefinitionRenderingMode.Omit;
        MarkdownBlockRenderBuiltInExtensions.AddCommonMarkTableMarkdownFallback(options);
        MarkdownBlockRenderBuiltInExtensions.AddCommonMarkTaskListMarkdownFallback(options);
        MarkdownBlockRenderBuiltInExtensions.AddCommonMarkDefinitionListMarkdownFallback(options);
        MarkdownBlockRenderBuiltInExtensions.AddCommonMarkFootnoteDefinitionMarkdownFallback(options);
        MarkdownBlockRenderBuiltInExtensions.AddCommonMarkDetailsMarkdownFallback(options);
        MarkdownBlockRenderBuiltInExtensions.AddCommonMarkCustomContainerMarkdownFallback(options);
        MarkdownBlockRenderBuiltInExtensions.AddCommonMarkParagraphLineStartMarkdownFallback(options);
        MarkdownBlockRenderBuiltInExtensions.AddCommonMarkAttributedBlockMarkdownFallback(options);
        MarkdownInlineRenderBuiltInExtensions.AddCommonMarkGfmInlineMarkdownFallbacks(options);
        MarkdownInlineRenderBuiltInExtensions.AddCommonMarkAttributedInlineMarkdownFallback(options);
        return options;
    }

    /// <summary>
    /// Creates a GitHub Flavored Markdown-oriented writer profile for README and GitHub documentation output.
    /// </summary>
    public static MarkdownSaveOptions CreateGitHubFlavoredMarkdownProfile() {
        var options = new MarkdownSaveOptions {
            ImageRenderingMode = MarkdownImageRenderingMode.PortableMarkdown,
            AbbreviationDefinitionRendering = MarkdownAbbreviationDefinitionRenderingMode.Omit,
            OutputLineEnding = "\n",
            UnorderedListMarker = '-'
        };
        MarkdownBlockRenderBuiltInExtensions.AddGitHubDefinitionListMarkdownFallback(options);
        MarkdownBlockRenderBuiltInExtensions.AddGitHubCustomContainerMarkdownFallback(options);
        MarkdownBlockRenderBuiltInExtensions.AddGitHubParagraphLineStartMarkdownFallback(options);
        MarkdownBlockRenderBuiltInExtensions.AddGitHubAttributedBlockMarkdownFallback(options);
        MarkdownInlineRenderBuiltInExtensions.AddGitHubOfficeInlineMarkdownFallbacks(options);
        MarkdownInlineRenderBuiltInExtensions.AddGitHubAttributedInlineMarkdownFallback(options);
        return options;
    }

    /// <summary>
    /// Creates a more portable markdown writer profile that degrades OfficeIMO-only blocks into broadly compatible markdown.
    /// </summary>
    public static MarkdownSaveOptions CreatePortableProfile() {
        var options = new MarkdownSaveOptions();
        options.ImageRenderingMode = MarkdownImageRenderingMode.PortableMarkdown;
        MarkdownBlockRenderBuiltInExtensions.AddPortableCalloutMarkdownFallback(options);
        return options;
    }

    /// <summary>
    /// Creates a markdown writer profile that emits raw HTML for image output.
    /// </summary>
    public static MarkdownSaveOptions CreateHtmlImageProfile() => new MarkdownSaveOptions {
        ImageRenderingMode = MarkdownImageRenderingMode.Html
    };

    /// <summary>
    /// Controls how image blocks and image inlines are serialized back to markdown text.
    /// </summary>
    public MarkdownImageRenderingMode ImageRenderingMode { get; set; } = MarkdownImageRenderingMode.RichMarkdown;

    /// <summary>
    /// Controls escaping of literal inline text and unstructured table strings. The default preserves existing
    /// output. <see cref="MarkdownTextEscapingMode.LiteralBackslash"/> prevents HTML and entity interpretation
    /// without introducing HTML character references or escaping ordinary parentheses.
    /// </summary>
    public MarkdownTextEscapingMode TextEscapingMode { get; set; } = MarkdownTextEscapingMode.Default;

    /// <summary>
    /// Controls how YAML front matter is emitted by <see cref="MarkdownDocument.ToMarkdown(MarkdownSaveOptions?)"/>.
    /// </summary>
    public MarkdownFrontMatterRenderingMode FrontMatterRendering { get; set; } = MarkdownFrontMatterRenderingMode.Preserve;

    /// <summary>
    /// Controls how parse-owned abbreviation definitions are emitted by <see cref="MarkdownDocument.ToMarkdown(MarkdownSaveOptions?)"/>.
    /// </summary>
    public MarkdownAbbreviationDefinitionRenderingMode AbbreviationDefinitionRendering { get; set; } = MarkdownAbbreviationDefinitionRenderingMode.Preserve;

    /// <summary>
    /// Optional line ending to use in rendered Markdown. When unset, the platform default is used.
    /// </summary>
    public string? OutputLineEnding {
        get => _outputLineEnding;
        set {
            if (value != null && value.Length == 0) {
                throw new ArgumentException("Output line ending cannot be empty.", nameof(value));
            }

            _outputLineEnding = value;
        }
    }

    /// <summary>
    /// Marker used when rendering unordered list items. CommonMark-compatible values are <c>-</c>, <c>*</c>, and <c>+</c>.
    /// </summary>
    public char UnorderedListMarker {
        get => _unorderedListMarker;
        set {
            if (value != '-' && value != '*' && value != '+') {
                throw new ArgumentOutOfRangeException(nameof(value), value, "Unordered list marker must be '-', '*', or '+'.");
            }

            _unorderedListMarker = value;
        }
    }

    /// <summary>
    /// Optional markdown block render extensions. Later registrations win when block types overlap.
    /// </summary>
    public List<MarkdownBlockMarkdownRenderExtension> BlockRenderExtensions { get; } = new();

    /// <summary>
    /// Optional markdown inline render extensions. Later registrations win when inline types overlap.
    /// </summary>
    public List<MarkdownInlineMarkdownRenderExtension> InlineRenderExtensions { get; } = new();

    /// <summary>
    /// Optional markdown block render extensions that match parsed blocks by final syntax kind.
    /// Later registrations win when syntax kinds overlap. These extensions run before type-based block extensions.
    /// </summary>
    public List<MarkdownSyntaxBlockMarkdownRenderExtension> SyntaxBlockRenderExtensions { get; } = new();

    /// <summary>
    /// Optional markdown inline render extensions that match parsed inlines by final syntax kind.
    /// Later registrations win when syntax kinds overlap. These extensions run before type-based inline extensions.
    /// </summary>
    public List<MarkdownSyntaxInlineMarkdownRenderExtension> SyntaxInlineRenderExtensions { get; } = new();

    /// <summary>
    /// Creates a shallow clone of the writer options while copying mutable collections.
    /// </summary>
    public MarkdownSaveOptions Clone() {
        var clone = new MarkdownSaveOptions {
            AllowLoss = AllowLoss,
            Encoding = Encoding,
            ImageRenderingMode = ImageRenderingMode,
            TextEscapingMode = TextEscapingMode,
            FrontMatterRendering = FrontMatterRendering,
            AbbreviationDefinitionRendering = AbbreviationDefinitionRendering,
            OutputLineEnding = OutputLineEnding,
            UnorderedListMarker = UnorderedListMarker
        };
        for (int i = 0; i < BlockRenderExtensions.Count; i++) {
            clone.BlockRenderExtensions.Add(BlockRenderExtensions[i]);
        }
        for (int i = 0; i < InlineRenderExtensions.Count; i++) {
            clone.InlineRenderExtensions.Add(InlineRenderExtensions[i]);
        }
        for (int i = 0; i < SyntaxBlockRenderExtensions.Count; i++) {
            clone.SyntaxBlockRenderExtensions.Add(SyntaxBlockRenderExtensions[i]);
        }
        for (int i = 0; i < SyntaxInlineRenderExtensions.Count; i++) {
            clone.SyntaxInlineRenderExtensions.Add(SyntaxInlineRenderExtensions[i]);
        }

        return clone;
    }
}
