namespace OfficeIMO.Markdown;

/// <summary>
/// Named block parser extension registration used by <see cref="MarkdownReaderPipeline.Default(MarkdownReaderOptions?)"/>.
/// </summary>
public sealed class MarkdownBlockParserExtension {
    /// <summary>
    /// Creates a block parser extension registration from a delegate-based parser.
    /// </summary>
    public MarkdownBlockParserExtension(
        string name,
        MarkdownBlockParserPlacement placement,
        MarkdownBlockParser parser,
        Func<MarkdownReaderOptions, bool>? isEnabled = null)
        : this(name, placement, new DelegateMarkdownBlockParser(parser), isEnabled) {
    }

    /// <summary>
    /// Creates a block parser extension registration.
    /// </summary>
    public MarkdownBlockParserExtension(
        string name,
        MarkdownBlockParserPlacement placement,
        IMarkdownBlockParser parser,
        Func<MarkdownReaderOptions, bool>? isEnabled = null) {
        if (string.IsNullOrWhiteSpace(name)) {
            throw new ArgumentException("Extension name cannot be null or whitespace.", nameof(name));
        }

        Name = name.Trim();
        Placement = placement;
        Parser = parser ?? throw new ArgumentNullException(nameof(parser));
        IsEnabled = isEnabled;
    }

    /// <summary>Stable extension name used for inspection or de-duplication.</summary>
    public string Name { get; }

    /// <summary>Placement anchor within the default reader pipeline.</summary>
    public MarkdownBlockParserPlacement Placement { get; }

    /// <summary>Parser instance contributed by this extension.</summary>
    public IMarkdownBlockParser Parser { get; }

    /// <summary>
    /// Optional predicate that decides whether the extension should apply for a specific options instance.
    /// </summary>
    public Func<MarkdownReaderOptions, bool>? IsEnabled { get; }

    /// <summary>Returns <see langword="true"/> when the extension should participate in parsing.</summary>
    public bool AppliesTo(MarkdownReaderOptions options) => IsEnabled?.Invoke(options) ?? true;
}
