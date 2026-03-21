namespace OfficeIMO.Markdown;

/// <summary>
/// Result returned by a custom inline parser when it recognizes a token at the current input position.
/// </summary>
public readonly struct MarkdownInlineParseResult {
    /// <summary>
    /// Creates an inline parse result.
    /// </summary>
    /// <param name="inline">Inline node produced by the extension.</param>
    /// <param name="consumedLength">Number of source characters consumed by the extension.</param>
    public MarkdownInlineParseResult(IMarkdownInline inline, int consumedLength) {
        if (inline == null) {
            throw new ArgumentNullException(nameof(inline));
        }

        if (consumedLength <= 0) {
            throw new ArgumentOutOfRangeException(nameof(consumedLength), consumedLength, "Consumed length must be greater than zero.");
        }

        Inline = inline;
        ConsumedLength = consumedLength;
    }

    /// <summary>Inline node produced by the extension.</summary>
    public IMarkdownInline Inline { get; }

    /// <summary>Number of source characters consumed by the extension.</summary>
    public int ConsumedLength { get; }
}

/// <summary>
/// Context passed to custom inline parser extensions.
/// </summary>
public sealed class MarkdownInlineParserContext {
    private readonly Func<int, int, bool, bool, InlineSequence> _parseNested;
    private readonly MarkdownInlineSourceMap? _sourceMap;

    internal MarkdownInlineParserContext(
        string text,
        int position,
        MarkdownReaderOptions options,
        MarkdownReaderState? state,
        bool allowLinks,
        bool allowImages,
        MarkdownInlineSourceMap? sourceMap,
        Func<int, int, bool, bool, InlineSequence> parseNested) {
        Text = text ?? string.Empty;
        Position = position;
        Options = options ?? throw new ArgumentNullException(nameof(options));
        State = state;
        AllowLinks = allowLinks;
        AllowImages = allowImages;
        _sourceMap = sourceMap;
        _parseNested = parseNested ?? throw new ArgumentNullException(nameof(parseNested));
    }

    /// <summary>Full inline source text being parsed.</summary>
    public string Text { get; }

    /// <summary>Current parser position within <see cref="Text"/>.</summary>
    public int Position { get; }

    /// <summary>Reader options active for the current parse.</summary>
    public MarkdownReaderOptions Options { get; }

    /// <summary>Reader state shared with the current parse, including reference definitions when available.</summary>
    public MarkdownReaderState? State { get; }

    /// <summary>Whether nested parsing at this point is allowed to produce links.</summary>
    public bool AllowLinks { get; }

    /// <summary>Whether nested parsing at this point is allowed to produce images.</summary>
    public bool AllowImages { get; }

    /// <summary>Character at the current parser position, or <c>'\0'</c> when positioned at the end of the input.</summary>
    public char CurrentChar => Position >= 0 && Position < Text.Length ? Text[Position] : '\0';

    /// <summary>Returns the remaining unparsed text starting at <see cref="Position"/>.</summary>
    public string RemainingText => Position >= 0 && Position < Text.Length ? Text.Substring(Position) : string.Empty;

    /// <summary>
    /// Returns the source span for a slice relative to the current parser position, when source mapping is available.
    /// </summary>
    public MarkdownSourceSpan? GetSourceSpan(int relativeStart, int length) {
        if (_sourceMap == null) {
            return null;
        }

        if (relativeStart < 0 || length <= 0) {
            return null;
        }

        return _sourceMap.GetSpan(Position + relativeStart, length);
    }

    /// <summary>
    /// Parses a nested inline segment relative to the current parser position while preserving source-map offsets.
    /// </summary>
    public InlineSequence ParseNestedInlines(int relativeStart, int length, bool allowLinks = true, bool allowImages = true) {
        if (relativeStart < 0 || length <= 0 || Position < 0 || Position >= Text.Length) {
            return new InlineSequence();
        }

        if (Position + relativeStart >= Text.Length) {
            return new InlineSequence();
        }

        var safeLength = Math.Min(length, Text.Length - (Position + relativeStart));
        if (safeLength <= 0) {
            return new InlineSequence();
        }

        return _parseNested(Position + relativeStart, safeLength, allowLinks, allowImages);
    }
}

/// <summary>
/// Delegate used by custom inline parser extensions.
/// Return <see langword="true"/> and a <see cref="MarkdownInlineParseResult"/> when the extension
/// recognizes an inline token at the current position; otherwise return <see langword="false"/>.
/// </summary>
public delegate bool MarkdownInlineParser(MarkdownInlineParserContext context, out MarkdownInlineParseResult result);

/// <summary>
/// Named custom inline parser registration used by <see cref="MarkdownReader"/>.
/// </summary>
public sealed class MarkdownInlineParserExtension {
    /// <summary>
    /// Creates an inline parser extension registration.
    /// </summary>
    public MarkdownInlineParserExtension(
        string name,
        MarkdownInlineParser parser,
        Func<MarkdownReaderOptions, bool>? isEnabled = null) {
        if (string.IsNullOrWhiteSpace(name)) {
            throw new ArgumentException("Extension name cannot be null or whitespace.", nameof(name));
        }

        Name = name.Trim();
        Parser = parser ?? throw new ArgumentNullException(nameof(parser));
        IsEnabled = isEnabled;
    }

    /// <summary>Stable extension name used for diagnostics and inspection.</summary>
    public string Name { get; }

    /// <summary>Parser delegate contributed by this extension.</summary>
    public MarkdownInlineParser Parser { get; }

    /// <summary>Optional predicate that decides whether the extension should apply for a specific options instance.</summary>
    public Func<MarkdownReaderOptions, bool>? IsEnabled { get; }

    /// <summary>Returns <see langword="true"/> when the extension should participate in parsing.</summary>
    public bool AppliesTo(MarkdownReaderOptions options) => IsEnabled?.Invoke(options) ?? true;
}
