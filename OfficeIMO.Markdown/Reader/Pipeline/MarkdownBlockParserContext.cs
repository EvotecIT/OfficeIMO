namespace OfficeIMO.Markdown;

/// <summary>
/// Result returned by a custom block parser when it recognizes a block at the current line.
/// </summary>
public readonly struct MarkdownBlockParseResult {
    /// <summary>
    /// Creates a block parse result for a single block.
    /// </summary>
    public MarkdownBlockParseResult(IMarkdownBlock block, int consumedLineCount)
        : this(new[] { block }, consumedLineCount) {
    }

    /// <summary>
    /// Creates a block parse result for one or more blocks.
    /// </summary>
    public MarkdownBlockParseResult(IReadOnlyList<IMarkdownBlock> blocks, int consumedLineCount) {
        if (blocks == null) {
            throw new ArgumentNullException(nameof(blocks));
        }

        if (blocks.Count == 0) {
            throw new ArgumentException("At least one block must be returned.", nameof(blocks));
        }

        for (int i = 0; i < blocks.Count; i++) {
            if (blocks[i] == null) {
                throw new ArgumentException("Returned blocks cannot contain null entries.", nameof(blocks));
            }
        }

        if (consumedLineCount <= 0) {
            throw new ArgumentOutOfRangeException(nameof(consumedLineCount), consumedLineCount, "Consumed line count must be greater than zero.");
        }

        Blocks = blocks;
        ConsumedLineCount = consumedLineCount;
    }

    /// <summary>
    /// Blocks produced by the parser.
    /// </summary>
    public IReadOnlyList<IMarkdownBlock> Blocks { get; }

    /// <summary>
    /// Number of source lines consumed by the parser.
    /// </summary>
    public int ConsumedLineCount { get; }
}

/// <summary>
/// Context passed to delegate-based custom block parser extensions.
/// </summary>
public sealed class MarkdownBlockParserContext {
    private readonly string[] _lines;

    internal MarkdownBlockParserContext(
        string[] lines,
        int lineIndex,
        MarkdownReaderOptions options,
        MarkdownDoc document,
        MarkdownReaderState state) {
        _lines = lines ?? Array.Empty<string>();
        LineIndex = lineIndex;
        Options = options ?? throw new ArgumentNullException(nameof(options));
        Document = document ?? throw new ArgumentNullException(nameof(document));
        State = state ?? throw new ArgumentNullException(nameof(state));
    }

    /// <summary>
    /// Full document lines being parsed.
    /// </summary>
    public IReadOnlyList<string> Lines => _lines;

    /// <summary>
    /// Zero-based current line index.
    /// </summary>
    public int LineIndex { get; }

    /// <summary>
    /// One-based current line number.
    /// </summary>
    public int LineNumber => LineIndex + 1;

    /// <summary>
    /// Current source line, or an empty string when positioned outside the input.
    /// </summary>
    public string CurrentLine => LineIndex >= 0 && LineIndex < _lines.Length ? _lines[LineIndex] : string.Empty;

    /// <summary>
    /// Reader options active for the current parse.
    /// </summary>
    public MarkdownReaderOptions Options { get; }

    /// <summary>
    /// Markdown document being built.
    /// </summary>
    public MarkdownDoc Document { get; }

    /// <summary>
    /// Mutable reader state shared across parsers.
    /// </summary>
    public MarkdownReaderState State { get; }

    /// <summary>
    /// Returns a source line relative to the current line.
    /// </summary>
    public bool TryGetLine(int relativeOffset, out string line) {
        var index = LineIndex + relativeOffset;
        if (index >= 0 && index < _lines.Length) {
            line = _lines[index];
            return true;
        }

        line = string.Empty;
        return false;
    }

    /// <summary>
    /// Returns a slice of source lines relative to the current line.
    /// </summary>
    public IReadOnlyList<string> GetLines(int relativeStart, int lineCount) {
        if (lineCount <= 0) {
            return Array.Empty<string>();
        }

        var startIndex = LineIndex + relativeStart;
        if (startIndex < 0 || startIndex >= _lines.Length) {
            return Array.Empty<string>();
        }

        var safeCount = Math.Min(lineCount, _lines.Length - startIndex);
        var slice = new string[safeCount];
        Array.Copy(_lines, startIndex, slice, 0, safeCount);
        return slice;
    }

    /// <summary>
    /// Parses a nested markdown block range relative to the current line using the same reader options,
    /// while preserving source spans for the nested content.
    /// </summary>
    public IReadOnlyList<IMarkdownBlock> ParseNestedBlocks(int relativeStartLine, int lineCount) =>
        MarkdownReader.ParseNestedBlocksFromLineRange(_lines, LineIndex + relativeStartLine, lineCount, Options, State);
}

/// <summary>
/// Delegate used by custom block parser extensions.
/// Return <see langword="true"/> and a <see cref="MarkdownBlockParseResult"/> when the extension
/// recognizes a block starting at the current line; otherwise return <see langword="false"/>.
/// </summary>
public delegate bool MarkdownBlockParser(MarkdownBlockParserContext context, out MarkdownBlockParseResult result);

internal sealed class DelegateMarkdownBlockParser : IMarkdownBlockParser {
    private readonly MarkdownBlockParser _parser;

    public DelegateMarkdownBlockParser(MarkdownBlockParser parser) {
        _parser = parser ?? throw new ArgumentNullException(nameof(parser));
    }

    public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
        var context = new MarkdownBlockParserContext(lines, i, options, doc, state);
        if (!_parser(context, out var result)) {
            return false;
        }

        if (i + result.ConsumedLineCount > lines.Length) {
            throw new ArgumentOutOfRangeException(nameof(result), result.ConsumedLineCount, "Consumed line count exceeds the remaining input.");
        }

        for (int blockIndex = 0; blockIndex < result.Blocks.Count; blockIndex++) {
            doc.Add(result.Blocks[blockIndex]);
        }

        i += result.ConsumedLineCount;
        return true;
    }
}
