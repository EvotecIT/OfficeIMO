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
    /// Creates a source span for a line range relative to the current parser position.
    /// </summary>
    /// <param name="relativeStartLine">Zero-based line offset from the current parser position.</param>
    /// <param name="lineCount">Number of source lines covered by the span.</param>
    /// <returns>A source span mapped to the normalized markdown input when source mapping is available.</returns>
    public MarkdownSourceSpan CreateLineSpan(int relativeStartLine, int lineCount) {
        if (lineCount <= 0) {
            throw new ArgumentOutOfRangeException(nameof(lineCount), lineCount, "Line count must be greater than zero.");
        }

        var startLine = ResolveAbsoluteLine(relativeStartLine);
        var endLine = ResolveAbsoluteLine(relativeStartLine + lineCount - 1);
        return State.SourceTextMap?.CreateLineSpan(startLine, endLine) ?? new MarkdownSourceSpan(startLine, endLine);
    }

    /// <summary>
    /// Creates a column-aware source span for a range relative to the current parser position.
    /// </summary>
    /// <param name="relativeStartLine">Zero-based start-line offset from the current parser position.</param>
    /// <param name="startColumn">One-based start column on the resolved start line.</param>
    /// <param name="relativeEndLine">Zero-based end-line offset from the current parser position.</param>
    /// <param name="endColumn">One-based end column on the resolved end line.</param>
    /// <returns>A source span mapped to the normalized markdown input when source mapping is available.</returns>
    public MarkdownSourceSpan CreateSourceSpan(
        int relativeStartLine,
        int startColumn,
        int relativeEndLine,
        int endColumn) {
        var startLine = ResolveAbsoluteLine(relativeStartLine);
        var endLine = ResolveAbsoluteLine(relativeEndLine);
        return State.SourceTextMap?.CreateSpan(startLine, startColumn, endLine, endColumn)
               ?? new MarkdownSourceSpan(startLine, startColumn, endLine, endColumn);
    }

    /// <summary>
    /// Creates a normalized source slice for a token or field source span captured during parsing.
    /// </summary>
    public bool TryCreateSourceSlice(MarkdownSourceSpan sourceSpan, out MarkdownSourceSlice slice) {
        var sourceMap = State.SourceTextMap;
        if (sourceMap == null) {
            slice = default;
            return false;
        }

        return MarkdownSourceSlice.TryCreate(sourceMap.Text, sourceSpan, MarkdownSourceTextKind.Normalized, out slice);
    }

    /// <summary>
    /// Creates a normalized source slice for a line range relative to the current parser position.
    /// </summary>
    /// <param name="relativeStartLine">Zero-based line offset from the current parser position.</param>
    /// <param name="lineCount">Number of source lines covered by the slice.</param>
    /// <param name="slice">Materialized normalized source slice when the method returns <c>true</c>.</param>
    /// <returns><c>true</c> when the normalized source text is available and the range can be materialized.</returns>
    public bool TryCreateSourceSlice(int relativeStartLine, int lineCount, out MarkdownSourceSlice slice) {
        if (lineCount <= 0) {
            slice = default;
            return false;
        }

        return TryCreateSourceSlice(CreateLineSpan(relativeStartLine, lineCount), out slice);
    }

    /// <summary>
    /// Parses an inline slice from a source line while preserving source spans for inline syntax.
    /// </summary>
    /// <param name="relativeLine">Zero-based line offset from the current parser position.</param>
    /// <param name="startColumn">One-based start column for the inline slice.</param>
    /// <param name="length">Number of characters to parse from the source line.</param>
    /// <returns>Parsed inline sequence with source-map-backed inline nodes when available.</returns>
    public InlineSequence ParseInlineText(int relativeLine, int startColumn, int length) {
        if (length <= 0 || startColumn < 1 || !TryGetLine(relativeLine, out var line)) {
            return new InlineSequence();
        }

        var startIndex = Math.Min(line.Length, startColumn - 1);
        var safeLength = Math.Min(length, line.Length - startIndex);
        if (safeLength <= 0) {
            return new InlineSequence();
        }

        var text = line.Substring(startIndex, safeLength);
        MarkdownInlineSourceMap? sourceMap = null;
        if (State.SourceTextMap != null) {
            var absoluteLine = ResolveAbsoluteLine(relativeLine);
            var points = new MarkdownSourcePoint?[text.Length];
            var sourceColumn = startColumn;
            for (var i = 0; i < text.Length; i++) {
                points[i] = State.SourceTextMap.CreatePoint(absoluteLine, sourceColumn);
                sourceColumn = MarkdownSourceColumns.AdvanceColumn(sourceColumn, text[i]);
            }

            sourceMap = new MarkdownInlineSourceMap(points);
        }

        return MarkdownReader.ParseInlineText(text, Options, State, sourceMap);
    }

    /// <summary>
    /// Parses a nested markdown block range relative to the current line using the same reader options,
    /// while preserving source spans for the nested content.
    /// </summary>
    public IReadOnlyList<IMarkdownBlock> ParseNestedBlocks(int relativeStartLine, int lineCount) =>
        MarkdownReader.ParseNestedBlocksFromLineRange(_lines, LineIndex + relativeStartLine, lineCount, Options, State);

    private int ResolveAbsoluteLine(int relativeLine) {
        var localLineIndex = LineIndex + relativeLine;
        var absoluteLines = State.SourceLineAbsoluteNumbers;
        if (absoluteLines != null
            && localLineIndex >= 0
            && localLineIndex < absoluteLines.Count
            && absoluteLines[localLineIndex] > 0) {
            return absoluteLines[localLineIndex];
        }

        return State.SourceLineOffset + localLineIndex + 1;
    }
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
