namespace OfficeIMO.Markdown;

/// <summary>
/// Identifies which source text a <see cref="MarkdownSourceSlice"/> was materialized from.
/// </summary>
public enum MarkdownSourceTextKind {
    /// <summary>Normalized markdown text used to compute source spans.</summary>
    Normalized,
    /// <summary>Original reader input retained for future lossless roundtrip support.</summary>
    Original
}

/// <summary>
/// Materialized source text for a span-backed syntax node.
/// </summary>
public readonly struct MarkdownSourceSlice {
    private readonly string? _sourceText;

    private MarkdownSourceSlice(
        string sourceText,
        MarkdownSourceSpan sourceSpan,
        MarkdownSourceTextKind textKind,
        int startOffset,
        int endOffsetInclusive) {
        _sourceText = sourceText ?? string.Empty;
        SourceSpan = sourceSpan;
        TextKind = textKind;
        StartOffset = startOffset;
        EndOffsetInclusive = endOffsetInclusive;
    }

    /// <summary>Source span represented by this slice.</summary>
    public MarkdownSourceSpan SourceSpan { get; }

    /// <summary>Source text kind used to materialize this slice.</summary>
    public MarkdownSourceTextKind TextKind { get; }

    /// <summary>0-based inclusive start offset in the backing source text.</summary>
    public int StartOffset { get; }

    /// <summary>0-based inclusive end offset in the backing source text.</summary>
    public int EndOffsetInclusive { get; }

    /// <summary>Exact source text covered by this slice.</summary>
    public string Text {
        get {
            var sourceText = _sourceText ?? string.Empty;
            if (sourceText.Length == 0 || EndOffsetInclusive < StartOffset) {
                return string.Empty;
            }

            var endExclusive = Math.Min(sourceText.Length, EndOffsetInclusive + 1);
            return sourceText.Substring(StartOffset, endExclusive - StartOffset);
        }
    }

    /// <summary>Attempts to create a source slice from the supplied source text and span.</summary>
    public static bool TryCreate(
        string? sourceText,
        MarkdownSourceSpan sourceSpan,
        MarkdownSourceTextKind textKind,
        out MarkdownSourceSlice slice) {
        sourceText ??= string.Empty;
        if (sourceText.Length == 0) {
            slice = default;
            return false;
        }

        if (TryResolveOffsets(sourceText, sourceSpan, out var startOffset, out var endOffsetInclusive)) {
            slice = new MarkdownSourceSlice(sourceText, sourceSpan, textKind, startOffset, endOffsetInclusive);
            return true;
        }

        slice = default;
        return false;
    }

    /// <summary>
    /// Attempts to create a source slice using only line and column coordinates, ignoring any normalized-text offsets.
    /// </summary>
    public static bool TryCreateFromLineColumns(
        string? sourceText,
        MarkdownSourceSpan sourceSpan,
        MarkdownSourceTextKind textKind,
        out MarkdownSourceSlice slice) {
        sourceText ??= string.Empty;
        if (sourceText.Length == 0) {
            slice = default;
            return false;
        }

        if (TryResolveLineColumnOffsets(sourceText, sourceSpan, out var startOffset, out var endOffsetInclusive)) {
            slice = new MarkdownSourceSlice(sourceText, sourceSpan, textKind, startOffset, endOffsetInclusive);
            return true;
        }

        slice = default;
        return false;
    }

    private static bool TryResolveOffsets(
        string sourceText,
        MarkdownSourceSpan span,
        out int startOffset,
        out int endOffsetInclusive) {
        if (span.StartOffset.HasValue && span.EndOffset.HasValue) {
            startOffset = ClampOffset(sourceText, span.StartOffset.Value);
            endOffsetInclusive = ClampOffset(sourceText, span.EndOffset.Value);
            return endOffsetInclusive >= startOffset;
        }

        return TryResolveLineColumnOffsets(sourceText, span, out startOffset, out endOffsetInclusive);
    }

    private static bool TryResolveLineColumnOffsets(
        string sourceText,
        MarkdownSourceSpan span,
        out int startOffset,
        out int endOffsetInclusive) {
        if (span.StartColumn.HasValue
            && span.EndColumn.HasValue
            && TryGetOffset(sourceText, span.StartLine, span.StartColumn.Value, out startOffset)
            && TryGetOffset(sourceText, span.EndLine, span.EndColumn.Value, out endOffsetInclusive)) {
            return endOffsetInclusive >= startOffset;
        }

        if (TryGetLineStartOffset(sourceText, span.StartLine, out startOffset)
            && TryGetLineEndOffset(sourceText, span.EndLine, out endOffsetInclusive)) {
            return endOffsetInclusive >= startOffset;
        }

        startOffset = 0;
        endOffsetInclusive = -1;
        return false;
    }

    private static int ClampOffset(string sourceText, int offset) {
        if (offset < 0) {
            return 0;
        }

        return offset >= sourceText.Length ? sourceText.Length - 1 : offset;
    }

    private static bool TryGetOffset(string sourceText, int lineNumber, int columnNumber, out int offset) {
        if (!TryGetLineStartOffset(sourceText, lineNumber, out var lineStart)) {
            offset = 0;
            return false;
        }

        offset = Math.Min(sourceText.Length - 1, lineStart + Math.Max(0, columnNumber - 1));
        return true;
    }

    private static bool TryGetLineStartOffset(string sourceText, int lineNumber, out int offset) {
        offset = 0;
        if (lineNumber < 1) {
            return false;
        }

        if (lineNumber == 1) {
            return true;
        }

        var currentLine = 1;
        for (var i = 0; i < sourceText.Length; i++) {
            if (!IsLineBreakStart(sourceText, i, out var lineBreakLength)) {
                continue;
            }

            currentLine++;
            if (currentLine == lineNumber) {
                offset = i + lineBreakLength;
                return offset <= sourceText.Length;
            }

            i += lineBreakLength - 1;
        }

        return false;
    }

    private static bool TryGetLineEndOffset(string sourceText, int lineNumber, out int offset) {
        if (!TryGetLineStartOffset(sourceText, lineNumber, out var lineStart)) {
            offset = 0;
            return false;
        }

        offset = sourceText.Length - 1;
        for (var i = lineStart; i < sourceText.Length; i++) {
            if (IsLineBreakStart(sourceText, i, out _)) {
                offset = Math.Max(lineStart, i - 1);
                return true;
            }
        }

        return true;
    }

    private static bool IsLineBreakStart(string sourceText, int offset, out int length) {
        if (sourceText[offset] == '\r') {
            length = offset + 1 < sourceText.Length && sourceText[offset + 1] == '\n'
                ? 2
                : 1;
            return true;
        }

        if (sourceText[offset] == '\n') {
            length = 1;
            return true;
        }

        length = 0;
        return false;
    }
}
