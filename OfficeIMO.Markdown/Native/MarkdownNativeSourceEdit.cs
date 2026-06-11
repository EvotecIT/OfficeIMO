namespace OfficeIMO.Markdown;

/// <summary>
/// Non-mutating source edit derived from a native markdown source span.
/// </summary>
public sealed class MarkdownNativeSourceEdit {
    internal MarkdownNativeSourceEdit(MarkdownSourceSpan sourceSpan, int startOffset, int endOffsetInclusive, string replacementMarkdown) {
        SourceSpan = sourceSpan;
        StartOffset = startOffset;
        EndOffsetInclusive = endOffsetInclusive;
        ReplacementMarkdown = replacementMarkdown ?? string.Empty;
    }

    /// <summary>Source span this edit replaces.</summary>
    public MarkdownSourceSpan SourceSpan { get; }

    /// <summary>0-based inclusive start offset in the source markdown.</summary>
    public int StartOffset { get; }

    /// <summary>0-based inclusive end offset in the source markdown.</summary>
    public int EndOffsetInclusive { get; }

    /// <summary>Replacement markdown.</summary>
    public string ReplacementMarkdown { get; }

    /// <summary>Applies this edit to the supplied markdown source and returns the edited text.</summary>
    public string Apply(string sourceMarkdown) {
        sourceMarkdown ??= string.Empty;
        if (StartOffset < 0 || StartOffset > sourceMarkdown.Length) {
            throw new InvalidOperationException("Edit start offset is outside the supplied markdown source.");
        }

        var endExclusive = Math.Min(sourceMarkdown.Length, EndOffsetInclusive + 1);
        if (endExclusive < StartOffset) {
            throw new InvalidOperationException("Edit end offset is before the start offset.");
        }

        return sourceMarkdown.Substring(0, StartOffset)
               + ReplacementMarkdown
               + sourceMarkdown.Substring(endExclusive);
    }
}
