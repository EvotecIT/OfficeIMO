namespace OfficeIMO.Markdown;

internal static class MarkdownOriginalSourceSliceMapper {
    public static bool TryCreateMapping(
        string originalMarkdown,
        string sourceMarkdown,
        bool preservesOriginalMarkdown,
        MarkdownSourceSpan span,
        out MarkdownSourceMapping mapping) {
        if (!MarkdownSourceSlice.TryCreate(sourceMarkdown, span, MarkdownSourceTextKind.Normalized, out var normalizedSlice)) {
            mapping = default;
            return false;
        }

        if (!preservesOriginalMarkdown) {
            mapping = CreateOriginalUnavailableMapping(
                span,
                normalizedSlice,
                MarkdownOriginalSourceSliceFailureReason.OriginalMarkdownNotPreserved);
            return true;
        }

        if (string.Equals(originalMarkdown, sourceMarkdown, StringComparison.Ordinal)) {
            if (MarkdownSourceSlice.TryCreate(originalMarkdown, span, MarkdownSourceTextKind.Original, out var originalSlice)) {
                mapping = new MarkdownSourceMapping(
                    span,
                    normalizedSlice,
                    hasOriginalSource: true,
                    originalSlice,
                    MarkdownOriginalSourceMappingKind.Exact,
                    MarkdownOriginalSourceSliceFailureReason.None);
                return true;
            }

            mapping = CreateOriginalUnavailableMapping(
                span,
                normalizedSlice,
                MarkdownOriginalSourceSliceFailureReason.OriginalSpanUnavailable);
            return true;
        }

        if (!LineEndingsAreEquivalent(originalMarkdown, sourceMarkdown)) {
            mapping = CreateOriginalUnavailableMapping(
                span,
                normalizedSlice,
                MarkdownOriginalSourceSliceFailureReason.OriginalTextNotEquivalent);
            return true;
        }

        if (TryCreateOriginalSourceSliceFromEquivalentLineEndings(originalMarkdown, sourceMarkdown, span, out var lineEndingMappedSlice)
            || MarkdownSourceSlice.TryCreateFromLineColumns(originalMarkdown, span, MarkdownSourceTextKind.Original, out lineEndingMappedSlice)) {
            mapping = new MarkdownSourceMapping(
                span,
                normalizedSlice,
                hasOriginalSource: true,
                lineEndingMappedSlice,
                MarkdownOriginalSourceMappingKind.LineEndingEquivalent,
                MarkdownOriginalSourceSliceFailureReason.None);
            return true;
        }

        mapping = CreateOriginalUnavailableMapping(
            span,
            normalizedSlice,
            MarkdownOriginalSourceSliceFailureReason.OriginalSpanUnavailable);
        return true;
    }

    public static bool TryCreate(
        string originalMarkdown,
        string sourceMarkdown,
        bool preservesOriginalMarkdown,
        MarkdownSourceSpan span,
        out MarkdownSourceSlice slice,
        out MarkdownOriginalSourceSliceFailureReason failureReason) {
        if (!TryCreateMapping(originalMarkdown, sourceMarkdown, preservesOriginalMarkdown, span, out var mapping)) {
            slice = default;
            failureReason = MarkdownOriginalSourceSliceFailureReason.OriginalSpanUnavailable;
            return false;
        }

        failureReason = mapping.OriginalSourceFailureReason;
        if (mapping.HasOriginalSource) {
            slice = mapping.OriginalSourceSlice;
            failureReason = MarkdownOriginalSourceSliceFailureReason.None;
            return true;
        }

        slice = default;
        return false;
    }

    private static MarkdownSourceMapping CreateOriginalUnavailableMapping(
        MarkdownSourceSpan span,
        MarkdownSourceSlice normalizedSlice,
        MarkdownOriginalSourceSliceFailureReason failureReason) =>
        new MarkdownSourceMapping(
            span,
            normalizedSlice,
            hasOriginalSource: false,
            originalSourceSlice: default,
            MarkdownOriginalSourceMappingKind.Unavailable,
            failureReason);

    private static bool LineEndingsAreEquivalent(string originalMarkdown, string sourceMarkdown) =>
        string.Equals(NormalizeLineEndings(originalMarkdown), sourceMarkdown, StringComparison.Ordinal);

    private static string NormalizeLineEndings(string value) =>
        value.Replace("\r\n", "\n").Replace('\r', '\n');

    private static bool TryCreateOriginalSourceSliceFromEquivalentLineEndings(
        string originalMarkdown,
        string sourceMarkdown,
        MarkdownSourceSpan span,
        out MarkdownSourceSlice slice) {
        if (TryMapEquivalentLineEndingOffsets(originalMarkdown, sourceMarkdown, span, out var originalStartOffset, out var originalEndOffsetInclusive)) {
            return MarkdownSourceSlice.TryCreateFromOffsets(
                originalMarkdown,
                span,
                MarkdownSourceTextKind.Original,
                originalStartOffset,
                originalEndOffsetInclusive,
                out slice);
        }

        slice = default;
        return false;
    }

    private static bool TryMapEquivalentLineEndingOffsets(
        string originalMarkdown,
        string sourceMarkdown,
        MarkdownSourceSpan span,
        out int originalStartOffset,
        out int originalEndOffsetInclusive) {
        originalStartOffset = 0;
        originalEndOffsetInclusive = -1;
        if (!span.StartOffset.HasValue || !span.EndOffset.HasValue) {
            return false;
        }

        var sourceStartOffset = span.StartOffset.Value;
        var sourceEndOffsetInclusive = span.EndOffset.Value;
        if (sourceStartOffset < 0
            || sourceEndOffsetInclusive < sourceStartOffset
            || sourceEndOffsetInclusive >= sourceMarkdown.Length) {
            return false;
        }

        var mappedStart = false;
        var mappedEnd = false;
        var originalIndex = 0;
        var sourceIndex = 0;
        while (sourceIndex < sourceMarkdown.Length && originalIndex < originalMarkdown.Length) {
            var sourceTokenStart = sourceIndex;
            var originalTokenStart = originalIndex;
            int sourceTokenLength;
            int originalTokenLength;
            var sourceLineBreak = IsLineBreakStart(sourceMarkdown, sourceIndex, out sourceTokenLength);
            var originalLineBreak = IsLineBreakStart(originalMarkdown, originalIndex, out originalTokenLength);
            if (sourceLineBreak != originalLineBreak) {
                return false;
            }

            if (!sourceLineBreak) {
                if (sourceMarkdown[sourceIndex] != originalMarkdown[originalIndex]) {
                    return false;
                }

                sourceTokenLength = 1;
                originalTokenLength = 1;
            }

            var sourceTokenEnd = sourceTokenStart + sourceTokenLength - 1;
            var originalTokenEnd = originalTokenStart + originalTokenLength - 1;
            if (!mappedStart && sourceStartOffset >= sourceTokenStart && sourceStartOffset <= sourceTokenEnd) {
                originalStartOffset = sourceLineBreak
                    ? originalTokenStart
                    : originalTokenStart + sourceStartOffset - sourceTokenStart;
                mappedStart = true;
            }

            if (!mappedEnd && sourceEndOffsetInclusive >= sourceTokenStart && sourceEndOffsetInclusive <= sourceTokenEnd) {
                originalEndOffsetInclusive = sourceLineBreak
                    ? originalTokenEnd
                    : originalTokenStart + sourceEndOffsetInclusive - sourceTokenStart;
                mappedEnd = true;
            }

            if (mappedStart && mappedEnd) {
                return true;
            }

            sourceIndex += sourceTokenLength;
            originalIndex += originalTokenLength;
        }

        return false;
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
