namespace OfficeIMO.Markdown;

internal static class MarkdownTransformSourceSpanHelper {
    internal readonly record struct ChangedBlockRange(
        int StartBefore,
        int CountBefore,
        int StartAfter,
        int CountAfter,
        int SuffixCount);

    internal static string[] CreateBlockFingerprints(IReadOnlyList<IMarkdownBlock> blocks) {
        var fingerprints = new string[blocks.Count];
        for (var i = 0; i < blocks.Count; i++) {
            var block = blocks[i];
            fingerprints[i] = block == null
                ? string.Empty
                : (block.GetType().FullName ?? block.GetType().Name) + "\n" + block.RenderMarkdown();
        }

        return fingerprints;
    }

    internal static ChangedBlockRange ComputeChangedRange(
        string[] before,
        string[] after) {
        var prefix = 0;
        while (prefix < before.Length
               && prefix < after.Length
               && string.Equals(before[prefix], after[prefix], StringComparison.Ordinal)) {
            prefix++;
        }

        var suffix = 0;
        while (suffix < before.Length - prefix
               && suffix < after.Length - prefix
               && string.Equals(before[before.Length - 1 - suffix], after[after.Length - 1 - suffix], StringComparison.Ordinal)) {
            suffix++;
        }

        return new ChangedBlockRange(
            StartBefore: prefix,
            CountBefore: before.Length - prefix - suffix,
            StartAfter: prefix,
            CountAfter: after.Length - prefix - suffix,
            SuffixCount: suffix);
    }

    internal static MarkdownSourceSpan? AggregateSpans(
        IReadOnlyList<MarkdownSourceSpan?>? spans,
        int start,
        int count) {
        if (spans == null || count <= 0 || start < 0 || start >= spans.Count) {
            return null;
        }

        MarkdownSourceSpan? first = null;
        MarkdownSourceSpan? last = null;
        var endExclusive = Math.Min(spans.Count, start + count);
        for (var i = start; i < endExclusive; i++) {
            if (!spans[i].HasValue) {
                continue;
            }

            var span = spans[i]!.Value;
            if (!first.HasValue || CompareStart(span, first.Value) < 0) {
                first = span;
            }

            if (!last.HasValue || CompareEnd(span, last.Value) > 0) {
                last = span;
            }
        }

        return first.HasValue && last.HasValue
            ? CreateAggregateSpan(first.Value, last.Value)
            : null;
    }

    internal static MarkdownSourceSpan? AggregateBlockSpans(
        IReadOnlyList<IMarkdownBlock> blocks,
        int start,
        int count) {
        if (count <= 0 || start < 0 || start >= blocks.Count) {
            return null;
        }

        MarkdownSourceSpan? first = null;
        MarkdownSourceSpan? last = null;
        var endExclusive = Math.Min(blocks.Count, start + count);
        for (var i = start; i < endExclusive; i++) {
            if (blocks[i] is not MarkdownObject blockObject || !blockObject.SourceSpan.HasValue) {
                continue;
            }

            var span = blockObject.SourceSpan.Value;
            if (!first.HasValue || CompareStart(span, first.Value) < 0) {
                first = span;
            }

            if (!last.HasValue || CompareEnd(span, last.Value) > 0) {
                last = span;
            }
        }

        return first.HasValue && last.HasValue
            ? CreateAggregateSpan(first.Value, last.Value)
            : null;
    }

    internal static List<MarkdownSourceSpan?> UpdateBlockSpans(
        IReadOnlyList<MarkdownSourceSpan?>? previousSpans,
        int newCount,
        ChangedBlockRange change,
        MarkdownSourceSpan? affectedSourceSpan) {
        var updated = new List<MarkdownSourceSpan?>(newCount);
        if (previousSpans == null || previousSpans.Count == 0) {
            for (var i = 0; i < newCount; i++) {
                updated.Add(null);
            }

            return updated;
        }

        var prefixCount = Math.Min(change.StartBefore, previousSpans.Count);
        for (var i = 0; i < prefixCount; i++) {
            updated.Add(previousSpans[i]);
        }

        for (var i = 0; i < change.CountAfter; i++) {
            updated.Add(affectedSourceSpan);
        }

        var suffixStart = Math.Max(prefixCount, previousSpans.Count - change.SuffixCount);
        for (var i = suffixStart; i < previousSpans.Count; i++) {
            updated.Add(previousSpans[i]);
        }

        while (updated.Count < newCount) {
            updated.Add(null);
        }

        if (updated.Count > newCount) {
            updated.RemoveRange(newCount, updated.Count - newCount);
        }

        return updated;
    }

    internal static void ApplyAffectedSpanToChangedBlocks(
        IReadOnlyList<IMarkdownBlock> blocks,
        ChangedBlockRange change,
        MarkdownSourceSpan? affectedSourceSpan,
        bool overwriteExistingSpans = false) {
        if (blocks == null || blocks.Count == 0 || !affectedSourceSpan.HasValue || change.CountAfter <= 0) {
            return;
        }

        var endExclusive = Math.Min(blocks.Count, change.StartAfter + change.CountAfter);
        for (var i = change.StartAfter; i < endExclusive; i++) {
            if (blocks[i] is not MarkdownObject blockObject) {
                continue;
            }

            if (overwriteExistingSpans || !blockObject.SourceSpan.HasValue) {
                blockObject.SourceSpan = affectedSourceSpan;
            }
        }
    }

    private static MarkdownSourceSpan CreateAggregateSpan(MarkdownSourceSpan first, MarkdownSourceSpan last) {
        if (first.StartColumn.HasValue && last.EndColumn.HasValue) {
            int? startOffset = first.StartOffset.HasValue && last.EndOffset.HasValue
                ? first.StartOffset
                : null;
            int? endOffset = first.StartOffset.HasValue && last.EndOffset.HasValue
                ? last.EndOffset
                : null;
            return new MarkdownSourceSpan(
                first.StartLine,
                first.StartColumn.Value,
                last.EndLine,
                last.EndColumn.Value,
                startOffset,
                endOffset);
        }

        return new MarkdownSourceSpan(first.StartLine, last.EndLine);
    }

    private static int CompareStart(MarkdownSourceSpan left, MarkdownSourceSpan right) {
        int lineCompare = left.StartLine.CompareTo(right.StartLine);
        if (lineCompare != 0) {
            return lineCompare;
        }

        return NormalizeStartColumn(left.StartColumn).CompareTo(NormalizeStartColumn(right.StartColumn));
    }

    private static int CompareEnd(MarkdownSourceSpan left, MarkdownSourceSpan right) {
        int lineCompare = left.EndLine.CompareTo(right.EndLine);
        if (lineCompare != 0) {
            return lineCompare;
        }

        return NormalizeEndColumn(left.EndColumn).CompareTo(NormalizeEndColumn(right.EndColumn));
    }

    private static int NormalizeStartColumn(int? column) => column ?? 1;

    private static int NormalizeEndColumn(int? column) => column ?? int.MaxValue;
}
