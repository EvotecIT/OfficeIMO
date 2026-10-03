namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkTextReader {
    private readonly struct ParagraphStyleCacheEntry(IWorkParagraphStyle value, bool isComplete, bool mappedPropertiesComplete) {
        internal IWorkParagraphStyle Value { get; } = value;
        internal bool IsComplete { get; } = isComplete;
        internal bool MappedPropertiesComplete { get; } = mappedPropertiesComplete;
    }

    /// <summary>Shares the paragraph decoder and bounded inheritance cache across table roles and selected cells.</summary>
    internal sealed class TableStyleResolver(IWorkObjectIndex index, IWorkProjectionBudget budget,
        IWorkSourceReferenceIssueCollector references) {
        private readonly Dictionary<ulong, ParagraphStyleCacheEntry> _cache = new();

        internal IWorkParagraphStyle? Read(ulong identifier, ref bool complete) {
            bool resolved = true;
            IWorkParagraphStyle style = ResolveParagraphStyle(index, identifier, budget, _cache,
                tolerateStyleDepth: true, references, ref resolved);
            if (!resolved) {
                complete = false;
                if (!_cache[identifier].MappedPropertiesComplete) return null;
            }
            return style;
        }
    }
}
