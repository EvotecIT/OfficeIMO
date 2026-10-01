namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkTextReader {
    /// <summary>Shares the paragraph decoder and bounded inheritance cache across table roles and selected cells.</summary>
    internal sealed class TableStyleResolver(IWorkObjectIndex index, IWorkProjectionBudget budget,
        IWorkSourceReferenceIssueCollector references) {
        private readonly Dictionary<ulong, Cached<IWorkParagraphStyle>> _cache = new();

        internal IWorkParagraphStyle? Read(ulong identifier, ref bool complete) {
            bool resolved = true;
            IWorkParagraphStyle style = ResolveParagraphStyle(index, identifier, budget, _cache,
                tolerateStyleDepth: true, references, ref resolved);
            if (!resolved) { complete = false; return null; }
            return style;
        }
    }
}
