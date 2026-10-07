namespace OfficeIMO.Pdf;

internal static partial class PdfDocumentSemanticEnricher {
    private readonly record struct TaggedListContent(int ObjectNumber, int Level, bool IsLabel);

    private static PdfTaggedListItemSource[] BuildTaggedListSources(
        PdfReadPage readPage,
        PdfUnderstandingPageResult page,
        TaggedContentRoleIndex? roles,
        int maximum,
        PdfUnderstandingWorkBudget budget) {
        if (roles is null) return Array.Empty<PdfTaggedListItemSource>();
        var items = new Dictionary<int, PdfTaggedListItemSource>();
        foreach (PdfTextSpan run in page.DecodedRuns) {
            budget.Consume();
            if (!run.MarkedContentId.HasValue) continue;
            TaggedListContent? membership = roles.GetListContent(page.PageNumber, readPage,
                new MarkedContentKey(run.ContentStreamObjectNumber, run.MarkedContentId.Value));
            if (!membership.HasValue) continue;
            TaggedListContent content = membership.Value;
            if (!items.TryGetValue(content.ObjectNumber, out PdfTaggedListItemSource? item)) {
                if (items.Count >= maximum) throw PdfReadLimitException.Create(PdfReadLimitKind.UnderstandingArtifacts, maximum, (long)items.Count + 1);
                item = new PdfTaggedListItemSource(content.ObjectNumber, content.Level);
                items.Add(content.ObjectNumber, item);
            }
            (content.IsLabel ? item.LabelRuns : item.BodyRuns).Add(run);
        }
        return items.Values.ToArray();
    }

    private sealed partial class TaggedContentRoleIndex {
        private readonly Dictionary<int, Dictionary<MarkedContentKey, TaggedListContent?>> _listContentByPage = new();

        private void AddListContent(int pageNumber, MarkedContentKey key, TaggedListContent? content) {
            if (!_listContentByPage.TryGetValue(pageNumber, out Dictionary<MarkedContentKey, TaggedListContent?>? page)) {
                page = new Dictionary<MarkedContentKey, TaggedListContent?>();
                _listContentByPage.Add(pageNumber, page);
            }
            // Conflicting structure references cannot establish a unique list owner.
            if (page.TryGetValue(key, out TaggedListContent? existing)) {
                if (existing != content) page[key] = null;
            } else {
                page.Add(key, content);
            }
        }

        internal TaggedListContent? GetListContent(int pageNumber, PdfReadPage readPage, MarkedContentKey key) {
            if (!_listContentByPage.TryGetValue(pageNumber, out Dictionary<MarkedContentKey, TaggedListContent?>? page)) return null;
            if (page.TryGetValue(key, out TaggedListContent? exact)) return exact;
            if (!key.ContentStreamObjectNumber.HasValue) {
                MarkedContentKey? scoped = ResolveUniquePageContentKey(pageNumber, readPage, key.MarkedContentId);
                return scoped.HasValue && page.TryGetValue(scoped.Value, out TaggedListContent? found) ? found : null;
            }
            return readPage.IsPageContentStreamObjectNumber(key.ContentStreamObjectNumber) &&
                page.TryGetValue(new MarkedContentKey(null, key.MarkedContentId), out TaggedListContent? unscoped)
                ? unscoped : null;
        }
    }
}

/// <summary>Nearest reachable tagged LI owner and its page-local label/body source runs.</summary>
internal sealed class PdfTaggedListItemSource {
    internal PdfTaggedListItemSource(int objectNumber, int level) { ObjectNumber = objectNumber; Level = level; }
    internal int ObjectNumber { get; }
    internal int Level { get; }
    internal HashSet<PdfTextSpan> LabelRuns { get; } = new();
    internal HashSet<PdfTextSpan> BodyRuns { get; } = new();
}
