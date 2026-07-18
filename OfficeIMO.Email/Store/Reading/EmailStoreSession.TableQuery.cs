namespace OfficeIMO.Email.Store;

public sealed partial class EmailStoreSession {
    /// <summary>
    /// Executes a bounded typed table query using lightweight summaries only. Complete item payloads are not read.
    /// </summary>
    public EmailStoreTablePage SearchPage(EmailStoreTableQuery query,
        CancellationToken cancellationToken = default) {
        if (query == null) throw new ArgumentNullException(nameof(query));
        ThrowIfDisposed();
        if (query.FolderId.HasValue) FolderCatalog.Get(query.FolderId.Value);

        int enumerationLimit = query.MaxItemsScanned == int.MaxValue
            ? int.MaxValue
            : query.MaxItemsScanned + 1;
        var enumeration = new EmailStoreEnumerationOptions(
            query.FolderId?.Value,
            query.IncludeDescendants,
            query.IncludeAssociatedItems,
            query.IncludeOrphanedItems,
            enumerationLimit);
        var matches = new List<EmailStoreQueryRow>();
        int scanned = 0;
        bool scanLimitReached = false;
        foreach (EmailStoreItemReference reference in EnumerateItems(enumeration, cancellationToken)) {
            cancellationToken.ThrowIfCancellationRequested();
            if (scanned >= query.MaxItemsScanned) {
                scanLimitReached = true;
                break;
            }
            scanned++;
            var row = new EmailStoreQueryRow(reference, ReadSummary(reference, cancellationToken));
            if (query.Filter.Evaluate(row)) matches.Add(row);
        }

        matches.Sort(new QueryRowComparer(query.EffectiveSorts));
        IReadOnlyList<object?>? continuationValues = query.ContinuationToken?.DecodeValues(query);
        IEnumerable<EmailStoreQueryRow> remaining = continuationValues == null
            ? matches
            : matches.Where(row => CompareToContinuation(row, continuationValues, query.EffectiveSorts) > 0);
        int take = query.PageSize == int.MaxValue ? int.MaxValue : query.PageSize + 1;
        List<EmailStoreQueryRow> pageCandidates = remaining.Take(take).ToList();
        bool hasMore = pageCandidates.Count > query.PageSize;
        if (hasMore) pageCandidates.RemoveAt(pageCandidates.Count - 1);

        EmailStoreContinuationToken? nextToken = hasMore && pageCandidates.Count > 0
            ? EmailStoreContinuationToken.Create(query, pageCandidates[pageCandidates.Count - 1])
            : null;
        EmailStoreTableRow[] rows = pageCandidates
            .Select(row => new EmailStoreTableRow(row, query.Projection))
            .ToArray();
        return new EmailStoreTablePage(
            Array.AsReadOnly(rows),
            nextToken,
            scanned,
            matches.Count,
            scanLimitReached,
            query.Explain());
    }

    private static int CompareToContinuation(EmailStoreQueryRow row,
        IReadOnlyList<object?> values, IReadOnlyList<EmailStoreSort> sorts) {
        for (int index = 0; index < sorts.Count; index++) {
            int result = sorts[index].CompareValues(sorts[index].Field.Read(row), values[index]);
            if (result != 0) return result;
        }
        return 0;
    }

    private sealed class QueryRowComparer : IComparer<EmailStoreQueryRow> {
        private readonly IReadOnlyList<EmailStoreSort> _sorts;

        internal QueryRowComparer(IReadOnlyList<EmailStoreSort> sorts) {
            _sorts = sorts;
        }

        public int Compare(EmailStoreQueryRow? left, EmailStoreQueryRow? right) {
            if (ReferenceEquals(left, right)) return 0;
            if (left == null) return -1;
            if (right == null) return 1;
            foreach (EmailStoreSort sort in _sorts) {
                int result = sort.Compare(left, right);
                if (result != 0) return result;
            }
            return 0;
        }
    }
}
