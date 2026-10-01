namespace OfficeIMO.Email.Store;

/// <summary>Summary matches and explicit completion evidence for a bounded metadata search.</summary>
public sealed class EmailStoreSearchReport {
    internal EmailStoreSearchReport(IReadOnlyList<EmailStoreSearchResult> results, int scanned, bool scanLimit, bool resultLimit) {
        Results = results; ItemsScanned = scanned; StoppedAtItemLimit = scanLimit; StoppedAtResultLimit = resultLimit;
    }
    /// <summary>Matches in enumeration order.</summary>
    public IReadOnlyList<EmailStoreSearchResult> Results { get; }
    /// <summary>References whose summaries were evaluated.</summary>
    public int ItemsScanned { get; }
    /// <summary>Whether more references existed beyond the scan budget.</summary>
    public bool StoppedAtItemLimit { get; }
    /// <summary>Whether more references existed beyond the result budget.</summary>
    public bool StoppedAtResultLimit { get; }
    /// <summary>Whether the selected enumeration was exhausted.</summary>
    public bool IsComplete => !StoppedAtItemLimit && !StoppedAtResultLimit;
}

public sealed partial class EmailStoreSession {
    /// <summary>Searches lightweight summaries and reports scan/result limits. Use content search for durable resumable batches.</summary>
    public EmailStoreSearchReport SearchWithReport(EmailStoreQuery query, CancellationToken cancellationToken = default) {
        if (query == null) throw new ArgumentNullException(nameof(query));
        ThrowIfDisposed();
        var enumeration = new EmailStoreEnumerationOptions(query.FolderId, query.IncludeDescendants,
            query.IncludeAssociatedItems, query.IncludeOrphanedItems,
            query.MaxItemsScanned == int.MaxValue ? int.MaxValue : query.MaxItemsScanned + 1);
        var results = new List<EmailStoreSearchResult>();
        int scanned = 0;
        bool scanLimit = false, resultLimit = false;
        using var references = EnumerateItems(enumeration, cancellationToken).GetEnumerator();
        while (references.MoveNext()) {
            cancellationToken.ThrowIfCancellationRequested();
            if (scanned >= query.MaxItemsScanned) { scanLimit = true; break; }
            var reference = references.Current;
            scanned++;
            var summary = ReadSummary(reference, cancellationToken);
            if (!Matches(query, summary)) continue;
            results.Add(new EmailStoreSearchResult(reference, summary));
            if (results.Count >= query.MaxResults) { resultLimit = references.MoveNext(); break; }
        }
        cancellationToken.ThrowIfCancellationRequested();
        return new EmailStoreSearchReport(results.AsReadOnly(), scanned, scanLimit, resultLimit);
    }
}
