namespace OfficeIMO.Email.Store;

/// <summary>
/// Bounded Store table query with a typed filter AST, deterministic ordering, projection, and keyset continuation.
/// </summary>
public sealed class EmailStoreTableQuery {
    private readonly IReadOnlyList<EmailStoreSort> _sorts;
    private readonly IReadOnlyList<EmailStoreSort> _effectiveSorts;

    /// <summary>Creates an immutable table query.</summary>
    public EmailStoreTableQuery(
        EmailStoreFolderId? folderId = null,
        bool includeDescendants = false,
        bool includeAssociatedItems = false,
        bool includeOrphanedItems = false,
        EmailStoreFilter? filter = null,
        IEnumerable<EmailStoreSort>? sorts = null,
        EmailStoreProjection? projection = null,
        EmailStoreContinuationToken? continuationToken = null,
        int maxItemsScanned = 1_000_000,
        int pageSize = 100) {
        if (maxItemsScanned <= 0) throw new ArgumentOutOfRangeException(nameof(maxItemsScanned));
        if (pageSize <= 0) throw new ArgumentOutOfRangeException(nameof(pageSize));
        if (folderId.HasValue && folderId.Value.IsEmpty) throw new ArgumentException("The folder identifier cannot be the default value.", nameof(folderId));
        FolderId = folderId;
        IncludeDescendants = includeDescendants;
        IncludeAssociatedItems = includeAssociatedItems;
        IncludeOrphanedItems = includeOrphanedItems;
        Filter = filter ?? EmailStoreFilter.All;
        EmailStoreSort[] requested = sorts?.ToArray() ?? Array.Empty<EmailStoreSort>();
        if (requested.Any(sort => sort == null)) throw new ArgumentException("Sorts cannot contain null.", nameof(sorts));
        string? duplicate = requested.GroupBy(sort => sort.Field.Key, StringComparer.Ordinal)
            .FirstOrDefault(group => group.Count() > 1)?.Key;
        if (duplicate != null) throw new ArgumentException(string.Concat("Sort field '", duplicate, "' is duplicated."), nameof(sorts));
        if (requested.Length == 0) requested = new[] { EmailStoreFields.ReceivedAt.Descending() };
        _sorts = Array.AsReadOnly(requested);

        var effective = requested.ToList();
        if (!effective.Any(sort => sort.Field.Key == EmailStoreFields.FolderId.Key)) {
            effective.Add(EmailStoreFields.FolderId.Ascending());
        }
        if (!effective.Any(sort => sort.Field.Key == EmailStoreFields.ItemId.Key)) {
            effective.Add(EmailStoreFields.ItemId.Ascending());
        }
        _effectiveSorts = effective.AsReadOnly();
        Projection = projection ?? EmailStoreProjection.Default;
        ContinuationToken = continuationToken;
        MaxItemsScanned = maxItemsScanned;
        PageSize = pageSize;
    }

    /// <summary>Optional typed folder scope.</summary>
    public EmailStoreFolderId? FolderId { get; }

    /// <summary>Whether descendants of <see cref="FolderId"/> are included.</summary>
    public bool IncludeDescendants { get; }

    /// <summary>Whether folder-associated information is included.</summary>
    public bool IncludeAssociatedItems { get; }

    /// <summary>Whether source-index orphans are included.</summary>
    public bool IncludeOrphanedItems { get; }

    /// <summary>Typed expression tree evaluated over lightweight references and summaries.</summary>
    public EmailStoreFilter Filter { get; }

    /// <summary>Caller-requested ordering. Stable folder/item tie-breakers are appended internally.</summary>
    public IReadOnlyList<EmailStoreSort> Sorts => _sorts;

    /// <summary>Selected output columns.</summary>
    public EmailStoreProjection Projection { get; }

    /// <summary>Optional keyset continuation from a compatible earlier page.</summary>
    public EmailStoreContinuationToken? ContinuationToken { get; }

    /// <summary>Maximum lightweight references evaluated per execution.</summary>
    public int MaxItemsScanned { get; }

    /// <summary>Maximum rows returned in one page.</summary>
    public int PageSize { get; }

    /// <summary>Returns a copy that resumes from the supplied continuation.</summary>
    public EmailStoreTableQuery ContinueFrom(EmailStoreContinuationToken? token) => Copy(continuationToken: token, replaceToken: true);

    /// <summary>Returns a copy with a different typed filter.</summary>
    public EmailStoreTableQuery Where(EmailStoreFilter filter) => Copy(filter: filter ?? throw new ArgumentNullException(nameof(filter)));

    /// <summary>Returns a copy with a different ordered column list.</summary>
    public EmailStoreTableQuery OrderBy(params EmailStoreSort[] sorts) => Copy(sorts: sorts ?? throw new ArgumentNullException(nameof(sorts)));

    /// <summary>Returns a copy with a different table projection.</summary>
    public EmailStoreTableQuery Select(params EmailStoreField[] fields) => Copy(projection: new EmailStoreProjection(fields));

    /// <summary>Returns the deterministic execution plan without reading item payloads.</summary>
    public EmailStoreQueryPlan Explain() => new EmailStoreQueryPlan(
        Filter,
        _effectiveSorts,
        Projection,
        MaxItemsScanned,
        PageSize,
        materializesMatchesForSort: true,
        readsItemPayloads: false);

    internal IReadOnlyList<EmailStoreSort> EffectiveSorts => _effectiveSorts;

    internal string Signature => string.Concat(
        "scope=", EmailStoreScalarCodec.Signature(FolderId),
        ";desc=", IncludeDescendants ? "1" : "0",
        ";fai=", IncludeAssociatedItems ? "1" : "0",
        ";orph=", IncludeOrphanedItems ? "1" : "0",
        ";max=", MaxItemsScanned.ToString(CultureInfo.InvariantCulture),
        ";filter=", Filter.Signature,
        ";sort=", string.Join(",", _effectiveSorts.Select(sort => sort.Signature)));

    private EmailStoreTableQuery Copy(
        EmailStoreFilter? filter = null,
        IEnumerable<EmailStoreSort>? sorts = null,
        EmailStoreProjection? projection = null,
        EmailStoreContinuationToken? continuationToken = null,
        bool replaceToken = false) =>
        new EmailStoreTableQuery(
            FolderId,
            IncludeDescendants,
            IncludeAssociatedItems,
            IncludeOrphanedItems,
            filter ?? Filter,
            sorts ?? _sorts,
            projection ?? Projection,
            replaceToken ? continuationToken : ContinuationToken,
            MaxItemsScanned,
            PageSize);
}

/// <summary>Read behavior and deterministic ordering selected for a Store table query.</summary>
public sealed class EmailStoreQueryPlan {
    internal EmailStoreQueryPlan(EmailStoreFilter filter, IReadOnlyList<EmailStoreSort> effectiveSorts,
        EmailStoreProjection projection, int maxItemsScanned, int pageSize,
        bool materializesMatchesForSort, bool readsItemPayloads) {
        Filter = filter;
        EffectiveSorts = effectiveSorts;
        Projection = projection;
        MaxItemsScanned = maxItemsScanned;
        PageSize = pageSize;
        MaterializesMatchesForSort = materializesMatchesForSort;
        ReadsItemPayloads = readsItemPayloads;
    }

    /// <summary>Typed filter AST.</summary>
    public EmailStoreFilter Filter { get; }

    /// <summary>Effective ordering including stable tie-break columns.</summary>
    public IReadOnlyList<EmailStoreSort> EffectiveSorts { get; }

    /// <summary>Output column projection.</summary>
    public EmailStoreProjection Projection { get; }

    /// <summary>Maximum references evaluated.</summary>
    public int MaxItemsScanned { get; }

    /// <summary>Maximum returned rows.</summary>
    public int PageSize { get; }

    /// <summary>Whether matching summary rows are buffered to establish global ordering.</summary>
    public bool MaterializesMatchesForSort { get; }

    /// <summary>Whether complete documents, bodies, recipients, or attachment payloads are read.</summary>
    public bool ReadsItemPayloads { get; }
}
