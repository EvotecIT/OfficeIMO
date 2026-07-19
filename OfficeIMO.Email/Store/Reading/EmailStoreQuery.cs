using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

/// <summary>Bounded query evaluated over lightweight summaries from an open store session.</summary>
public sealed class EmailStoreQuery {
    /// <summary>Creates a bounded store query.</summary>
    public EmailStoreQuery(
        string? folderId = null,
        bool includeDescendants = false,
        bool includeAssociatedItems = false,
        bool includeOrphanedItems = false,
        OutlookItemKind? itemKind = null,
        string? subjectContains = null,
        string? senderContains = null,
        DateTimeOffset? since = null,
        DateTimeOffset? before = null,
        bool? hasAttachments = null,
        bool? isRead = null,
        int maxItemsScanned = 1_000_000,
        int maxResults = 1_000) {
        if (maxItemsScanned <= 0) throw new ArgumentOutOfRangeException(nameof(maxItemsScanned));
        if (maxResults <= 0) throw new ArgumentOutOfRangeException(nameof(maxResults));
        if (since.HasValue && before.HasValue && since.Value > before.Value) {
            throw new ArgumentException("The query start must not be later than its end.", nameof(since));
        }
        FolderId = string.IsNullOrWhiteSpace(folderId) ? null : folderId;
        IncludeDescendants = includeDescendants;
        IncludeAssociatedItems = includeAssociatedItems;
        IncludeOrphanedItems = includeOrphanedItems;
        ItemKind = itemKind;
        SubjectContains = EmptyToNull(subjectContains);
        SenderContains = EmptyToNull(senderContains);
        Since = since;
        Before = before;
        HasAttachments = hasAttachments;
        IsRead = isRead;
        MaxItemsScanned = maxItemsScanned;
        MaxResults = maxResults;
    }

    /// <summary>Optional folder identifier. Null searches every folder.</summary>
    public string? FolderId { get; }

    /// <summary>Whether descendants of <see cref="FolderId"/> are included.</summary>
    public bool IncludeDescendants { get; }

    /// <summary>Whether folder-associated information items are searched.</summary>
    public bool IncludeAssociatedItems { get; }

    /// <summary>Whether items recovered outside folder contents tables are searched.</summary>
    public bool IncludeOrphanedItems { get; }

    /// <summary>Optional typed Outlook item filter.</summary>
    public OutlookItemKind? ItemKind { get; }

    /// <summary>Optional case-insensitive subject fragment.</summary>
    public string? SubjectContains { get; }

    /// <summary>Optional case-insensitive sender name or address fragment.</summary>
    public string? SenderContains { get; }

    /// <summary>Optional inclusive lower bound applied to received time, then sent time.</summary>
    public DateTimeOffset? Since { get; }

    /// <summary>Optional exclusive upper bound applied to received time, then sent time.</summary>
    public DateTimeOffset? Before { get; }

    /// <summary>Optional declared attachment-presence filter.</summary>
    public bool? HasAttachments { get; }

    /// <summary>Optional read-state filter.</summary>
    public bool? IsRead { get; }

    /// <summary>Maximum references whose summaries may be evaluated.</summary>
    public int MaxItemsScanned { get; }

    /// <summary>Maximum matching results returned.</summary>
    public int MaxResults { get; }

    private static string? EmptyToNull(string? value) =>
        string.IsNullOrWhiteSpace(value) ? null : value;
}
