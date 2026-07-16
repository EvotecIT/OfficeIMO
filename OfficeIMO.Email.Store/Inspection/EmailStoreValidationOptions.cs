namespace OfficeIMO.Email.Store;

/// <summary>Scope and safety bounds for email-store validation.</summary>
public sealed class EmailStoreValidationOptions {
    /// <summary>Creates validation options.</summary>
    public EmailStoreValidationOptions(
        EmailStoreValidationMode mode = EmailStoreValidationMode.Summaries,
        string? folderId = null,
        bool includeDescendants = false,
        bool includeAssociatedItems = false,
        bool includeOrphanedItems = true,
        int maxItems = 100_000) {
        if (!Enum.IsDefined(typeof(EmailStoreValidationMode), mode)) {
            throw new ArgumentOutOfRangeException(nameof(mode));
        }
        if (maxItems <= 0) throw new ArgumentOutOfRangeException(nameof(maxItems));
        Mode = mode;
        FolderId = string.IsNullOrWhiteSpace(folderId) ? null : folderId;
        IncludeDescendants = includeDescendants;
        IncludeAssociatedItems = includeAssociatedItems;
        IncludeOrphanedItems = includeOrphanedItems;
        MaxItems = maxItems;
    }

    /// <summary>Validation depth.</summary>
    public EmailStoreValidationMode Mode { get; }

    /// <summary>Optional folder identifier. Null validates every folder.</summary>
    public string? FolderId { get; }

    /// <summary>Whether descendants of <see cref="FolderId"/> are included.</summary>
    public bool IncludeDescendants { get; }

    /// <summary>Whether folder-associated information items are validated.</summary>
    public bool IncludeAssociatedItems { get; }

    /// <summary>Whether recoverable items absent from contents tables are validated.</summary>
    public bool IncludeOrphanedItems { get; }

    /// <summary>Maximum item references examined.</summary>
    public int MaxItems { get; }
}
