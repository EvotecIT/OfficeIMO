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
        int maxItems = 100_000,
        bool verifyStructuralIntegrity = false,
        int maxStructuralPages = 100_000,
        int maxStructuralBlocks = 100_000,
        long maxStructuralBytes = 1024L * 1024 * 1024) {
        if (!Enum.IsDefined(typeof(EmailStoreValidationMode), mode)) {
            throw new ArgumentOutOfRangeException(nameof(mode));
        }
        if (maxItems <= 0) throw new ArgumentOutOfRangeException(nameof(maxItems));
        if (maxStructuralPages <= 0) throw new ArgumentOutOfRangeException(nameof(maxStructuralPages));
        if (maxStructuralBlocks <= 0) throw new ArgumentOutOfRangeException(nameof(maxStructuralBlocks));
        if (maxStructuralBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maxStructuralBytes));
        Mode = mode;
        FolderId = string.IsNullOrWhiteSpace(folderId) ? null : folderId;
        IncludeDescendants = includeDescendants;
        IncludeAssociatedItems = includeAssociatedItems;
        IncludeOrphanedItems = includeOrphanedItems;
        MaxItems = maxItems;
        VerifyStructuralIntegrity = verifyStructuralIntegrity;
        MaxStructuralPages = maxStructuralPages;
        MaxStructuralBlocks = maxStructuralBlocks;
        MaxStructuralBytes = maxStructuralBytes;
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

    /// <summary>
    /// Whether PST/OST page and block trailers, CRCs, signatures, identifiers, lengths, and bounds are verified.
    /// This is opt-in because it reads each selected structure rather than only the data needed for browsing.
    /// </summary>
    public bool VerifyStructuralIntegrity { get; }

    /// <summary>Maximum BBT and NBT pages examined by structural validation.</summary>
    public int MaxStructuralPages { get; }

    /// <summary>Maximum BBT-referenced blocks examined by structural validation.</summary>
    public int MaxStructuralBlocks { get; }

    /// <summary>Maximum page, payload, and trailer bytes read by structural validation.</summary>
    public long MaxStructuralBytes { get; }
}
