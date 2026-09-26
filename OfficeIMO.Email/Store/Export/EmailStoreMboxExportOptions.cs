using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

/// <summary>Acceptance policy applied before a staged mbox artifact is committed.</summary>
public enum EmailStoreMboxExportFidelityPolicy {
    /// <summary>Commits a usable partial mailbox and preserves every omission or approximation in the report.</summary>
    AllowDiagnosedLoss = 0,
    /// <summary>Rejects publication when typed diagnostics report approximation, omission, or failure.</summary>
    RequireNoLoss = 1
}

/// <summary>Selection and streaming policy for mbox export.</summary>
public sealed class EmailStoreMboxExportOptions {
    /// <summary>Creates mbox export options.</summary>
    public EmailStoreMboxExportOptions(
        string? folderId = null,
        bool includeDescendants = false,
        bool includeAssociatedItems = false,
        bool includeOrphanedItems = false,
        bool overwriteExisting = false,
        bool continueOnError = true,
        int maxItems = 100_000,
        EmailMailboxWriterOptions? writerOptions = null) {
        if (maxItems <= 0) throw new ArgumentOutOfRangeException(nameof(maxItems));
        FolderId = string.IsNullOrWhiteSpace(folderId) ? null : folderId;
        IncludeDescendants = includeDescendants;
        IncludeAssociatedItems = includeAssociatedItems;
        IncludeOrphanedItems = includeOrphanedItems;
        OverwriteExisting = overwriteExisting;
        ContinueOnError = continueOnError;
        MaxItems = maxItems;
        WriterOptions = writerOptions ?? EmailMailboxWriterOptions.Default;
        FidelityPolicy = EmailStoreMboxExportFidelityPolicy.AllowDiagnosedLoss;
    }

    /// <summary>Creates mbox export options with an explicit pre-commit fidelity policy.</summary>
    public EmailStoreMboxExportOptions(
        EmailStoreMboxExportFidelityPolicy fidelityPolicy,
        string? folderId = null,
        bool includeDescendants = false,
        bool includeAssociatedItems = false,
        bool includeOrphanedItems = false,
        bool overwriteExisting = false,
        bool continueOnError = true,
        int maxItems = 100_000,
        EmailMailboxWriterOptions? writerOptions = null)
        : this(folderId, includeDescendants, includeAssociatedItems, includeOrphanedItems,
            overwriteExisting, continueOnError, maxItems, writerOptions) {
        if (!Enum.IsDefined(typeof(EmailStoreMboxExportFidelityPolicy), fidelityPolicy)) {
            throw new ArgumentOutOfRangeException(nameof(fidelityPolicy));
        }
        FidelityPolicy = fidelityPolicy;
    }

    /// <summary>Optional source folder identifier. Null exports every folder.</summary>
    public string? FolderId { get; }

    /// <summary>Whether descendants of <see cref="FolderId"/> are included.</summary>
    public bool IncludeDescendants { get; }

    /// <summary>Whether folder-associated information items are exported.</summary>
    public bool IncludeAssociatedItems { get; }

    /// <summary>Whether recoverable items absent from contents tables are exported.</summary>
    public bool IncludeOrphanedItems { get; }

    /// <summary>Whether an existing destination may be atomically replaced.</summary>
    public bool OverwriteExisting { get; }

    /// <summary>Whether later items are attempted after a read or conversion failure.</summary>
    public bool ContinueOnError { get; }

    /// <summary>Maximum item references attempted.</summary>
    public int MaxItems { get; }

    /// <summary>Underlying mbox variant and OfficeIMO.Email message writer policy.</summary>
    public EmailMailboxWriterOptions WriterOptions { get; }

    /// <summary>Acceptance policy checked against typed fidelity diagnostics before the staged mailbox is committed.</summary>
    public EmailStoreMboxExportFidelityPolicy FidelityPolicy { get; }
}
