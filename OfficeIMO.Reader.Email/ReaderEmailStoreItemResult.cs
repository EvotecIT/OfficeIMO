using OfficeIMO.Email;
using OfficeIMO.Email.Store;

namespace OfficeIMO.Reader.Email;

/// <summary>Reader projection of one selected email-store item.</summary>
public sealed class ReaderEmailStoreItemResult {
    internal ReaderEmailStoreItemResult(
        EmailStoreItemReference reference,
        EmailStoreItemSummary? summary,
        string logicalPath,
        IReadOnlyList<ReaderChunk> chunks,
        IReadOnlyList<EmailDiagnostic> itemDiagnostics,
        IReadOnlyList<EmailDiagnostic>? storeDiagnostics = null) {
        Reference = reference;
        Summary = summary;
        LogicalPath = logicalPath;
        Chunks = chunks;
        ItemDiagnostics = itemDiagnostics ?? throw new ArgumentNullException(nameof(itemDiagnostics));
        StoreDiagnostics = storeDiagnostics ?? Array.Empty<EmailDiagnostic>();
        Diagnostics = StoreDiagnostics.Concat(ItemDiagnostics).ToArray();
    }

    /// <summary>Stable reference within the source store.</summary>
    public EmailStoreItemReference Reference { get; }
    /// <summary>Lightweight summary when it could be read.</summary>
    public EmailStoreItemSummary? Summary { get; }
    /// <summary>Escaped logical store/folder/item path used by projected chunks.</summary>
    public string LogicalPath { get; }
    /// <summary>Chunks for this item only.</summary>
    public IReadOnlyList<ReaderChunk> Chunks { get; }
    /// <summary>Item-scoped diagnostics emitted while reading and projecting this item.</summary>
    public IReadOnlyList<EmailDiagnostic> ItemDiagnostics { get; }
    /// <summary>
    /// Store-open and hierarchy diagnostics. These are attached to the first streamed result only so callers see
    /// them without receiving the same diagnostics for every item.
    /// </summary>
    public IReadOnlyList<EmailDiagnostic> StoreDiagnostics { get; }
    /// <summary>Combined store and item diagnostics for compatibility and consolidated reporting.</summary>
    public IReadOnlyList<EmailDiagnostic> Diagnostics { get; }
    /// <summary>Whether this item produced Reader chunks without an item-scoped error diagnostic.</summary>
    public bool Succeeded => Chunks.Count > 0 &&
        !ItemDiagnostics.Any(diagnostic => diagnostic.Severity == EmailDiagnosticSeverity.Error);
}
