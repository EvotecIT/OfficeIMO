using OfficeIMO.Email;
using OfficeIMO.Email.Store;

namespace OfficeIMO.Reader.EmailStore;

/// <summary>Reader projection of one selected email-store item.</summary>
public sealed class ReaderEmailStoreItemResult {
    internal ReaderEmailStoreItemResult(
        EmailStoreItemReference reference,
        EmailStoreItemSummary? summary,
        string logicalPath,
        IReadOnlyList<ReaderChunk> chunks,
        IReadOnlyList<EmailDiagnostic> diagnostics) {
        Reference = reference;
        Summary = summary;
        LogicalPath = logicalPath;
        Chunks = chunks;
        Diagnostics = diagnostics;
    }

    /// <summary>Stable reference within the source store.</summary>
    public EmailStoreItemReference Reference { get; }
    /// <summary>Lightweight summary when it could be read.</summary>
    public EmailStoreItemSummary? Summary { get; }
    /// <summary>Escaped logical store/folder/item path used by projected chunks.</summary>
    public string LogicalPath { get; }
    /// <summary>Chunks for this item only.</summary>
    public IReadOnlyList<ReaderChunk> Chunks { get; }
    /// <summary>Store and item diagnostics attributed to this item.</summary>
    public IReadOnlyList<EmailDiagnostic> Diagnostics { get; }
    /// <summary>Whether the item produced Reader chunks without an error diagnostic.</summary>
    public bool Succeeded => Chunks.Count > 0 &&
        !Diagnostics.Any(diagnostic => diagnostic.Severity == EmailDiagnosticSeverity.Error);
}
