using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

/// <summary>How much of one source item a salvage run brought back.</summary>
public enum EmailStoreSalvageStatus {
    /// <summary>The item's source data was read completely and verified where the format allows it.</summary>
    Recovered = 0,
    /// <summary>Some of the item was written, but part of its source data was damaged, missing, or unverified.</summary>
    Partial = 1,
    /// <summary>Nothing usable was written for the item.</summary>
    NotRecovered = 2
}

/// <summary>Outcome and source provenance for one item found in a damaged store.</summary>
public sealed class EmailStoreSalvageEntry {
    internal EmailStoreSalvageEntry(string itemId, EmailStoreSalvageStatus status, string sourceLocation,
        long? sourceOffset, long recoveredBytes, long? declaredBytes, bool? isInActiveIndex,
        long? damageOffset, string? damageCode, OutlookItemKind? kind, string? subject,
        string? destinationPath, long bytesWritten, IReadOnlyList<EmailStoreDiagnostic> diagnostics) {
        ItemId = itemId;
        Status = status;
        SourceLocation = sourceLocation;
        SourceOffset = sourceOffset;
        RecoveredBytes = recoveredBytes;
        DeclaredBytes = declaredBytes;
        IsInActiveIndex = isInActiveIndex;
        DamageOffset = damageOffset;
        DamageCode = damageCode;
        Kind = kind;
        Subject = subject;
        DestinationPath = destinationPath;
        BytesWritten = bytesWritten;
        Diagnostics = diagnostics;
    }

    /// <summary>
    /// Stable source identity: <c>dbx:item:&lt;record&gt;</c> and <c>dbx:chain:&lt;fragment&gt;</c> for DBX,
    /// the store item id for OLM records, and <c>olm:entry:&lt;path&gt;</c> for an unreadable OLM XML entry.
    /// </summary>
    public string ItemId { get; }
    /// <summary>Whether the item was recovered fully, partially, or not at all.</summary>
    public EmailStoreSalvageStatus Status { get; }
    /// <summary>Human-readable source structure, such as a DBX record address or an OLM entry path and index.</summary>
    public string SourceLocation { get; }
    /// <summary>Byte offset of the item's source structure in the damaged file, when known.</summary>
    public long? SourceOffset { get; }
    /// <summary>Source bytes recovered for the item: MIME bytes for DBX, decoded property bytes for OLM records.</summary>
    public long RecoveredBytes { get; }
    /// <summary>Size the source declared for the item, when it declared one. DBX sizes are native hints.</summary>
    public long? DeclaredBytes { get; }
    /// <summary>Whether a readable active index listed the item; null when the index could not be read.</summary>
    public bool? IsInActiveIndex { get; }
    /// <summary>Byte offset of the first damaged structure that stopped or degraded recovery.</summary>
    public long? DamageOffset { get; }
    /// <summary>Diagnostic code describing the damage at <see cref="DamageOffset"/>.</summary>
    public string? DamageCode { get; }
    /// <summary>Outlook item kind when it is known.</summary>
    public OutlookItemKind? Kind { get; }
    /// <summary>Subject or display title, when one could be read.</summary>
    public string? Subject { get; }
    /// <summary>Absolute artifact path, or null when nothing was written.</summary>
    public string? DestinationPath { get; }
    /// <summary>Artifact length.</summary>
    public long BytesWritten { get; }
    /// <summary>Damage, conversion, and write diagnostics for this item.</summary>
    public IReadOnlyList<EmailStoreDiagnostic> Diagnostics { get; }
}
