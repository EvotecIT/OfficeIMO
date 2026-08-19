namespace OfficeIMO.Email.Store;

/// <summary>Privacy-safe identity of the exact source accepted by a resumable Store migration.</summary>
public sealed class EmailStoreSourceIdentity {
    internal EmailStoreSourceIdentity(EmailStoreFormat format, long length,
        string catalogFingerprint, string durableFingerprint) {
        Format = format;
        Length = length;
        CatalogFingerprint = catalogFingerprint;
        DurableFingerprint = durableFingerprint;
    }

    /// <summary>Detected source format.</summary>
    public EmailStoreFormat Format { get; }
    /// <summary>Validated aggregate source length.</summary>
    public long Length { get; }
    /// <summary>SHA-256 fingerprint of the bounded source catalog.</summary>
    public string CatalogFingerprint { get; }
    /// <summary>SHA-256 fingerprint of all persisted source bytes.</summary>
    public string DurableFingerprint { get; }
}

/// <summary>Controls incomplete migration artifacts when an operation stops before commit.</summary>
public enum EmailStorePartialResultPolicy {
    /// <summary>Retain the integrity-checked checkpoint and writer-owned working files for resume.</summary>
    RetainResumableState = 0,
    /// <summary>Delete the checkpoint and writer-owned working files.</summary>
    DiscardIncompleteState = 1
}

/// <summary>Final strict-loss disposition of a Store migration.</summary>
public enum EmailStoreMigrationDisposition {
    /// <summary>Every selected item was written and any requested verification succeeded.</summary>
    Completed = 0,
    /// <summary>The caller allowed one or more selected items to be skipped with diagnostics.</summary>
    CompletedWithAcceptedLoss = 1
}
