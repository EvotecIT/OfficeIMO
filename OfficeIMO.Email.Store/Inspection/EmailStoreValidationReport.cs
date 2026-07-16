namespace OfficeIMO.Email.Store;

/// <summary>Result of bounded validation over an open store.</summary>
public sealed class EmailStoreValidationReport {
    internal EmailStoreValidationReport(EmailStoreValidationMode mode, int itemsExamined,
        int itemsFailed, int orphanedItems, bool wasTruncated,
        IReadOnlyList<EmailStoreDiagnostic> diagnostics,
        bool structuralIntegrityRequested = false,
        EmailStoreStructuralValidationResult? structural = null) {
        Mode = mode;
        ItemsExamined = itemsExamined;
        ItemsFailed = itemsFailed;
        OrphanedItems = orphanedItems;
        WasTruncated = wasTruncated;
        Diagnostics = diagnostics;
        StructuralIntegrityRequested = structuralIntegrityRequested;
        StructuralIntegritySupported = structural?.Supported == true;
        StructuralPagesExamined = structural?.PagesExamined ?? 0;
        StructuralBlocksExamined = structural?.BlocksExamined ?? 0;
        StructuralBytesExamined = structural?.BytesExamined ?? 0;
        StructuralFailures = structural?.Failures ?? 0;
        StructuralValidationWasTruncated = structural?.WasTruncated == true;
    }

    /// <summary>Validation depth that was executed.</summary>
    public EmailStoreValidationMode Mode { get; }

    /// <summary>Number of item references examined.</summary>
    public int ItemsExamined { get; }

    /// <summary>Number of selected items that could not be validated at the requested depth.</summary>
    public int ItemsFailed { get; }

    /// <summary>Number of examined items recovered outside folder contents tables.</summary>
    public int OrphanedItems { get; }

    /// <summary>Whether validation stopped at the configured item bound.</summary>
    public bool WasTruncated { get; }

    /// <summary>Opening, parsing, and per-item diagnostics observed by this validation.</summary>
    public IReadOnlyList<EmailStoreDiagnostic> Diagnostics { get; }

    /// <summary>Whether trailer-level structural integrity verification was requested.</summary>
    public bool StructuralIntegrityRequested { get; }

    /// <summary>Whether the source format supports the requested trailer-level validation.</summary>
    public bool StructuralIntegritySupported { get; }

    /// <summary>Number of BBT and NBT pages examined.</summary>
    public int StructuralPagesExamined { get; }

    /// <summary>Number of BBT-referenced blocks examined.</summary>
    public int StructuralBlocksExamined { get; }

    /// <summary>Number of page, block-payload, and trailer bytes examined.</summary>
    public long StructuralBytesExamined { get; }

    /// <summary>Number of pages or blocks that failed one or more structural checks.</summary>
    public int StructuralFailures { get; }

    /// <summary>Whether a structural page, block, or byte bound stopped verification early.</summary>
    public bool StructuralValidationWasTruncated { get; }

    /// <summary>Whether validation covered its complete selected scope.</summary>
    public bool IsComplete => !WasTruncated && ItemsFailed == 0 &&
        (!StructuralIntegrityRequested ||
            (StructuralIntegritySupported && !StructuralValidationWasTruncated && StructuralFailures == 0));

    /// <summary>Whether no error-severity diagnostics were observed.</summary>
    public bool IsValid => StructuralFailures == 0 &&
        !Diagnostics.Any(item => item.Severity == EmailStoreDiagnosticSeverity.Error);
}
