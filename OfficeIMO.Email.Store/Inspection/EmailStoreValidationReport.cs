namespace OfficeIMO.Email.Store;

/// <summary>Result of bounded validation over an open store.</summary>
public sealed class EmailStoreValidationReport {
    internal EmailStoreValidationReport(EmailStoreValidationMode mode, int itemsExamined,
        int itemsFailed, int orphanedItems, bool wasTruncated,
        IReadOnlyList<EmailStoreDiagnostic> diagnostics) {
        Mode = mode;
        ItemsExamined = itemsExamined;
        ItemsFailed = itemsFailed;
        OrphanedItems = orphanedItems;
        WasTruncated = wasTruncated;
        Diagnostics = diagnostics;
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

    /// <summary>Whether validation covered its complete selected scope.</summary>
    public bool IsComplete => !WasTruncated && ItemsFailed == 0;

    /// <summary>Whether no error-severity diagnostics were observed.</summary>
    public bool IsValid => !Diagnostics.Any(item => item.Severity == EmailStoreDiagnosticSeverity.Error);
}
