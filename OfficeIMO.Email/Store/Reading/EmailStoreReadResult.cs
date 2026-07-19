namespace OfficeIMO.Email.Store;

/// <summary>Result of a bounded email-store read.</summary>
public sealed class EmailStoreReadResult {
    internal EmailStoreReadResult(EmailStore store, IReadOnlyList<EmailStoreDiagnostic> diagnostics, long bytesRead) {
        Store = store;
        Diagnostics = diagnostics;
        BytesRead = bytesRead;
    }

    /// <summary>Materialized store.</summary>
    public EmailStore Store { get; }

    /// <summary>Structured diagnostics.</summary>
    public IReadOnlyList<EmailStoreDiagnostic> Diagnostics { get; }

    /// <summary>Source length validated by the reader.</summary>
    public long BytesRead { get; }

    /// <summary>True when an error diagnostic was emitted.</summary>
    public bool HasErrors => Diagnostics.Any(diagnostic => diagnostic.Severity == EmailStoreDiagnosticSeverity.Error);
}
