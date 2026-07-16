namespace OfficeIMO.Email.Store;

/// <summary>Outcome for one source item appended to an mbox export.</summary>
public sealed class EmailStoreMboxExportEntry {
    internal EmailStoreMboxExportEntry(EmailStoreItemReference reference, long bytesWritten,
        IReadOnlyList<EmailStoreDiagnostic> diagnostics) {
        Reference = reference;
        BytesWritten = bytesWritten;
        Diagnostics = diagnostics;
    }

    /// <summary>Stable source item reference.</summary>
    public EmailStoreItemReference Reference { get; }

    /// <summary>Bytes appended for this entry.</summary>
    public long BytesWritten { get; }

    /// <summary>Read and conversion diagnostics for this entry.</summary>
    public IReadOnlyList<EmailStoreDiagnostic> Diagnostics { get; }

    /// <summary>Whether an mbox entry was appended without an error diagnostic.</summary>
    public bool Succeeded => BytesWritten > 0 &&
        !Diagnostics.Any(item => item.Severity == EmailStoreDiagnosticSeverity.Error);
}
