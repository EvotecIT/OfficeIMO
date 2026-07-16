namespace OfficeIMO.Email.Store;

/// <summary>Outcome for one item in a directory export.</summary>
public sealed class EmailStoreExportEntry {
    internal EmailStoreExportEntry(EmailStoreItemReference reference, string? destinationPath,
        long bytesWritten, IReadOnlyList<EmailStoreDiagnostic> diagnostics) {
        Reference = reference;
        DestinationPath = destinationPath;
        BytesWritten = bytesWritten;
        Diagnostics = diagnostics;
    }

    /// <summary>Stable source item reference.</summary>
    public EmailStoreItemReference Reference { get; }

    /// <summary>Absolute destination path, or null when no artifact was created.</summary>
    public string? DestinationPath { get; }

    /// <summary>Serialized artifact length.</summary>
    public long BytesWritten { get; }

    /// <summary>Read, conversion, and write diagnostics for this item.</summary>
    public IReadOnlyList<EmailStoreDiagnostic> Diagnostics { get; }

    /// <summary>Whether an artifact was written without an error diagnostic.</summary>
    public bool Succeeded => DestinationPath != null &&
        !Diagnostics.Any(item => item.Severity == EmailStoreDiagnosticSeverity.Error);
}
