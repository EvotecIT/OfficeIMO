namespace OfficeIMO.Email.Store;

/// <summary>Outcome for one item in a directory export.</summary>
public sealed class EmailStoreExportEntry {
    internal EmailStoreExportEntry(EmailStoreItemReference reference, string? destinationPath,
        long bytesWritten, IReadOnlyList<EmailStoreDiagnostic> diagnostics,
        string? maildirFlags = null) {
        Reference = reference;
        DestinationPath = destinationPath;
        BytesWritten = bytesWritten;
        Diagnostics = diagnostics;
        MaildirFlags = maildirFlags;
    }

    /// <summary>Stable source item reference.</summary>
    public EmailStoreItemReference Reference { get; }

    /// <summary>Absolute destination path, or null when no artifact was created.</summary>
    public string? DestinationPath { get; }

    /// <summary>Serialized artifact length.</summary>
    public long BytesWritten { get; }

    /// <summary>Read, conversion, and write diagnostics for this item.</summary>
    public IReadOnlyList<EmailStoreDiagnostic> Diagnostics { get; }

    /// <summary>
    /// Maildir information flags in canonical D,F,P,R,S,T order when the export format is Maildir.
    /// This remains populated when the destination file system cannot encode the <c>:2,</c> suffix.
    /// </summary>
    public string? MaildirFlags { get; }

    /// <summary>Whether an artifact was written without an error diagnostic.</summary>
    public bool Succeeded => DestinationPath != null &&
        !Diagnostics.Any(item => item.Severity == EmailStoreDiagnosticSeverity.Error);
}
