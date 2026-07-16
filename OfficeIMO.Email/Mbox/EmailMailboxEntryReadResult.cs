namespace OfficeIMO.Email;

/// <summary>One streamed mbox entry together with message-scoped diagnostics and source byte count.</summary>
public sealed class EmailMailboxEntryReadResult {
    internal EmailMailboxEntryReadResult(EmailMailboxEntry entry, IReadOnlyList<EmailDiagnostic> diagnostics,
        long bytesRead) {
        Entry = entry;
        Diagnostics = diagnostics;
        BytesRead = bytesRead;
    }

    /// <summary>Parsed mailbox entry.</summary>
    public EmailMailboxEntry Entry { get; }

    /// <summary>Diagnostics produced while reading this message.</summary>
    public IReadOnlyList<EmailDiagnostic> Diagnostics { get; }

    /// <summary>Envelope and message bytes consumed for this entry.</summary>
    public long BytesRead { get; }

    /// <summary>True when this message produced at least one error diagnostic.</summary>
    public bool HasErrors => Diagnostics.Any(diagnostic => diagnostic.Severity == EmailDiagnosticSeverity.Error);
}
