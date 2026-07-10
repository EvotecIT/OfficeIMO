namespace OfficeIMO.Email;

/// <summary>Result of bounded mailbox ingestion.</summary>
public sealed class EmailMailboxReadResult {
    internal EmailMailboxReadResult(EmailMailbox mailbox, IReadOnlyList<EmailDiagnostic> diagnostics, long bytesRead) {
        Mailbox = mailbox;
        Diagnostics = diagnostics;
        BytesRead = bytesRead;
    }

    /// <summary>Parsed mailbox.</summary>
    public EmailMailbox Mailbox { get; }
    /// <summary>Aggregate and message diagnostics.</summary>
    public IReadOnlyList<EmailDiagnostic> Diagnostics { get; }
    /// <summary>Number of source bytes consumed.</summary>
    public long BytesRead { get; }
}
