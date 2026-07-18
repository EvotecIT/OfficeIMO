namespace OfficeIMO.Email;

/// <summary>Supported Unix mailbox escaping conventions.</summary>
public enum MboxVariant {
    /// <summary>Detect mboxo or mboxrd escaping while reading.</summary>
    Auto = 0,
    /// <summary>Escape only body lines that begin with "From ".</summary>
    Mboxo = 1,
    /// <summary>Escape every body line matching one or more angle brackets followed by "From ".</summary>
    Mboxrd = 2
}

/// <summary>One message and its mbox envelope metadata.</summary>
public sealed class EmailMailboxEntry {
    /// <summary>Creates a mailbox entry.</summary>
    public EmailMailboxEntry(EmailDocument document) {
        Document = document ?? throw new ArgumentNullException(nameof(document));
    }

    /// <summary>Parsed email document.</summary>
    public EmailDocument Document { get; }
    /// <summary>Envelope sender from the separator line.</summary>
    public string? EnvelopeSender { get; set; }
    /// <summary>Envelope timestamp when parseable.</summary>
    public DateTimeOffset? EnvelopeDate { get; set; }
    /// <summary>Original separator line without its line ending; a safe ASCII value is preserved on write.</summary>
    public string? RawFromLine { get; set; }
}

/// <summary>Ordered Unix mailbox aggregate.</summary>
public sealed class EmailMailbox {
    private readonly List<EmailMailboxEntry> _messages = new List<EmailMailboxEntry>();

    /// <summary>Ordered mailbox messages.</summary>
    public IList<EmailMailboxEntry> Messages => _messages;

    /// <summary>Loads an mbox mailbox from a file.</summary>
    public static EmailMailbox Load(string path, EmailMailboxReaderOptions? options = null, CancellationToken cancellationToken = default) =>
        RequireMailbox(new EmailMailboxReader(options ?? EmailMailboxReaderOptions.Default).Read(path, cancellationToken));

    /// <summary>Loads an mbox mailbox from memory.</summary>
    public static EmailMailbox Load(byte[] bytes, EmailMailboxReaderOptions? options = null, CancellationToken cancellationToken = default) =>
        RequireMailbox(new EmailMailboxReader(options ?? EmailMailboxReaderOptions.Default).Read(bytes, cancellationToken));

    /// <summary>Loads from the beginning of a seekable caller-owned stream and restores its position; non-seekable streams are read forward.</summary>
    public static EmailMailbox Load(Stream stream, EmailMailboxReaderOptions? options = null, CancellationToken cancellationToken = default) =>
        RequireMailbox(new EmailMailboxReader(options ?? EmailMailboxReaderOptions.Default).Read(stream, cancellationToken));

    /// <summary>Asynchronously loads an mbox mailbox from a file.</summary>
    public static async Task<EmailMailbox> LoadAsync(string path, EmailMailboxReaderOptions? options = null, CancellationToken cancellationToken = default) =>
        RequireMailbox(await new EmailMailboxReader(options ?? EmailMailboxReaderOptions.Default).ReadAsync(path, cancellationToken).ConfigureAwait(false));

    /// <summary>Asynchronously loads from the beginning of a seekable caller-owned stream and restores its position; non-seekable streams are read forward.</summary>
    public static async Task<EmailMailbox> LoadAsync(Stream stream, EmailMailboxReaderOptions? options = null, CancellationToken cancellationToken = default) =>
        RequireMailbox(await new EmailMailboxReader(options ?? EmailMailboxReaderOptions.Default).ReadAsync(stream, cancellationToken).ConfigureAwait(false));

    /// <summary>Serializes this mailbox to mbox bytes.</summary>
    public byte[] ToBytes(EmailMailboxWriterOptions? options = null) =>
        new EmailMailboxWriter(options ?? EmailMailboxWriterOptions.Default).ToBytes(this);

    /// <summary>Serializes the mailbox to a new writable memory stream positioned at the beginning.</summary>
    public MemoryStream ToStream(EmailMailboxWriterOptions? options = null) =>
        new MemoryStream(ToBytes(options));

    /// <summary>Saves this mailbox to a file.</summary>
    public EmailWriteResult Save(string path, EmailMailboxWriterOptions? options = null) =>
        new EmailMailboxWriter(options ?? EmailMailboxWriterOptions.Default).Write(this, path);

    /// <summary>Saves this mailbox to a caller-owned stream without closing it.</summary>
    public EmailWriteResult Save(Stream stream, EmailMailboxWriterOptions? options = null) =>
        new EmailMailboxWriter(options ?? EmailMailboxWriterOptions.Default).Write(this, stream);

    /// <summary>Asynchronously saves this mailbox to a file.</summary>
    public Task<EmailWriteResult> SaveAsync(string path, EmailMailboxWriterOptions? options = null, CancellationToken cancellationToken = default) =>
        new EmailMailboxWriter(options ?? EmailMailboxWriterOptions.Default).WriteAsync(this, path, cancellationToken);

    /// <summary>Asynchronously saves this mailbox to a caller-owned stream without closing it.</summary>
    public Task<EmailWriteResult> SaveAsync(Stream stream, EmailMailboxWriterOptions? options = null, CancellationToken cancellationToken = default) =>
        new EmailMailboxWriter(options ?? EmailMailboxWriterOptions.Default).WriteAsync(this, stream, cancellationToken);

    private static EmailMailbox RequireMailbox(EmailMailboxReadResult result) {
        EmailDiagnostic? error = result.Diagnostics.FirstOrDefault(static diagnostic =>
            diagnostic.Severity == EmailDiagnosticSeverity.Error);
        if (error != null) {
            throw new InvalidDataException("The mailbox could not be loaded: " + error.Code + ": " + error.Message);
        }
        return result.Mailbox;
    }
}
