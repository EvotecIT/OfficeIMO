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
    /// <summary>Original separator line without its line ending.</summary>
    public string? RawFromLine { get; set; }
}

/// <summary>Ordered Unix mailbox aggregate.</summary>
public sealed class EmailMailbox {
    private readonly List<EmailMailboxEntry> _messages = new List<EmailMailboxEntry>();

    /// <summary>Ordered mailbox messages.</summary>
    public IList<EmailMailboxEntry> Messages => _messages;
}
