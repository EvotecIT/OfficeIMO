namespace OfficeIMO.Email;

/// <summary>Format-neutral representation of an email or Outlook item.</summary>
public sealed class EmailDocument {
    private readonly List<EmailHeader> _headers = new List<EmailHeader>();
    private readonly List<EmailRecipient> _recipients = new List<EmailRecipient>();
    private readonly List<EmailAttachment> _attachments = new List<EmailAttachment>();
    private readonly Dictionary<string, object?> _properties = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);

    /// <summary>Source format used to create the document.</summary>
    public EmailFileFormat Format { get; set; }

    /// <summary>Typed Outlook item classification.</summary>
    public OutlookItemKind OutlookItemKind { get; set; } = OutlookItemKind.Message;

    /// <summary>Outlook message class, such as IPM.Note.</summary>
    public string? MessageClass { get; set; }

    /// <summary>Message subject.</summary>
    public string? Subject { get; set; }

    /// <summary>Represented author/from address.</summary>
    public EmailAddress? From { get; set; }

    /// <summary>Actual sender address when different from From.</summary>
    public EmailAddress? Sender { get; set; }

    /// <summary>Message-ID value.</summary>
    public string? MessageId { get; set; }

    /// <summary>Date declared by the message.</summary>
    public DateTimeOffset? Date { get; set; }

    /// <summary>Delivery/received timestamp when available.</summary>
    public DateTimeOffset? ReceivedDate { get; set; }

    /// <summary>Body alternatives.</summary>
    public EmailBody Body { get; } = new EmailBody();

    /// <summary>Ordered, duplicate-preserving headers.</summary>
    public IList<EmailHeader> Headers => _headers;

    /// <summary>Typed recipient collection.</summary>
    public IList<EmailRecipient> Recipients => _recipients;

    /// <summary>Attachment and embedded-item collection.</summary>
    public IList<EmailAttachment> Attachments => _attachments;

    /// <summary>Extensible source property bag for MAPI and format-specific metadata.</summary>
    public IDictionary<string, object?> Properties => _properties;

    /// <summary>Original artifact bytes when raw preservation was requested.</summary>
    public byte[]? RawSource { get; internal set; }
}
