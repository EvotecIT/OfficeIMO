namespace OfficeIMO.Email;

/// <summary>Format-neutral representation of an email or Outlook item.</summary>
public sealed class EmailDocument {
    private readonly List<EmailHeader> _headers = new List<EmailHeader>();
    private readonly List<EmailRecipient> _recipients = new List<EmailRecipient>();
    private readonly List<EmailAttachment> _attachments = new List<EmailAttachment>();
    private readonly List<MapiProperty> _mapiProperties = new List<MapiProperty>();
    private readonly List<TnefAttribute> _tnefAttributes = new List<TnefAttribute>();
    private readonly Dictionary<string, object?> _properties = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);

    /// <summary>Source format used to create the document.</summary>
    public EmailFileFormat Format { get; set; }

    /// <summary>Typed Outlook item classification.</summary>
    public OutlookItemKind OutlookItemKind { get; set; } = OutlookItemKind.Message;

    /// <summary>Outlook message class, such as IPM.Note.</summary>
    public string? MessageClass { get; set; }

    /// <summary>Primary MAPI code page used for legacy PT_STRING8 properties.</summary>
    public int? OutlookCodePage { get; set; }

    /// <summary>Message subject.</summary>
    public string? Subject { get; set; }

    /// <summary>Represented author/from address.</summary>
    public EmailAddress? From { get; set; }

    /// <summary>Actual sender address when different from From.</summary>
    public EmailAddress? Sender { get; set; }

    /// <summary>Recipient stamped by the message store as having received the message.</summary>
    public EmailAddress? ReceivedBy { get; set; }

    /// <summary>Represented recipient stamped by the message store.</summary>
    public EmailAddress? ReceivedRepresenting { get; set; }

    /// <summary>Message-ID value.</summary>
    public string? MessageId { get; set; }

    /// <summary>Date declared by the message.</summary>
    public DateTimeOffset? Date { get; set; }

    /// <summary>Delivery/received timestamp when available.</summary>
    public DateTimeOffset? ReceivedDate { get; set; }

    /// <summary>Body alternatives.</summary>
    public EmailBody Body { get; } = new EmailBody();

    /// <summary>Common Outlook/MAPI message metadata.</summary>
    public EmailMessageMetadata MessageMetadata { get; } = new EmailMessageMetadata();

    /// <summary>Protected-message classification and cryptographic payload handoff.</summary>
    public EmailProtectionInfo Protection { get; } = new EmailProtectionInfo();

    /// <summary>Ordered, duplicate-preserving headers.</summary>
    public IList<EmailHeader> Headers => _headers;

    /// <summary>Typed recipient collection.</summary>
    public IList<EmailRecipient> Recipients => _recipients;

    /// <summary>Attachment and embedded-item collection.</summary>
    public IList<EmailAttachment> Attachments => _attachments;

    /// <summary>Root MAPI properties, including properties not projected onto common fields.</summary>
    public IList<MapiProperty> MapiProperties => _mapiProperties;

    /// <summary>Ordered raw TNEF attributes when the source format is TNEF.</summary>
    public IList<TnefAttribute> TnefAttributes => _tnefAttributes;

    /// <summary>Typed appointment projection when <see cref="OutlookItemKind"/> is Appointment.</summary>
    public OutlookAppointment? Appointment { get; set; }

    /// <summary>Typed contact projection when <see cref="OutlookItemKind"/> is Contact.</summary>
    public OutlookContact? Contact { get; set; }

    /// <summary>Typed task projection when <see cref="OutlookItemKind"/> is Task.</summary>
    public OutlookTask? Task { get; set; }

    /// <summary>Typed journal projection when <see cref="OutlookItemKind"/> is Journal.</summary>
    public OutlookJournal? Journal { get; set; }

    /// <summary>Typed note projection when <see cref="OutlookItemKind"/> is Note.</summary>
    public OutlookNote? Note { get; set; }

    /// <summary>Extensible source property bag for MAPI and format-specific metadata.</summary>
    public IDictionary<string, object?> Properties => _properties;

    /// <summary>Original artifact bytes when raw preservation was requested.</summary>
    public byte[]? RawSource { get; internal set; }
}
