using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

/// <summary>Small projected item view intended for browsing and bounded store queries.</summary>
public sealed class EmailStoreItemSummary {
    internal EmailStoreItemSummary(EmailDocument document, bool? hasAttachments, bool? isRead) {
        OutlookItemKind = document.OutlookItemKind;
        MessageClass = document.MessageClass;
        Subject = document.Subject;
        MessageId = document.MessageId;
        From = document.From;
        Sender = document.Sender;
        SentAt = document.Date;
        ReceivedAt = document.ReceivedDate;
        HasAttachments = hasAttachments;
        IsRead = isRead;
        DeclaredSize = document.MessageMetadata.DeclaredSize;
    }

    internal static EmailStoreItemSummary FromItem(EmailStoreItem item) =>
        new EmailStoreItemSummary(
            item.Document,
            item.Document.Attachments.Count > 0,
            item.Document.MessageMetadata.IsRead);

    /// <summary>Projected Outlook item kind.</summary>
    public OutlookItemKind OutlookItemKind { get; }

    /// <summary>MAPI message class when available.</summary>
    public string? MessageClass { get; }

    /// <summary>Item subject.</summary>
    public string? Subject { get; }

    /// <summary>Internet message identifier when available.</summary>
    public string? MessageId { get; }

    /// <summary>Represented sender.</summary>
    public EmailAddress? From { get; }

    /// <summary>Actual sender.</summary>
    public EmailAddress? Sender { get; }

    /// <summary>Sent or created time.</summary>
    public DateTimeOffset? SentAt { get; }

    /// <summary>Received time.</summary>
    public DateTimeOffset? ReceivedAt { get; }

    /// <summary>Whether the source declares attachments.</summary>
    public bool? HasAttachments { get; }

    /// <summary>Whether the source declares the item read.</summary>
    public bool? IsRead { get; }

    /// <summary>Declared source size when available.</summary>
    public int? DeclaredSize { get; }
}
