using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

/// <summary>Small projected item view intended for browsing and bounded store queries.</summary>
public sealed class EmailStoreItemSummary {
    internal EmailStoreItemSummary(EmailDocument document, bool? hasAttachments, bool? isRead,
        bool? isHeaderOnly = null, bool isMarkedForDownload = false,
        bool isMarkedForRemoteDeletion = false) {
        OutlookItemKind = document.OutlookItemKind;
        MessageClass = document.MessageClass;
        Subject = document.Subject;
        MessageId = document.MessageId;
        NormalizedSubject = document.MessageMetadata.NormalizedSubject;
        ConversationTopic = document.MessageMetadata.ConversationTopic;
        ConversationIndex = Copy(document.MessageMetadata.ConversationIndex);
        ConversationId = Copy(document.MessageMetadata.ConversationId);
        InternetReferences = document.MessageMetadata.InternetReferences;
        InReplyToId = document.MessageMetadata.InReplyToId;
        MeetingGlobalObjectId = Copy(document.Appointment?.GlobalObjectId);
        MeetingCleanGlobalObjectId = Copy(document.Appointment?.CleanGlobalObjectId);
        TaskGlobalId = document.Task?.GlobalId;
        From = document.From;
        Sender = document.Sender;
        SentAt = document.Date;
        ReceivedAt = document.ReceivedDate;
        HasAttachments = hasAttachments;
        IsRead = isRead;
        DeclaredSize = document.MessageMetadata.DeclaredSize;
        IsHeaderOnly = isHeaderOnly;
        IsMarkedForDownload = isMarkedForDownload;
        IsMarkedForRemoteDeletion = isMarkedForRemoteDeletion;
    }

    internal static EmailStoreItemSummary FromItem(EmailStoreItem item) =>
        new EmailStoreItemSummary(
            item.Document,
            item.Document.Attachments.Count > 0,
            item.Document.MessageMetadata.IsRead,
            item.ContentAvailability.IsHeaderOnly,
            item.ContentAvailability.IsMarkedForDownload,
            item.ContentAvailability.IsMarkedForRemoteDeletion);

    internal static EmailStoreItemSummary FromMetadata(EmailDocument document,
        EmailStoreItemSummary fallback) => new EmailStoreItemSummary(
            document,
            fallback.HasAttachments,
            fallback.IsRead,
            fallback.IsHeaderOnly,
            fallback.IsMarkedForDownload,
            fallback.IsMarkedForRemoteDeletion);

    /// <summary>Projected Outlook item kind.</summary>
    public OutlookItemKind OutlookItemKind { get; }

    /// <summary>MAPI message class when available.</summary>
    public string? MessageClass { get; }

    /// <summary>Item subject.</summary>
    public string? Subject { get; }

    /// <summary>Internet message identifier when available.</summary>
    public string? MessageId { get; }

    /// <summary>Subject normalized by the source or Outlook projection.</summary>
    public string? NormalizedSubject { get; }

    /// <summary>Outlook conversation topic.</summary>
    public string? ConversationTopic { get; }

    /// <summary>Outlook conversation index.</summary>
    public byte[]? ConversationIndex { get; }

    /// <summary>Binary Outlook conversation identifier.</summary>
    public byte[]? ConversationId { get; }

    /// <summary>Internet References field.</summary>
    public string? InternetReferences { get; }

    /// <summary>Internet In-Reply-To identifier.</summary>
    public string? InReplyToId { get; }

    /// <summary>Meeting Global Object ID.</summary>
    public byte[]? MeetingGlobalObjectId { get; }

    /// <summary>Clean Meeting Global Object ID, preferred for recurrence-series correlation.</summary>
    public byte[]? MeetingCleanGlobalObjectId { get; }

    /// <summary>Task lifecycle correlation identifier.</summary>
    public Guid? TaskGlobalId { get; }

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

    /// <summary>Whether Outlook explicitly marked an OST item as header-only; null means unspecified.</summary>
    public bool? IsHeaderOnly { get; }

    /// <summary>Whether the MAPI message status requests downloading remote content.</summary>
    public bool IsMarkedForDownload { get; }

    /// <summary>Whether the MAPI message status requests remote deletion.</summary>
    public bool IsMarkedForRemoteDeletion { get; }

    private static byte[]? Copy(byte[]? value) => value == null ? null : (byte[])value.Clone();
}
