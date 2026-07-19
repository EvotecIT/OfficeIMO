using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

/// <summary>Canonical fields available to Store query, sort, and projection contracts.</summary>
public static class EmailStoreFields {
    /// <summary>Stable item identifier.</summary>
    public static EmailStoreField<EmailStoreItemId> ItemId { get; } =
        new EmailStoreField<EmailStoreItemId>("item.id", "Item ID", row => row.Reference.Key);

    /// <summary>Containing folder identifier.</summary>
    public static EmailStoreField<EmailStoreFolderId> FolderId { get; } =
        new EmailStoreField<EmailStoreFolderId>("folder.id", "Folder ID", row => row.Reference.FolderKey);

    /// <summary>Whether the item is folder-associated information.</summary>
    public static EmailStoreField<bool> IsAssociated { get; } =
        new EmailStoreField<bool>("item.associated", "Associated", row => row.Reference.IsAssociated);

    /// <summary>Whether the item was recovered outside its folder contents table.</summary>
    public static EmailStoreField<bool> IsOrphaned { get; } =
        new EmailStoreField<bool>("item.orphaned", "Orphaned", row => row.Reference.IsOrphaned);

    /// <summary>Typed Outlook item family.</summary>
    public static EmailStoreField<OutlookItemKind> OutlookItemKind { get; } =
        new EmailStoreField<OutlookItemKind>("item.kind", "Item kind", row => row.Summary.OutlookItemKind);

    /// <summary>MAPI message class.</summary>
    public static EmailStoreStringField MessageClass { get; } =
        new EmailStoreStringField("item.messageClass", "Message class", row => row.Summary.MessageClass);

    /// <summary>Item subject.</summary>
    public static EmailStoreStringField Subject { get; } =
        new EmailStoreStringField("item.subject", "Subject", row => row.Summary.Subject);

    /// <summary>Internet message identifier.</summary>
    public static EmailStoreStringField MessageId { get; } =
        new EmailStoreStringField("item.messageId", "Message ID", row => row.Summary.MessageId);

    /// <summary>Represented sender email address.</summary>
    public static EmailStoreStringField FromAddress { get; } =
        new EmailStoreStringField("from.address", "From address", row => row.Summary.From?.Address);

    /// <summary>Represented sender display name.</summary>
    public static EmailStoreStringField FromDisplayName { get; } =
        new EmailStoreStringField("from.name", "From name", row => row.Summary.From?.DisplayName);

    /// <summary>Actual sender email address.</summary>
    public static EmailStoreStringField SenderAddress { get; } =
        new EmailStoreStringField("sender.address", "Sender address", row => row.Summary.Sender?.Address);

    /// <summary>Actual sender display name.</summary>
    public static EmailStoreStringField SenderDisplayName { get; } =
        new EmailStoreStringField("sender.name", "Sender name", row => row.Summary.Sender?.DisplayName);

    /// <summary>Sent or created time.</summary>
    public static EmailStoreField<DateTimeOffset?> SentAt { get; } =
        new EmailStoreField<DateTimeOffset?>("item.sentAt", "Sent at", row => row.Summary.SentAt);

    /// <summary>Received time.</summary>
    public static EmailStoreField<DateTimeOffset?> ReceivedAt { get; } =
        new EmailStoreField<DateTimeOffset?>("item.receivedAt", "Received at", row => row.Summary.ReceivedAt);

    /// <summary>Declared attachment presence.</summary>
    public static EmailStoreField<bool?> HasAttachments { get; } =
        new EmailStoreField<bool?>("item.hasAttachments", "Has attachments", row => row.Summary.HasAttachments);

    /// <summary>Declared read state.</summary>
    public static EmailStoreField<bool?> IsRead { get; } =
        new EmailStoreField<bool?>("item.isRead", "Read", row => row.Summary.IsRead);

    /// <summary>Declared source size.</summary>
    public static EmailStoreField<int?> DeclaredSize { get; } =
        new EmailStoreField<int?>("item.declaredSize", "Declared size", row => row.Summary.DeclaredSize);

    /// <summary>OST header-only cache state.</summary>
    public static EmailStoreField<bool?> IsHeaderOnly { get; } =
        new EmailStoreField<bool?>("item.headerOnly", "Header only", row => row.Summary.IsHeaderOnly);

    /// <summary>Remote-download marker.</summary>
    public static EmailStoreField<bool> IsMarkedForDownload { get; } =
        new EmailStoreField<bool>("item.markedForDownload", "Marked for download", row => row.Summary.IsMarkedForDownload);

    /// <summary>Remote-deletion marker.</summary>
    public static EmailStoreField<bool> IsMarkedForRemoteDeletion { get; } =
        new EmailStoreField<bool>("item.markedForRemoteDeletion", "Marked for remote deletion", row => row.Summary.IsMarkedForRemoteDeletion);
}
