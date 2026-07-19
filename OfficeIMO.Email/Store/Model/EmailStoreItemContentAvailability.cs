using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

/// <summary>Describes which requested item parts were locally available and where a cached source is inconclusive.</summary>
public sealed class EmailStoreItemContentAvailability {
    internal static readonly Guid PsetidCommon =
        MapiPropertySets.Common;
    private const int RemoteDownload = 0x00001000;
    private const int RemoteDelete = 0x00002000;

    private EmailStoreItemContentAvailability(
        EmailStoreItemReadParts availableParts,
        EmailStoreItemReadParts unavailableParts,
        EmailStoreItemReadParts indeterminateParts,
        bool? isHeaderOnly,
        bool isMarkedForDownload,
        bool isMarkedForRemoteDeletion,
        bool isPotentiallyPartial) {
        AvailableParts = availableParts;
        UnavailableParts = unavailableParts;
        IndeterminateParts = indeterminateParts;
        IsHeaderOnly = isHeaderOnly;
        IsMarkedForDownload = isMarkedForDownload;
        IsMarkedForRemoteDeletion = isMarkedForRemoteDeletion;
        IsPotentiallyPartial = isPotentiallyPartial;
    }

    /// <summary>Requested parts that were fully represented by the local source.</summary>
    public EmailStoreItemReadParts AvailableParts { get; }

    /// <summary>Requested parts known not to be available from the local source or current bounded read.</summary>
    public EmailStoreItemReadParts UnavailableParts { get; }

    /// <summary>Requested parts whose completeness cannot be established from an offline or partial source.</summary>
    public EmailStoreItemReadParts IndeterminateParts { get; }

    /// <summary>Whether Outlook explicitly marked an OST item as header-only; null means the source did not say.</summary>
    public bool? IsHeaderOnly { get; }

    /// <summary>Whether the MAPI message status requests downloading remote content.</summary>
    public bool IsMarkedForDownload { get; }

    /// <summary>Whether the MAPI message status requests remote deletion.</summary>
    public bool IsMarkedForRemoteDeletion { get; }

    /// <summary>Whether the underlying artifact can legitimately omit server- or sibling-stored content.</summary>
    public bool IsPotentiallyPartial { get; }

    internal static EmailStoreItemContentAvailability Create(
        EmailStoreFormat format,
        EmailDocument document,
        EmailStoreItemReadParts loadedParts,
        EmailStoreItemSummary? summary = null) {
        bool isPartialEmlx = TryGetBooleanProperty(document, "Emlx:IsPartial") == true;
        bool? headerOnly = summary?.IsHeaderOnly ?? TryGetHeaderOnly(document.MapiProperties);
        if (!headerOnly.HasValue && format != EmailStoreFormat.Ost && !isPartialEmlx) {
            headerOnly = false;
        }

        int messageStatus = GetMessageStatus(document.MapiProperties) ?? 0;
        bool markedForDownload = summary?.IsMarkedForDownload == true ||
            (messageStatus & RemoteDownload) != 0;
        bool markedForRemoteDeletion = summary?.IsMarkedForRemoteDeletion == true ||
            (messageStatus & RemoteDelete) != 0;
        bool potentiallyPartial = format == EmailStoreFormat.Ost || isPartialEmlx ||
            headerOnly == true || markedForDownload;

        EmailStoreItemReadParts available = EmailStoreItemReadParts.None;
        EmailStoreItemReadParts unavailable = EmailStoreItemReadParts.None;
        EmailStoreItemReadParts indeterminate = EmailStoreItemReadParts.None;
        foreach (EmailStoreItemReadParts part in EnumerateParts(loadedParts)) {
            switch (part) {
                case EmailStoreItemReadParts.Bodies:
                    if (HasBody(document)) available |= part;
                    else ClassifyMissing(part, headerOnly, potentiallyPartial,
                        emptyIsComplete: true, ref available, ref unavailable, ref indeterminate);
                    break;
                case EmailStoreItemReadParts.AttachmentContent:
                    if (HasAllAttachmentContent(document)) available |= part;
                    else ClassifyMissing(part, headerOnly, potentiallyPartial,
                        emptyIsComplete: false, ref available, ref unavailable, ref indeterminate);
                    break;
                case EmailStoreItemReadParts.EmbeddedItems:
                    if (HasAllEmbeddedItems(document)) available |= part;
                    else ClassifyMissing(part, headerOnly, potentiallyPartial,
                        emptyIsComplete: false, ref available, ref unavailable, ref indeterminate);
                    break;
                default:
                    available |= part;
                    break;
            }
        }

        return new EmailStoreItemContentAvailability(
            available, unavailable, indeterminate, headerOnly,
            markedForDownload, markedForRemoteDeletion, potentiallyPartial);
    }

    internal static bool? TryGetHeaderOnly(IEnumerable<MapiProperty> properties) {
        return properties.GetNullableMapiValue(MapiKnownProperties.PidLid.HeaderItem);
    }

    internal static int? GetMessageStatus(IEnumerable<MapiProperty> properties) =>
        properties.GetNullableMapiValue(MapiKnownProperties.PidTag.MessageStatus);

    private static void ClassifyMissing(EmailStoreItemReadParts part,
        bool? headerOnly, bool potentiallyPartial, bool emptyIsComplete,
        ref EmailStoreItemReadParts available,
        ref EmailStoreItemReadParts unavailable,
        ref EmailStoreItemReadParts indeterminate) {
        if (headerOnly == true) unavailable |= part;
        else if (potentiallyPartial) indeterminate |= part;
        else if (emptyIsComplete) available |= part;
        else unavailable |= part;
    }

    private static IEnumerable<EmailStoreItemReadParts> EnumerateParts(EmailStoreItemReadParts parts) {
        EmailStoreItemReadParts[] known = {
            EmailStoreItemReadParts.Metadata,
            EmailStoreItemReadParts.Bodies,
            EmailStoreItemReadParts.Recipients,
            EmailStoreItemReadParts.AttachmentMetadata,
            EmailStoreItemReadParts.AttachmentContent,
            EmailStoreItemReadParts.EmbeddedItems,
            EmailStoreItemReadParts.ExtendedMapiProperties
        };
        foreach (EmailStoreItemReadParts part in known) {
            if ((parts & part) != 0) yield return part;
        }
    }

    private static bool HasBody(EmailDocument document) =>
        !string.IsNullOrEmpty(document.Body.Text) ||
        !string.IsNullOrEmpty(document.Body.Html) ||
        !string.IsNullOrEmpty(document.Body.Rtf);

    private static bool HasAllAttachmentContent(EmailDocument document) {
        foreach (EmailAttachment attachment in document.Attachments) {
            if (attachment.MapiAttachMethod == 5) continue;
            if (attachment.Content == null && attachment.ContentSource == null) return false;
        }
        return true;
    }

    private static bool HasAllEmbeddedItems(EmailDocument document) {
        foreach (EmailAttachment attachment in document.Attachments) {
            if (attachment.MapiAttachMethod == 5 && attachment.EmbeddedDocument == null) return false;
        }
        return true;
    }

    private static bool? TryGetBooleanProperty(EmailDocument document, string key) {
        if (!document.Properties.TryGetValue(key, out object? value)) return null;
        return ToBoolean(value);
    }

    private static bool? ToBoolean(object? value) {
        if (value is bool boolean) return boolean;
        if (value is int integer) return integer != 0;
        if (value is short shortInteger) return shortInteger != 0;
        if (value is uint unsignedInteger) return unsignedInteger != 0;
        return null;
    }

}
