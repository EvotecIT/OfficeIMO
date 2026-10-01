namespace OfficeIMO.Email;

internal static partial class MimeWriter {
    internal static bool CanPreservePartHeaders(EmailAttachment attachment) =>
        attachment.PreserveMimeHeadersOnWrite && attachment.MimeHeaders.Count > 0 &&
        (attachment.EmbeddedDocument == null || attachment.Content != null || attachment.ContentSource != null ||
         EmailAttachmentStreamScope.HasStagedContent(attachment));
}
