namespace OfficeIMO.Email;

internal static class EmailAttachmentContent {
    internal static byte[]? ReadOrNull(EmailAttachment attachment, long maximumBytes) {
        if (attachment.Content != null && !EmailAttachmentStreamScope.HasStagedContent(attachment)) return attachment.Content;
        if (attachment.ContentSource == null && !EmailAttachmentStreamScope.HasStagedContent(attachment)) return null;
        long? length = EmailAttachmentStreamScope.GetLength(attachment);
        if (length.HasValue && length.Value > maximumBytes) {
            throw new EmailLimitExceededException(nameof(EmailWriterOptions.MaxOutputBytes),
                length.Value, maximumBytes);
        }

        using (Stream input = EmailAttachmentStreamScope.OpenRead(attachment)) {
            if (input == null || !input.CanRead) {
                throw new InvalidDataException("The attachment content source did not return a readable stream.");
            }
            using (var output = new EmailBoundedMemoryStream(maximumBytes)) {
                byte[] buffer = new byte[81920];
                while (true) {
                    int read = input.Read(buffer, 0, buffer.Length);
                    if (read == 0) break;
                    output.Write(buffer, 0, read);
                }
                return output.ToArray();
            }
        }
    }
}
