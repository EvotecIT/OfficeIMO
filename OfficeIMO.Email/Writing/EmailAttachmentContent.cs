namespace OfficeIMO.Email;

internal static class EmailAttachmentContent {
    internal static byte[]? ReadOrNull(EmailAttachment attachment, long maximumBytes) {
        if (attachment.Content != null) return attachment.Content;
        if (attachment.ContentSource == null) return null;
        if (attachment.ContentSource.Length.HasValue && attachment.ContentSource.Length.Value > maximumBytes) {
            throw new EmailLimitExceededException(nameof(EmailWriterOptions.MaxOutputBytes),
                attachment.ContentSource.Length.Value, maximumBytes);
        }

        using (Stream input = attachment.ContentSource.OpenRead()) {
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
