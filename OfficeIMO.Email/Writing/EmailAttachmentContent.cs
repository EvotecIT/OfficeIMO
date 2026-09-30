namespace OfficeIMO.Email;

internal static class EmailAttachmentContent {
    internal static byte[]? ReadOrNull(EmailAttachment attachment, long maximumBytes, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (maximumBytes < 0) throw new ArgumentOutOfRangeException(nameof(maximumBytes));
        if (attachment.Content != null && !EmailAttachmentStreamScope.HasStagedContent(attachment)) {
            if (attachment.Content.LongLength > maximumBytes) throw new EmailLimitExceededException(nameof(EmailWriterOptions.MaxOutputBytes), attachment.Content.LongLength, maximumBytes);
            return attachment.Content;
        }
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
                    cancellationToken.ThrowIfCancellationRequested();
                    long remaining = maximumBytes - output.Length;
                    int read = input.Read(buffer, 0, remaining >= buffer.Length ? buffer.Length : (int)remaining + 1);
                    cancellationToken.ThrowIfCancellationRequested();
                    if (read == 0) break;
                    output.Write(buffer, 0, read);
                }
                return output.ToArray();
            }
        }
    }
}
