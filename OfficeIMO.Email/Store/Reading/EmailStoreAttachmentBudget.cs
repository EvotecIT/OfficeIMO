using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

internal static class EmailStoreAttachmentBudget {
    internal static long AddDocument(EmailDocument document, long current, long maximum) {
        foreach (EmailAttachment attachment in document.Attachments) {
            long length = attachment.Content != null
                ? attachment.Content.LongLength
                : attachment.ContentSource?.Length ?? attachment.Length;
            if (length < 0) length = 0;
            if (length > maximum || current > maximum - length) {
                long actual = length > long.MaxValue - current
                    ? long.MaxValue
                    : current + length;
                throw new EmailStoreLimitExceededException(
                    nameof(EmailStoreReaderOptions.MaxTotalAttachmentBytes), actual, maximum);
            }
            current += length;
            if (attachment.EmbeddedDocument != null) {
                current = AddDocument(attachment.EmbeddedDocument, current, maximum);
            }
        }
        return current;
    }
}
