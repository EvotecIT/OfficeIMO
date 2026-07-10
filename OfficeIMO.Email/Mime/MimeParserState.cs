namespace OfficeIMO.Email;

internal sealed class MimeParserState {
    internal MimeParserState(EmailReaderOptions options, IList<EmailDiagnostic> diagnostics) {
        Options = options;
        Diagnostics = diagnostics;
    }

    internal EmailReaderOptions Options { get; }

    internal IList<EmailDiagnostic> Diagnostics { get; }

    internal int PartCount { get; private set; }

    internal long TotalAttachmentBytes { get; private set; }

    internal void CountPart() {
        PartCount++;
        if (PartCount > Options.MaxPartCount) {
            throw new EmailLimitExceededException(nameof(EmailReaderOptions.MaxPartCount), PartCount, Options.MaxPartCount);
        }
    }

    internal void CountAttachment(long length) {
        if (length > Options.MaxAttachmentBytes) {
            throw new EmailLimitExceededException(nameof(EmailReaderOptions.MaxAttachmentBytes), length, Options.MaxAttachmentBytes);
        }
        TotalAttachmentBytes += length;
        if (TotalAttachmentBytes > Options.MaxTotalAttachmentBytes) {
            throw new EmailLimitExceededException(nameof(EmailReaderOptions.MaxTotalAttachmentBytes),
                TotalAttachmentBytes, Options.MaxTotalAttachmentBytes);
        }
    }
}
