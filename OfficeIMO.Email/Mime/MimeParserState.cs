namespace OfficeIMO.Email;

internal sealed class MimeParserState {
    internal MimeParserState(EmailReaderOptions options, IList<EmailDiagnostic> diagnostics, CancellationToken cancellationToken) {
        Options = options;
        Diagnostics = diagnostics;
        CancellationToken = cancellationToken;
    }

    internal EmailReaderOptions Options { get; }

    internal IList<EmailDiagnostic> Diagnostics { get; }

    internal CancellationToken CancellationToken { get; }

    internal int PartCount { get; private set; }

    internal int AttachmentCount { get; private set; }

    internal long TotalAttachmentBytes { get; private set; }

    internal void CountPart() {
        ThrowIfCancellationRequested();
        PartCount++;
        if (PartCount > Options.MaxPartCount) {
            throw new EmailLimitExceededException(nameof(EmailReaderOptions.MaxPartCount), PartCount, Options.MaxPartCount);
        }
    }

    internal void EnsurePendingPartCount(int pendingPartCount) {
        ThrowIfCancellationRequested();
        long total = checked((long)PartCount + pendingPartCount);
        if (total > Options.MaxPartCount) {
            throw new EmailLimitExceededException(nameof(EmailReaderOptions.MaxPartCount),
                total, Options.MaxPartCount);
        }
    }

    internal void CountAttachment() {
        ThrowIfCancellationRequested();
        AttachmentCount++;
        if (AttachmentCount > Options.MaxAttachmentCount) {
            throw new EmailLimitExceededException(nameof(EmailReaderOptions.MaxAttachmentCount),
                AttachmentCount, Options.MaxAttachmentCount);
        }
    }

    internal void CountAttachmentBytes(long length) {
        EnsureAttachmentWithinLimits(length);
        TotalAttachmentBytes = checked(TotalAttachmentBytes + length);
    }

    internal void EnsureAttachmentWithinLimits(long length) {
        ThrowIfCancellationRequested();
        if (length > Options.MaxAttachmentBytes) {
            throw new EmailLimitExceededException(nameof(EmailReaderOptions.MaxAttachmentBytes), length, Options.MaxAttachmentBytes);
        }
        long total = checked(TotalAttachmentBytes + length);
        if (total > Options.MaxTotalAttachmentBytes) {
            throw new EmailLimitExceededException(nameof(EmailReaderOptions.MaxTotalAttachmentBytes),
                total, Options.MaxTotalAttachmentBytes);
        }
    }

    internal void ThrowIfCancellationRequested() {
        CancellationToken.ThrowIfCancellationRequested();
    }
}
