namespace OfficeIMO.Email;

internal sealed class MsgParserState {
    internal MsgParserState(EmailReaderOptions options, IList<EmailDiagnostic> diagnostics, CancellationToken cancellationToken) {
        Options = options;
        Diagnostics = diagnostics;
        CancellationToken = cancellationToken;
    }

    internal EmailReaderOptions Options { get; }

    internal IList<EmailDiagnostic> Diagnostics { get; }

    internal CancellationToken CancellationToken { get; }

    internal int PropertyCount { get; private set; }

    internal long DecodedPropertyBytes { get; private set; }

    internal int AttachmentCount { get; private set; }

    internal long TotalAttachmentBytes { get; private set; }

    internal int TnefAttributeCount { get; private set; }

    internal void CountProperty(int bytes) {
        ThrowIfCancellationRequested();
        PropertyCount++;
        if (PropertyCount > Options.MaxMapiPropertyCount) {
            throw new EmailLimitExceededException(nameof(EmailReaderOptions.MaxMapiPropertyCount),
                PropertyCount, Options.MaxMapiPropertyCount);
        }
        CountDecodedBytes(bytes);
    }

    internal void CountDecodedBytes(int bytes) {
        ThrowIfCancellationRequested();
        DecodedPropertyBytes = checked(DecodedPropertyBytes + bytes);
        if (DecodedPropertyBytes > Options.MaxDecodedPropertyBytes) {
            throw new EmailLimitExceededException(nameof(EmailReaderOptions.MaxDecodedPropertyBytes),
                DecodedPropertyBytes, Options.MaxDecodedPropertyBytes);
        }
    }

    internal void CountAttachment(long bytes) {
        ThrowIfCancellationRequested();
        AttachmentCount++;
        if (AttachmentCount > Options.MaxPartCount) {
            throw new EmailLimitExceededException(nameof(EmailReaderOptions.MaxPartCount), AttachmentCount, Options.MaxPartCount);
        }
        if (bytes > Options.MaxAttachmentBytes) {
            throw new EmailLimitExceededException(nameof(EmailReaderOptions.MaxAttachmentBytes), bytes, Options.MaxAttachmentBytes);
        }
        TotalAttachmentBytes = checked(TotalAttachmentBytes + bytes);
        if (TotalAttachmentBytes > Options.MaxTotalAttachmentBytes) {
            throw new EmailLimitExceededException(nameof(EmailReaderOptions.MaxTotalAttachmentBytes),
                TotalAttachmentBytes, Options.MaxTotalAttachmentBytes);
        }
    }

    internal void CountTnefAttribute() {
        ThrowIfCancellationRequested();
        TnefAttributeCount++;
        if (TnefAttributeCount > Options.MaxTnefAttributeCount) {
            throw new EmailLimitExceededException(nameof(EmailReaderOptions.MaxTnefAttributeCount),
                TnefAttributeCount, Options.MaxTnefAttributeCount);
        }
    }

    internal void ThrowIfCancellationRequested() {
        CancellationToken.ThrowIfCancellationRequested();
    }
}
