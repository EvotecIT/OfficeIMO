namespace OfficeIMO.Email;

internal sealed class MsgParserState {
    internal MsgParserState(EmailReaderOptions options, IList<EmailDiagnostic> diagnostics,
        CancellationToken cancellationToken, EmailProcessingBudget? budget = null) {
        Options = options;
        Diagnostics = diagnostics;
        CancellationToken = cancellationToken;
        Budget = budget ?? new EmailProcessingBudget(options);
    }

    internal EmailReaderOptions Options { get; }

    internal IList<EmailDiagnostic> Diagnostics { get; }

    internal CancellationToken CancellationToken { get; }

    internal EmailProcessingBudget Budget { get; }

    internal int PropertyCount { get; private set; }

    internal long DecodedPropertyBytes { get; private set; }

    internal long RemainingDecodedPropertyBytes => Options.MaxDecodedPropertyBytes - DecodedPropertyBytes;

    internal int AttachmentCount { get; private set; }

    internal long TotalAttachmentBytes { get; private set; }

    internal int TnefAttributeCount { get; private set; }

    internal void CountProperty(int bytes) {
        ThrowIfCancellationRequested();
        Budget.CountProperty(bytes);
        PropertyCount++;
        DecodedPropertyBytes = checked(DecodedPropertyBytes + bytes);
    }

    internal void CountDecodedBytes(int bytes) {
        ThrowIfCancellationRequested();
        Budget.CountDecodedPropertyBytes(bytes);
        DecodedPropertyBytes = checked(DecodedPropertyBytes + bytes);
    }

    internal void EnsureDecodedPropertyBytesWithinLimits(long bytes) {
        ThrowIfCancellationRequested();
        Budget.EnsureDecodedPropertyBytes(bytes);
    }

    internal void CountAttachment(long bytes) {
        ThrowIfCancellationRequested();
        Budget.CountAttachment(bytes);
        AttachmentCount++;
        TotalAttachmentBytes = checked(TotalAttachmentBytes + bytes);
    }

    internal void EnsureAttachmentBytesWithinLimits(long bytes, long pendingTotalBytes = 0) {
        ThrowIfCancellationRequested();
        Budget.EnsureAttachmentBytes(bytes, pendingTotalBytes);
    }

    internal void CountTnefAttribute() {
        ThrowIfCancellationRequested();
        Budget.CountTnefAttribute();
        TnefAttributeCount++;
    }

    internal void ThrowIfCancellationRequested() {
        CancellationToken.ThrowIfCancellationRequested();
    }
}
