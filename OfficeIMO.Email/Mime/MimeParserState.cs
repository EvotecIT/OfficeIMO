namespace OfficeIMO.Email;

internal sealed class MimeParserState {
    internal MimeParserState(EmailReaderOptions options, IList<EmailDiagnostic> diagnostics,
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

    internal int PartCount { get; private set; }

    internal int AttachmentCount { get; private set; }

    internal long TotalAttachmentBytes { get; private set; }

    internal void CountPart() {
        ThrowIfCancellationRequested();
        Budget.CountPart();
        PartCount++;
    }

    internal void EnsurePendingPartCount(int pendingPartCount) {
        ThrowIfCancellationRequested();
        Budget.EnsurePendingParts(pendingPartCount);
    }

    internal void CountAttachment() {
        ThrowIfCancellationRequested();
        Budget.CountAttachmentOnly();
        AttachmentCount++;
    }

    internal void CountAttachmentBytes(long length) {
        Budget.CountAttachmentBytes(length);
        TotalAttachmentBytes = checked(TotalAttachmentBytes + length);
    }

    internal void EnsureAttachmentWithinLimits(long length) {
        ThrowIfCancellationRequested();
        Budget.EnsureAttachmentBytes(length);
    }

    internal void ThrowIfCancellationRequested() {
        CancellationToken.ThrowIfCancellationRequested();
    }
}
