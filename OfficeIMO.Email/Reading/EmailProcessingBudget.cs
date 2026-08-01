namespace OfficeIMO.Email;

/// <summary>Immutable observed resource usage for one bounded email read.</summary>
public sealed class EmailProcessingBudgetSnapshot {
    internal EmailProcessingBudgetSnapshot(long inputBytes, int partCount, int propertyCount,
        long decodedPropertyBytes, int attachmentCount, long attachmentBytes, int tnefAttributeCount) {
        InputBytes = inputBytes;
        PartCount = partCount;
        PropertyCount = propertyCount;
        DecodedPropertyBytes = decodedPropertyBytes;
        AttachmentCount = attachmentCount;
        AttachmentBytes = attachmentBytes;
        TnefAttributeCount = tnefAttributeCount;
    }

    /// <summary>Source bytes accounted to the operation.</summary>
    public long InputBytes { get; }
    /// <summary>MIME entities accounted across the message and embedded messages.</summary>
    public int PartCount { get; }
    /// <summary>MAPI properties accounted across MSG and TNEF structures.</summary>
    public int PropertyCount { get; }
    /// <summary>Decoded MAPI property bytes accounted across the operation.</summary>
    public long DecodedPropertyBytes { get; }
    /// <summary>Attachments accounted across nested MIME, MSG, and TNEF structures.</summary>
    public int AttachmentCount { get; }
    /// <summary>Decoded attachment bytes accounted across the operation.</summary>
    public long AttachmentBytes { get; }
    /// <summary>TNEF attributes accounted across the operation.</summary>
    public int TnefAttributeCount { get; }
}

/// <summary>
/// Shared resource ledger for all parsers participating in one email read. The ledger prevents an embedded
/// parser from receiving a fresh attachment, property, or structural allowance.
/// </summary>
internal sealed class EmailProcessingBudget {
    private readonly EmailReaderOptions _limits;

    internal EmailProcessingBudget(EmailReaderOptions limits) {
        _limits = limits ?? throw new ArgumentNullException(nameof(limits));
    }

    internal long InputBytes { get; private set; }
    internal int PartCount { get; private set; }
    internal int PropertyCount { get; private set; }
    internal long DecodedPropertyBytes { get; private set; }
    internal int AttachmentCount { get; private set; }
    internal long AttachmentBytes { get; private set; }
    internal int TnefAttributeCount { get; private set; }

    internal void CountInput(long bytes) {
        if (bytes < 0) throw new ArgumentOutOfRangeException(nameof(bytes));
        long total = checked(InputBytes + bytes);
        Ensure(nameof(EmailReaderOptions.MaxInputBytes), total, _limits.MaxInputBytes);
        InputBytes = total;
    }

    internal void CountPart(int count = 1) {
        long total = checked((long)PartCount + count);
        Ensure(nameof(EmailReaderOptions.MaxPartCount), total, _limits.MaxPartCount);
        PartCount = checked((int)total);
    }

    internal void EnsurePendingParts(int count) =>
        Ensure(nameof(EmailReaderOptions.MaxPartCount), checked((long)PartCount + count), _limits.MaxPartCount);

    internal void CountProperty(long decodedBytes) {
        long properties = checked((long)PropertyCount + 1);
        Ensure(nameof(EmailReaderOptions.MaxMapiPropertyCount), properties, _limits.MaxMapiPropertyCount);
        EnsureDecodedPropertyBytes(decodedBytes);
        PropertyCount = checked((int)properties);
        DecodedPropertyBytes = checked(DecodedPropertyBytes + decodedBytes);
    }

    internal void CountDecodedPropertyBytes(long bytes) {
        EnsureDecodedPropertyBytes(bytes);
        DecodedPropertyBytes = checked(DecodedPropertyBytes + bytes);
    }

    internal void EnsureDecodedPropertyBytes(long bytes) =>
        Ensure(nameof(EmailReaderOptions.MaxDecodedPropertyBytes),
            checked(DecodedPropertyBytes + bytes), _limits.MaxDecodedPropertyBytes);

    internal void CountAttachment(long bytes) {
        long count = checked((long)AttachmentCount + 1);
        Ensure(nameof(EmailReaderOptions.MaxAttachmentCount), count, _limits.MaxAttachmentCount);
        EnsureAttachmentBytes(bytes);
        AttachmentCount = checked((int)count);
        AttachmentBytes = checked(AttachmentBytes + bytes);
    }

    internal void CountAttachmentOnly() {
        long count = checked((long)AttachmentCount + 1);
        Ensure(nameof(EmailReaderOptions.MaxAttachmentCount), count, _limits.MaxAttachmentCount);
        AttachmentCount = checked((int)count);
    }

    internal void CountAttachmentBytes(long bytes) {
        EnsureAttachmentBytes(bytes);
        AttachmentBytes = checked(AttachmentBytes + bytes);
    }

    internal void EnsureAttachmentBytes(long bytes, long pendingBytes = 0) {
        Ensure(nameof(EmailReaderOptions.MaxAttachmentBytes), bytes, _limits.MaxAttachmentBytes);
        Ensure(nameof(EmailReaderOptions.MaxTotalAttachmentBytes),
            checked(AttachmentBytes + pendingBytes + bytes), _limits.MaxTotalAttachmentBytes);
    }

    internal void CountTnefAttribute() {
        long count = checked((long)TnefAttributeCount + 1);
        Ensure(nameof(EmailReaderOptions.MaxTnefAttributeCount), count, _limits.MaxTnefAttributeCount);
        TnefAttributeCount = checked((int)count);
    }

    internal EmailProcessingBudgetSnapshot Snapshot() => new EmailProcessingBudgetSnapshot(
        InputBytes, PartCount, PropertyCount, DecodedPropertyBytes,
        AttachmentCount, AttachmentBytes, TnefAttributeCount);

    private static void Ensure(string name, long actual, long maximum) {
        if (actual > maximum) throw new EmailLimitExceededException(name, actual, maximum);
    }
}
