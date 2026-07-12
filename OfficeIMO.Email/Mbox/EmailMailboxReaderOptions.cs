namespace OfficeIMO.Email;

/// <summary>Immutable policy for bounded mbox reading.</summary>
public sealed class EmailMailboxReaderOptions {
    /// <summary>Default mailbox reader policy.</summary>
    public static EmailMailboxReaderOptions Default { get; } = new EmailMailboxReaderOptions();

    /// <summary>Creates mailbox reader options.</summary>
    public EmailMailboxReaderOptions(EmailReaderOptions? messageOptions = null, MboxVariant variant = MboxVariant.Auto,
        int maxMessageCount = 100000) {
        if (maxMessageCount <= 0) throw new ArgumentOutOfRangeException(nameof(maxMessageCount));
        MessageOptions = messageOptions ?? EmailReaderOptions.Default;
        Variant = variant;
        MaxMessageCount = maxMessageCount;
    }

    /// <summary>Bounded policy applied to each message and the aggregate input.</summary>
    public EmailReaderOptions MessageOptions { get; }
    /// <summary>Escaping convention to decode.</summary>
    public MboxVariant Variant { get; }
    /// <summary>Maximum messages in one mailbox.</summary>
    public int MaxMessageCount { get; }
}
