namespace OfficeIMO.Email;

/// <summary>Immutable policy for bounded mbox reading.</summary>
public sealed class EmailMailboxReaderOptions {
    /// <summary>Default mailbox reader policy.</summary>
    public static EmailMailboxReaderOptions Default { get; } = new EmailMailboxReaderOptions();

    /// <summary>Creates mailbox reader options.</summary>
    public EmailMailboxReaderOptions(EmailReaderOptions? messageOptions = null, MboxVariant variant = MboxVariant.Auto,
        int maxMessageCount = 100000)
        : this(512L * 1024L * 1024L, messageOptions, variant, maxMessageCount) { }

    /// <summary>Creates mailbox reader options with an explicit aggregate byte limit.</summary>
    public EmailMailboxReaderOptions(long maxMailboxBytes, EmailReaderOptions? messageOptions = null,
        MboxVariant variant = MboxVariant.Auto, int maxMessageCount = 100000) {
        if (maxMessageCount <= 0) throw new ArgumentOutOfRangeException(nameof(maxMessageCount));
        if (maxMailboxBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maxMailboxBytes));
        MessageOptions = messageOptions ?? EmailReaderOptions.Default;
        Variant = variant;
        MaxMessageCount = maxMessageCount;
        MaxMailboxBytes = maxMailboxBytes;
    }

    /// <summary>Bounded policy applied independently to each message.</summary>
    public EmailReaderOptions MessageOptions { get; }
    /// <summary>Escaping convention to decode.</summary>
    public MboxVariant Variant { get; }
    /// <summary>Maximum messages in one mailbox.</summary>
    public int MaxMessageCount { get; }
    /// <summary>Maximum aggregate source bytes consumed from one mailbox.</summary>
    public long MaxMailboxBytes { get; }
}
