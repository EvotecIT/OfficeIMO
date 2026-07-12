namespace OfficeIMO.Email;

/// <summary>Immutable deterministic mbox writer policy.</summary>
public sealed class EmailMailboxWriterOptions {
    /// <summary>Default mboxrd writer policy.</summary>
    public static EmailMailboxWriterOptions Default { get; } = new EmailMailboxWriterOptions();

    /// <summary>Creates mailbox writer options.</summary>
    public EmailMailboxWriterOptions(EmailWriterOptions? messageOptions = null, MboxVariant variant = MboxVariant.Mboxrd) {
        if (variant == MboxVariant.Auto) throw new ArgumentException("A concrete mbox variant is required for writing.", nameof(variant));
        MessageOptions = messageOptions ?? EmailWriterOptions.Default;
        Variant = variant;
    }

    /// <summary>EML serialization policy for each message.</summary>
    public EmailWriterOptions MessageOptions { get; }
    /// <summary>Escaping convention to write.</summary>
    public MboxVariant Variant { get; }
}
