using OfficeIMO.Email.AddressBook;
using OfficeIMO.Email.Store;

namespace OfficeIMO.Email.Data;

/// <summary>Immutable owner-specific policies for <see cref="EmailDataArtifact.Open"/>.</summary>
public sealed class EmailDataOpenOptions {
    /// <summary>Default bounded discovery and reader policy.</summary>
    public static EmailDataOpenOptions Default { get; } = new EmailDataOpenOptions();

    /// <summary>Creates a discovery policy without replacing any underlying owner's limits.</summary>
    public EmailDataOpenOptions(
        EmailReaderOptions? email = null,
        ContentLineReaderOptions? contentLines = null,
        EmailStoreReaderOptions? store = null,
        OfflineAddressBookReaderOptions? addressBook = null,
        EmailDataArtifactKind? expectedKind = null,
        bool useStreamingEmailReader = false) {
        if (expectedKind == EmailDataArtifactKind.Unknown)
            throw new ArgumentOutOfRangeException(nameof(expectedKind));
        Email = email ?? EmailReaderOptions.Default;
        ContentLines = contentLines ?? ContentLineReaderOptions.Default;
        Store = store ?? EmailStoreReaderOptions.Default;
        AddressBook = addressBook ?? OfflineAddressBookReaderOptions.Default;
        ExpectedKind = expectedKind;
        UseStreamingEmailReader = useStreamingEmailReader;
    }

    /// <summary>Policy used by the individual email artifact reader.</summary>
    public EmailReaderOptions Email { get; }
    /// <summary>Policy used by iCalendar and vCard readers.</summary>
    public ContentLineReaderOptions ContentLines { get; }
    /// <summary>Policy used by mailbox-store sessions.</summary>
    public EmailStoreReaderOptions Store { get; }
    /// <summary>Policy used by OAB inspection and sessions.</summary>
    public OfflineAddressBookReaderOptions AddressBook { get; }
    /// <summary>Optional explicit owner for ambiguous paths or extension-free content.</summary>
    public EmailDataArtifactKind? ExpectedKind { get; }
    /// <summary>
    /// Whether individual EML/MSG/OFT/TNEF payloads use file-backed streaming reads. The default full read preserves
    /// the existing protected-artifact pass-through contract.
    /// </summary>
    public bool UseStreamingEmailReader { get; }
}
