using OfficeIMO.Email;

namespace OfficeIMO.Reader.Email;

/// <summary>Controls direct email, mailbox, calendar, and vCard ingestion.</summary>
public sealed class ReaderEmailOptions {
    /// <summary>Bounded policy for EML, MSG/OFT, and TNEF artifacts.</summary>
    public EmailReaderOptions? MessageOptions { get; set; }

    /// <summary>Bounded policy for Mbox and MBX mailboxes.</summary>
    public EmailMailboxReaderOptions? MailboxOptions { get; set; }

    /// <summary>Bounded policy for standalone iCalendar and vCard streams.</summary>
    public ContentLineReaderOptions? ContentLineOptions { get; set; }

    /// <summary>Retains decoded attachment payloads in Reader assets.</summary>
    public bool IncludeAttachmentContent { get; set; } = true;
}

/// <summary>Options for registering every email-related handler from this package.</summary>
public sealed class ReaderEmailHandlersOptions {
    /// <summary>Direct message, mailbox, calendar, and vCard options.</summary>
    public ReaderEmailOptions? Artifacts { get; set; }

    /// <summary>PST, OST, OLM, EMLX, and mailbox-directory options.</summary>
    public ReaderEmailStoreOptions? Stores { get; set; }

    /// <summary>Outlook Offline Address Book options.</summary>
    public ReaderEmailAddressBookOptions? AddressBooks { get; set; }
}
