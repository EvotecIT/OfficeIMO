using OfficeIMO.Email;
using OfficeIMO.Reader.Html;

namespace OfficeIMO.Reader.Email;

/// <summary>Controls direct email, mailbox, calendar, and vCard ingestion.</summary>
public sealed class ReaderEmailOptions {
    /// <summary>Controls concealed HTML in derived Reader content. Selected bodies are inspected under either policy.</summary>
    public EmailConcealedTextPolicy ConcealedTextPolicy { get; set; }
    /// <summary>Bounded policy for EML, MSG/OFT, and TNEF artifacts. Disabling embedded messages also omits nested mail attachment text projection.</summary>
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

    /// <summary>MHTML archive projection options.</summary>
    public ReaderHtmlOptions? Mhtml { get; set; }

    /// <summary>PST, OST, OLM, EMLX, and mailbox-directory options.</summary>
    public ReaderEmailStoreOptions? Stores { get; set; }

    /// <summary>Outlook Offline Address Book options.</summary>
    public ReaderEmailAddressBookOptions? AddressBooks { get; set; }
}
