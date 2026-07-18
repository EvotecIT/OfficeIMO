namespace OfficeIMO.Email.Store;

/// <summary>Identifies an email-store or store-item format.</summary>
public enum EmailStoreFormat {
    /// <summary>The source format could not be determined.</summary>
    Unknown = 0,
    /// <summary>Outlook Personal Storage Table.</summary>
    Pst = 1,
    /// <summary>Outlook Offline Storage Table.</summary>
    Ost = 2,
    /// <summary>Outlook for Mac archive.</summary>
    Olm = 3,
    /// <summary>Apple Mail EMLX item.</summary>
    Emlx = 4,
    /// <summary>Apple Mail, Maildir, or RFC message directory tree.</summary>
    MailboxDirectory = 5,
    /// <summary>Unix mbox mailbox archive.</summary>
    Mbox = 6
}
