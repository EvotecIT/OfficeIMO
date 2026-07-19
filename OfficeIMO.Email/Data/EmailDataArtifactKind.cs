namespace OfficeIMO.Email.Data;

/// <summary>Identifies the OfficeIMO owner selected for an opened email-data artifact.</summary>
public enum EmailDataArtifactKind {
    /// <summary>No supported owner was selected.</summary>
    Unknown = 0,
    /// <summary>An EML, MSG, OFT, or TNEF item owned by OfficeIMO.Email.</summary>
    EmailDocument = 1,
    /// <summary>An iCalendar document owned by OfficeIMO.Email.</summary>
    Calendar = 2,
    /// <summary>A vCard document owned by OfficeIMO.Email.</summary>
    Contact = 3,
    /// <summary>A PST, OST, OLM, EMLX, mbox, Maildir, or Apple Mail directory store.</summary>
    Store = 4,
    /// <summary>An Outlook Offline Address Book v4 Full Details component or cache directory.</summary>
    OfflineAddressBook = 5
}
