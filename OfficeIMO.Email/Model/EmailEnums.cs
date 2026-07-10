namespace OfficeIMO.Email;

/// <summary>Identifies a supported email or Outlook artifact format.</summary>
public enum EmailFileFormat {
    /// <summary>The input format could not be determined.</summary>
    Unknown = 0,
    /// <summary>RFC 5322 message with MIME content.</summary>
    Eml,
    /// <summary>Outlook compound MSG item.</summary>
    OutlookMsg,
    /// <summary>Transport Neutral Encapsulation Format payload.</summary>
    Tnef,
    /// <summary>Unix mailbox archive.</summary>
    Mbox
}

/// <summary>Identifies the logical Outlook item represented by a document.</summary>
public enum OutlookItemKind {
    /// <summary>No typed Outlook item projection is available.</summary>
    Unknown = 0,
    /// <summary>E-mail message, normally IPM.Note.</summary>
    Message,
    /// <summary>Calendar appointment or meeting.</summary>
    Appointment,
    /// <summary>Contact card.</summary>
    Contact,
    /// <summary>Task item.</summary>
    Task,
    /// <summary>Journal item.</summary>
    Journal,
    /// <summary>Sticky note.</summary>
    Note
}

/// <summary>Classifies an address on a message or Outlook item.</summary>
public enum EmailRecipientKind {
    /// <summary>Unknown or source-specific recipient kind.</summary>
    Unknown = 0,
    /// <summary>Primary recipient.</summary>
    To,
    /// <summary>Carbon-copy recipient.</summary>
    Cc,
    /// <summary>Blind-carbon-copy recipient.</summary>
    Bcc,
    /// <summary>Reply-to address.</summary>
    ReplyTo,
    /// <summary>Resource recipient.</summary>
    Resource,
    /// <summary>Room recipient.</summary>
    Room
}

/// <summary>Severity assigned to a structured email diagnostic.</summary>
public enum EmailDiagnosticSeverity {
    /// <summary>Informational observation.</summary>
    Information = 0,
    /// <summary>Recoverable compatibility or fidelity warning.</summary>
    Warning,
    /// <summary>Content could not be interpreted completely.</summary>
    Error
}
