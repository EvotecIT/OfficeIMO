namespace OfficeIMO.Email;

/// <summary>Identifies a supported email or Outlook artifact format.</summary>
public enum EmailFileFormat {
    /// <summary>The input format could not be determined.</summary>
    Unknown = 0,
    /// <summary>RFC 5322 message with MIME content.</summary>
    Eml = 1,
    /// <summary>Outlook compound MSG item.</summary>
    OutlookMsg = 2,
    /// <summary>Transport Neutral Encapsulation Format payload.</summary>
    Tnef = 3,
    /// <summary>Unix mailbox archive.</summary>
    Mbox = 4,
    /// <summary>Outlook item template stored in the MSG compound-file representation.</summary>
    OutlookTemplate = 5
}

/// <summary>Identifies the logical Outlook item represented by a document.</summary>
public enum OutlookItemKind {
    /// <summary>No typed Outlook item projection is available.</summary>
    Unknown = 0,
    /// <summary>E-mail message, normally IPM.Note.</summary>
    Message = 1,
    /// <summary>Calendar appointment or meeting.</summary>
    Appointment = 2,
    /// <summary>Contact card.</summary>
    Contact = 3,
    /// <summary>Task item.</summary>
    Task = 4,
    /// <summary>Journal item.</summary>
    Journal = 5,
    /// <summary>Sticky note.</summary>
    Note = 6,
    /// <summary>Personal Outlook distribution list.</summary>
    DistributionList = 7
}

/// <summary>Classifies an address on a message or Outlook item.</summary>
public enum EmailRecipientKind {
    /// <summary>Unknown or source-specific recipient kind.</summary>
    Unknown = 0,
    /// <summary>Primary recipient.</summary>
    To = 1,
    /// <summary>Carbon-copy recipient.</summary>
    Cc = 2,
    /// <summary>Blind-carbon-copy recipient.</summary>
    Bcc = 3,
    /// <summary>Reply-to address.</summary>
    ReplyTo = 4,
    /// <summary>Resource recipient.</summary>
    Resource = 5,
    /// <summary>Room recipient.</summary>
    Room = 6
}

/// <summary>Outlook/MAPI message importance.</summary>
public enum EmailMessageImportance {
    /// <summary>Low importance.</summary>
    Low = 0,
    /// <summary>Normal importance.</summary>
    Normal = 1,
    /// <summary>High importance.</summary>
    High = 2
}

/// <summary>Outlook/MAPI transport priority.</summary>
public enum EmailMessagePriority {
    /// <summary>Non-urgent priority.</summary>
    NonUrgent = -1,
    /// <summary>Normal priority.</summary>
    Normal = 0,
    /// <summary>Urgent priority.</summary>
    Urgent = 1
}

/// <summary>Severity assigned to a structured email diagnostic.</summary>
public enum EmailDiagnosticSeverity {
    /// <summary>Informational observation.</summary>
    Information = 0,
    /// <summary>Recoverable compatibility or fidelity warning.</summary>
    Warning = 1,
    /// <summary>Content could not be interpreted completely.</summary>
    Error = 2
}
