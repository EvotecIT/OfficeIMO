namespace OfficeIMO.Pdf;

/// <summary>Selectable PDF action categories understood by the sanitization engine.</summary>
[System.Flags]
public enum PdfSanitizationActionKind {
    /// <summary>Do not select any action category.</summary>
    None = 0,
    /// <summary>JavaScript actions and document-level JavaScript name trees.</summary>
    JavaScript = 1 << 0,
    /// <summary>URI actions and catalog URI base targets.</summary>
    Uri = 1 << 1,
    /// <summary>Launch actions that open a file or application.</summary>
    Launch = 1 << 2,
    /// <summary>SubmitForm actions.</summary>
    SubmitForm = 1 << 3,
    /// <summary>GoToR actions that navigate to another PDF.</summary>
    GoToR = 1 << 4,
    /// <summary>GoToE actions that navigate into an embedded PDF.</summary>
    GoToE = 1 << 5,
    /// <summary>ImportData actions.</summary>
    ImportData = 1 << 6,
    /// <summary>Movie actions.</summary>
    Movie = 1 << 7,
    /// <summary>Rendition actions.</summary>
    Rendition = 1 << 8,
    /// <summary>RichMedia actions.</summary>
    RichMedia = 1 << 9,
    /// <summary>All active action kinds removed by the legacy default policy, excluding URI actions.</summary>
    DefaultActiveContent = JavaScript | Launch | SubmitForm | GoToR | GoToE | ImportData | Movie | Rendition | RichMedia,
    /// <summary>Every selectable action kind, including all URI targets.</summary>
    All = DefaultActiveContent | Uri
}
