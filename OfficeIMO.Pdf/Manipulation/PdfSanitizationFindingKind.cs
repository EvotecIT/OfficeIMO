namespace OfficeIMO.Pdf;

/// <summary>Potentially unsafe PDF feature reported by the sanitization engine.</summary>
public enum PdfSanitizationFindingKind {
    /// <summary>An active action such as JavaScript, Launch, GoToR, SubmitForm, or ImportData.</summary>
    ActiveAction,

    /// <summary>A URI action or catalog base URI whose scheme is not allowed.</summary>
    UnsafeUri,

    /// <summary>An embedded-file or associated-file reference.</summary>
    EmbeddedFile,

    /// <summary>A rich-media, movie, sound, screen, 3D, or file-attachment annotation.</summary>
    RichMedia
}
