namespace OfficeIMO.Pdf;

/// <summary>Selectable document-content categories understood by the PDF sanitization engine.</summary>
[System.Flags]
public enum PdfSanitizationContentKind {
    /// <summary>Do not select any document-content category.</summary>
    None = 0,
    /// <summary>User-authored Info fields and the catalog XMP metadata packet.</summary>
    UserMetadata = 1 << 0,
    /// <summary>Embedded and associated file attachments.</summary>
    EmbeddedFiles = 1 << 1,
    /// <summary>JavaScript, URI, launch, submit, remote-navigation, and other selected actions.</summary>
    Actions = 1 << 2,
    /// <summary>Comment and markup annotations, excluding Link, Widget, and FileAttachment annotations.</summary>
    CommentsAndMarkup = 1 << 3,
    /// <summary>The document outline/bookmark tree.</summary>
    Bookmarks = 1 << 4,
    /// <summary>Optional-content layer definitions and associations.</summary>
    OptionalContent = 1 << 5,
    /// <summary>Every selectable before-sharing category.</summary>
    All = UserMetadata | EmbeddedFiles | Actions | CommentsAndMarkup | Bookmarks | OptionalContent
}
