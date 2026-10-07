namespace OfficeIMO.Epub;

/// <summary>The semantic partition to which an EPUB content document belongs.</summary>
public enum EpubDocumentMatter {
    /// <summary>Preliminary material before the main content.</summary>
    FrontMatter,
    /// <summary>The main content of the publication.</summary>
    BodyMatter,
    /// <summary>Supplementary material after the main content.</summary>
    BackMatter
}
