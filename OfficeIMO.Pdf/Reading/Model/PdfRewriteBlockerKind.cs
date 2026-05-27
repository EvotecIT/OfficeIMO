namespace OfficeIMO.Pdf;

/// <summary>
/// Categories of PDF features that OfficeIMO.Pdf will not rewrite until preservation support exists.
/// </summary>
public enum PdfRewriteBlockerKind {
    /// <summary>Encrypted PDFs cannot be read or rewritten yet.</summary>
    Encryption,

    /// <summary>Digital signature markers are present.</summary>
    Signatures,

    /// <summary>AcroForm or form field markers are present.</summary>
    Forms,

    /// <summary>Outline or bookmark markers are present.</summary>
    Outlines,

    /// <summary>Catalog page mode or layout markers are present.</summary>
    CatalogViewSettings,

    /// <summary>Page label markers are present.</summary>
    PageLabels,

    /// <summary>Unsupported catalog name-tree markers are present.</summary>
    CatalogNameTrees,

    /// <summary>Named destination markers are present.</summary>
    NamedDestinations,

    /// <summary>Document open action markers are present.</summary>
    OpenActions,

    /// <summary>Viewer preference markers are present.</summary>
    ViewerPreferences,

    /// <summary>Tagged PDF structure markers are present.</summary>
    TaggedContent,

    /// <summary>XMP metadata stream markers are present.</summary>
    XmpMetadata,

    /// <summary>Unsupported catalog URI dictionary markers are present.</summary>
    CatalogUri,

    /// <summary>Output intent markers are present.</summary>
    OutputIntents,

    /// <summary>Embedded file markers are present.</summary>
    EmbeddedFiles,

    /// <summary>Optional content/layer markers are present.</summary>
    OptionalContent,

    /// <summary>Active content markers such as JavaScript actions are present.</summary>
    ActiveContent,

    /// <summary>Rewrite object graph contains missing or wrong-generation indirect references.</summary>
    InvalidObjectReferences
}
