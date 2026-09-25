namespace OfficeIMO.Pdf;

/// <summary>Source features omitted when pages become content on new sheets.</summary>
[Flags]
public enum PdfImpositionSourceFeatureLoss {
    /// <summary>No known source feature was omitted.</summary>
    None = 0,
    /// <summary>Page annotations and their appearances were not transferred.</summary>
    Annotations = 1,
    /// <summary>AcroForm fields and widgets were not transferred.</summary>
    Forms = 2,
    /// <summary>Document structure tags were not transferred.</summary>
    StructureTags = 4,
    /// <summary>Source signatures were removed in an explicit unsigned derivative.</summary>
    Signatures = 8,
    /// <summary>Document-level embedded files and associated attachments were not transferred.</summary>
    EmbeddedFiles = 16,
    /// <summary>Document-level output intent profiles were not transferred.</summary>
    OutputIntents = 32,
    /// <summary>Document outline and bookmark navigation was not transferred.</summary>
    Outlines = 64,
    /// <summary>Document page numbering labels were not transferred.</summary>
    PageLabels = 128,
    /// <summary>Named destinations were not transferred to the new sheet pages.</summary>
    NamedDestinations = 256,
    /// <summary>Catalog view settings, actions, name trees, or URI settings were not transferred.</summary>
    CatalogFeatures = 512,
    /// <summary>Info-dictionary or XMP metadata was not transferred.</summary>
    DocumentMetadata = 1024,
    /// <summary>Selected-page actions, metadata, or presentation settings were not transferred.</summary>
    PageFeatures = 2048,
    /// <summary>Source encryption was not transferred to the derivative sheets.</summary>
    Encryption = 4096
}
