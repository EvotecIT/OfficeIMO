namespace OfficeIMO.Pdf;

/// <summary>Shared existing-document capability families evaluated by the mutation planner.</summary>
public enum PdfMutationCapabilityKind {
    /// <summary>Page membership, ordering, geometry, and inherited page-tree state.</summary>
    PageTreeChanges,

    /// <summary>Page content streams, resources, marked content, and visible drawing state.</summary>
    ContentChanges,

    /// <summary>Catalog-owned navigation, names, labels, layers, preferences, and related state.</summary>
    CatalogChanges,

    /// <summary>AcroForm fields, widgets, appearances, and calculation state.</summary>
    FormChanges,

    /// <summary>Page annotations, actions, geometry, replies, and appearances.</summary>
    AnnotationChanges,

    /// <summary>Info-dictionary and XMP metadata.</summary>
    MetadataChanges,

    /// <summary>Embedded and associated files plus catalog attachment indexes.</summary>
    AttachmentChanges,

    /// <summary>Security dictionaries, passwords, permissions, and encrypted objects.</summary>
    EncryptionChanges,

    /// <summary>Signature fields, byte ranges, permissions, timestamps, and validation evidence.</summary>
    SignatureChanges
}
