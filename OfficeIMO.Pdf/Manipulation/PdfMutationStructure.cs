namespace OfficeIMO.Pdf;

/// <summary>PDF structures that a planned mutation can affect.</summary>
public enum PdfMutationStructure {
    /// <summary>Trailer and indirect-object reachability state.</summary>
    ObjectGraph,

    /// <summary>Document information dictionary.</summary>
    InfoDictionary,

    /// <summary>XMP metadata stream and schemas.</summary>
    XmpMetadata,

    /// <summary>Catalog dictionary and catalog-owned structures.</summary>
    Catalog,

    /// <summary>Outlines, destinations, page labels, and related navigation.</summary>
    Navigation,

    /// <summary>Page tree membership and ordering.</summary>
    PageTree,

    /// <summary>Page content streams.</summary>
    PageContent,

    /// <summary>Page resource dictionaries and referenced resources.</summary>
    PageResources,

    /// <summary>Page annotations, including widget annotations.</summary>
    Annotations,

    /// <summary>AcroForm fields and form dictionaries.</summary>
    AcroForm,

    /// <summary>Form or annotation appearance streams.</summary>
    AppearanceStreams,

    /// <summary>Digital signature dictionaries, byte ranges, and signature fields.</summary>
    Signatures,

    /// <summary>Embedded files, associated files, and their name trees.</summary>
    Attachments,

    /// <summary>Standard or public-key security dictionaries and encrypted objects.</summary>
    Encryption,

    /// <summary>Tagged-PDF structure and marked-content references.</summary>
    TaggedContent
}
