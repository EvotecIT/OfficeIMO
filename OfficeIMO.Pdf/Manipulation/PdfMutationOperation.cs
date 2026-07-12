namespace OfficeIMO.Pdf;

/// <summary>Existing-document mutation families understood by <see cref="PdfMutationPlanner"/>.</summary>
public enum PdfMutationOperation {
    /// <summary>Update document information metadata.</summary>
    UpdateMetadata,

    /// <summary>Update supported AcroForm field values and appearances.</summary>
    FillFormFields,

    /// <summary>Flatten supported AcroForm widgets into page content.</summary>
    FlattenFormFields,

    /// <summary>Update and then flatten supported AcroForm fields.</summary>
    FillAndFlattenFormFields,

    /// <summary>Append an external-signature placeholder revision.</summary>
    PrepareExternalSignature,

    /// <summary>Fill a prepared all-zero external-signature contents reservation without changing file length.</summary>
    FinalizeExternalSignature,

    /// <summary>Create one or more independent output PDFs from selected source pages.</summary>
    ExtractPages,

    /// <summary>Change page membership or order, including delete, move, duplicate, merge, or import.</summary>
    ModifyPageTree,

    /// <summary>Merge complete documents with explicit catalog, navigation, form, and attachment policies.</summary>
    MergeDocuments,

    /// <summary>Change page content streams or resources, including stamps and watermarks.</summary>
    ModifyPageContent,

    /// <summary>Change catalog-level navigation, viewer, layer, or related document state.</summary>
    ModifyCatalog,

    /// <summary>Create, update, remove, or flatten annotations.</summary>
    ModifyAnnotations,

    /// <summary>Create, replace, rename, or remove embedded or associated files.</summary>
    ModifyAttachments,

    /// <summary>Change document encryption, passwords, or permissions.</summary>
    ChangeEncryption,

    /// <summary>Apply lossless document optimization.</summary>
    Optimize,

    /// <summary>Apply destructive content redaction.</summary>
    Redact,

    /// <summary>Append DSS/VRI evidence for cryptographically verified signatures.</summary>
    EnrichLongTermValidation,

    /// <summary>Synchronize document information metadata with catalog XMP metadata.</summary>
    SynchronizeMetadata,

    /// <summary>Remove or quarantine active content and embedded payloads according to an explicit policy.</summary>
    Sanitize
}
