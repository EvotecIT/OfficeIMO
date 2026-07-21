namespace OfficeIMO.Pdf;

/// <summary>Permission or authorization checks relevant to a planned PDF mutation.</summary>
public enum PdfMutationPermissionCheck {
    /// <summary>The document must be readable with the supplied options.</summary>
    ReadDocument,

    /// <summary>The document security permissions must allow general modification.</summary>
    ModifyDocument,

    /// <summary>The document security permissions must allow document assembly.</summary>
    AssembleDocument,

    /// <summary>The document security permissions must allow annotation changes.</summary>
    ModifyAnnotations,

    /// <summary>The document security permissions must allow form filling.</summary>
    FillForms,

    /// <summary>Certification-signature DocMDP rules must allow the requested change.</summary>
    DocMdp,

    /// <summary>Signature-field FieldMDP rules must allow the requested target fields.</summary>
    FieldMdp,

    /// <summary>The operation requires a valid append-only revision chain.</summary>
    AppendRevision,

    /// <summary>The operation may only fill a prepared signature contents reservation.</summary>
    FillSignatureContentsReservation,

    /// <summary>The operation requires owner-level authorization for encryption changes.</summary>
    OwnerAuthorization,

    /// <summary>The document security permissions must allow content copying or extraction.</summary>
    CopyContents
}
