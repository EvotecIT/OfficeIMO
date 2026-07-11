namespace OfficeIMO.Pdf;

/// <summary>Evidence required after a planned PDF mutation.</summary>
public enum PdfMutationProof {
    /// <summary>The produced artifact must be readable by the OfficeIMO.Pdf parser.</summary>
    ReadableOutput,

    /// <summary>The rewrite-preservation matrix must accept the requested preservation policy.</summary>
    RewritePreservation,

    /// <summary>The original PDF bytes must remain an exact prefix of the output.</summary>
    BytePrefixPreservation,

    /// <summary>The incremental revision and cross-reference chain must remain valid.</summary>
    RevisionChain,

    /// <summary>Requested metadata must be present after readback.</summary>
    MetadataReadback,

    /// <summary>Requested form values and widget states must be present after readback.</summary>
    FormFieldReadback,

    /// <summary>Page count, ordering, geometry, and destinations must match the requested edit.</summary>
    PageStructureReadback,

    /// <summary>Changed page or appearance content must be rendered for visual proof.</summary>
    VisualRendering,

    /// <summary>Annotation objects, geometry, actions, and appearances must match after readback.</summary>
    AnnotationReadback,

    /// <summary>Attachment names, relationships, metadata, and payload hashes must match after readback.</summary>
    AttachmentReadback,

    /// <summary>Signature byte ranges and revision coverage must remain structurally valid.</summary>
    SignatureByteRanges,

    /// <summary>Only the prepared signature /Contents reservation may change and the file length must remain unchanged.</summary>
    ReservedSignatureContentsPatch,

    /// <summary>DocMDP and FieldMDP permission results must remain valid for the applied change.</summary>
    SignaturePermissions,

    /// <summary>Encryption, password, permission, metadata, and content round trips must pass.</summary>
    EncryptionRoundTrip,

    /// <summary>Removed content must not remain extractable or recoverable from decoded objects and streams.</summary>
    RedactionResidue,

    /// <summary>DSS/VRI certificate, OCSP, CRL, and timestamp references must match after readback.</summary>
    LongTermValidationReadback
}
