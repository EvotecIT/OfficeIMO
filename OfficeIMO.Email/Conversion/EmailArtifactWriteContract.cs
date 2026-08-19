namespace OfficeIMO.Email;

/// <summary>Identifies which bytes were selected for an email artifact write.</summary>
public enum EmailArtifactSourceSelection {
    /// <summary>No artifact was produced because policy blocked the operation.</summary>
    None = 0,
    /// <summary>The unchanged source artifact was emitted byte for byte.</summary>
    PreservedSource = 1,
    /// <summary>The artifact was regenerated from the structured <see cref="EmailDocument"/> model.</summary>
    Regenerated = 2
}

/// <summary>Describes the final disposition of known semantic loss.</summary>
public enum EmailConversionLossDisposition {
    /// <summary>No known semantic loss was required by the completed operation.</summary>
    None = 0,
    /// <summary>Known semantic loss was explicitly accepted by the configured policy.</summary>
    Accepted = 1,
    /// <summary>Known semantic loss blocked artifact creation.</summary>
    Blocked = 2
}

/// <summary>Describes how attachment content sources participated in one write.</summary>
public enum EmailAttachmentContentLifetime {
    /// <summary>Attachment sources were not opened because preserved source bytes were reused or writing was blocked.</summary>
    NotAccessed = 0,
    /// <summary>
    /// Attachment streams were opened, consumed, and disposed inside the write operation. The result does not retain
    /// store handles, streams, or attachment payloads.
    /// </summary>
    OperationScoped = 1
}

/// <summary>Stable diagnostic identifiers emitted by artifact preservation and conversion.</summary>
public static class EmailArtifactDiagnosticCodes {
    /// <summary>The preserved source was not selected because the structured model changed.</summary>
    public const string PreservedSourceModelChanged = "EMAIL_RAW_SOURCE_SKIPPED_MODEL_CHANGED";
    /// <summary>A protected wrapper would need an invalidating rewrite.</summary>
    public const string ProtectedContentRewrite = "EMAIL_PROTECTED_CONTENT_REWRITE";
}
