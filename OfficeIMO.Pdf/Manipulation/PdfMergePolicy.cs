namespace OfficeIMO.Pdf;

/// <summary>Controls how one document-level structure participates in a PDF merge.</summary>
public enum PdfMergeStructureMode {
    /// <summary>Retains the structure from the primary source and ignores incoming copies.</summary>
    KeepPrimary,
    /// <summary>Combines compatible values from every source using structure-specific rules.</summary>
    Combine,
    /// <summary>Removes the structure from the merged artifact.</summary>
    Drop,
    /// <summary>Rejects the merge when a non-primary source also contains the structure.</summary>
    RejectIncoming
}

/// <summary>Controls collisions between named items imported from multiple PDFs.</summary>
public enum PdfMergeCollisionMode {
    /// <summary>Keeps the first item and drops later items with the same name.</summary>
    KeepFirst,
    /// <summary>Renames later items deterministically by appending a source and sequence suffix.</summary>
    RenameIncoming,
    /// <summary>Rejects the merge when two items use the same name.</summary>
    Reject
}

/// <summary>
/// Shared document-level policy for first-party PDF merge and page-import workflows.
/// Defaults preserve the primary source, matching the historical OfficeIMO.Pdf merge behavior.
/// </summary>
public sealed class PdfMergePolicy {
    /// <summary>Metadata scalar handling. Combine fills missing primary values from later sources.</summary>
    public PdfMergeStructureMode Metadata { get; set; } = PdfMergeStructureMode.KeepPrimary;
    /// <summary>Outline/bookmark tree handling. Combine appends source roots in source order.</summary>
    public PdfMergeStructureMode Outlines { get; set; } = PdfMergeStructureMode.KeepPrimary;
    /// <summary>Named destination handling.</summary>
    public PdfMergeStructureMode NamedDestinations { get; set; } = PdfMergeStructureMode.KeepPrimary;
    /// <summary>Page-label number-tree handling.</summary>
    public PdfMergeStructureMode PageLabels { get; set; } = PdfMergeStructureMode.KeepPrimary;
    /// <summary>AcroForm field-tree handling.</summary>
    public PdfMergeStructureMode Forms { get; set; } = PdfMergeStructureMode.KeepPrimary;
    /// <summary>Embedded and associated-file handling.</summary>
    public PdfMergeStructureMode Attachments { get; set; } = PdfMergeStructureMode.KeepPrimary;
    /// <summary>Viewer preference and initial-view handling.</summary>
    public PdfMergeStructureMode ViewerPreferences { get; set; } = PdfMergeStructureMode.KeepPrimary;
    /// <summary>Remaining catalog state such as output intents and optional-content properties.</summary>
    public PdfMergeStructureMode CatalogState { get; set; } = PdfMergeStructureMode.KeepPrimary;
    /// <summary>Collision behavior for named destinations.</summary>
    public PdfMergeCollisionMode NamedDestinationCollisions { get; set; } = PdfMergeCollisionMode.RenameIncoming;
    /// <summary>Collision behavior for attachment file names.</summary>
    public PdfMergeCollisionMode AttachmentCollisions { get; set; } = PdfMergeCollisionMode.RenameIncoming;
}
