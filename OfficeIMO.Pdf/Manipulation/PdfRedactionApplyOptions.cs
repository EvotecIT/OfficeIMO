namespace OfficeIMO.Pdf;

/// <summary>Additional document-level residue removed by an explicit redaction policy.</summary>
[Flags]
public enum PdfRedactionCleanupScope {
    /// <summary>Only intersecting page content and annotations are changed.</summary>
    None = 0,
    /// <summary>Clear Info metadata and remove catalog/page XMP streams.</summary>
    Metadata = 1,
    /// <summary>Remove embedded/associated files and file-attachment annotations.</summary>
    Attachments = 2,
    /// <summary>Remove the tagged structure tree and page structure-parent references.</summary>
    StructureTree = 4,
    /// <summary>Remove alternate/actual text and form alternate-name entries.</summary>
    AlternateText = 8,
    /// <summary>Remove optional-content catalog, resource, and object references.</summary>
    OptionalContent = 16,
    /// <summary>Apply every document-level cleanup category.</summary>
    All = Metadata | Attachments | StructureTree | AlternateText | OptionalContent
}

/// <summary>Fail-closed behavior when an intersecting image cannot be safely rewritten at pixel level.</summary>
public enum PdfRedactionUnsupportedImagePolicy {
    /// <summary>Reject the operation.</summary>
    FailClosed,
    /// <summary>Remove the entire intersecting placement and prune its resource when unused.</summary>
    RemoveWholePlacement,
    /// <summary>Keep the image bytes and paint a visual-only overlay.</summary>
    VisualOverlay
}

/// <summary>
/// Options controlling how planned PDF redaction areas are applied.
/// </summary>
public sealed class PdfRedactionApplyOptions {
    /// <summary>
    /// Fill color used for the visible redaction mark. Defaults to black.
    /// </summary>
    public PdfColor FillColor { get; set; } = PdfColor.Black;

    /// <summary>
    /// When true, redaction areas are painted even when no text or annotation match is found in the area.
    /// </summary>
    public bool PaintUnmatchedAreas { get; set; } = true;

    /// <summary>
    /// When true, redaction areas that intersect image placements are allowed to be painted as visual overlays even though image pixels and resources are not rewritten.
    /// </summary>
    public bool AllowImagePlacementOverlays { get; set; }

    /// <summary>Explicit fallback for JPEG, transformed, masked, indexed, or otherwise unsupported partial-image rewrites.</summary>
    public PdfRedactionUnsupportedImagePolicy UnsupportedImagePolicy { get; set; } = PdfRedactionUnsupportedImagePolicy.FailClosed;

    /// <summary>Explicit document-level residue cleanup. Defaults to intersecting content only.</summary>
    public PdfRedactionCleanupScope CleanupScope { get; set; }

    /// <summary>Remove intersecting painted vector paths and rectangles from page content before applying the redaction mark.</summary>
    public bool RemoveIntersectingPaths { get; set; } = true;

    /// <summary>Optional decoder used for image codecs, such as JPEG, that are intentionally not dependencies of the PDF core.</summary>
    public IPdfRedactionImageDecoder? ImageDecoder { get; set; }

    /// <summary>Maximum RGBA bytes accepted from an optional image decoder. Defaults to 256 MiB.</summary>
    public int MaximumDecodedImageBytes { get; set; } = 256 * 1024 * 1024;
}
