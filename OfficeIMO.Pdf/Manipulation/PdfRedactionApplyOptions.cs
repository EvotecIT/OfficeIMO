namespace OfficeIMO.Pdf;

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
}
