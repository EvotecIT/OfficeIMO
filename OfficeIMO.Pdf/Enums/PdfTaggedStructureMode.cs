namespace OfficeIMO.Pdf;

/// <summary>
/// Controls generated tagged-PDF groundwork emitted into the document catalog.
/// </summary>
public enum PdfTaggedStructureMode {
    /// <summary>Do not emit tagged-PDF catalog markers.</summary>
    None = 0,
    /// <summary>
    /// Emit tagged-PDF catalog markers. Generated pages emit structure-order tab hints; paragraphs, headings, list labels and bodies nested below generated L/LI containers, table captions, table cell slices, generated Table/TR containers, table header scope attributes, table span attributes, and image figures with alternate text can also receive limited MCID and parent-tree references.
    /// This is readiness groundwork and does not provide complete tagged page content.
    /// </summary>
    CatalogMarkers = 1
}
