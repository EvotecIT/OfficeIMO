namespace OfficeIMO.Pdf;

/// <summary>
/// Describes how a table cell border line is drawn.
/// </summary>
public enum PdfCellBorderLineStyle {
    /// <summary>Draw a single line.</summary>
    Standard,

    /// <summary>Draw two parallel strokes. Internal boundaries share centred tracks; outer borders retain inward tracks. Cell content keeps clear of the painted extent.</summary>
    TwoLine
}
