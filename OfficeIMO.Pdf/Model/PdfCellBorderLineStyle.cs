namespace OfficeIMO.Pdf;

/// <summary>
/// Describes how a table cell border line is drawn.
/// </summary>
public enum PdfCellBorderLineStyle {
    /// <summary>Draw a single line.</summary>
    Standard,

    /// <summary>Draw two parallel strokes centred on the cell boundary. Cell content keeps clear of their inward painted extent.</summary>
    TwoLine
}
