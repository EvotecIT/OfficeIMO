namespace OfficeIMO.Pdf;

/// <summary>How a text element follows the one before it in extraction order.</summary>
public enum PdfTextBreak {
    /// <summary>The element continues the previous one: their glyphs touch along one baseline.</summary>
    None = 0,
    /// <summary>The element is further along the same line after a gap, as between separately positioned words.</summary>
    Space = 1,
    /// <summary>The element starts another line.</summary>
    Line = 2
}
