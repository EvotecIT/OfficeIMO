namespace OfficeIMO.Pdf;

/// <summary>Semantic label style for a tagged list, independent from its painted marker.</summary>
public enum PdfListNumbering {
    /// <summary>No prescribed numbering scheme; labels may contain arbitrary text.</summary>
    None,
    /// <summary>Solid circular bullets.</summary>
    Disc,
    /// <summary>Hollow circular bullets.</summary>
    Circle,
    /// <summary>Square bullets.</summary>
    Square,
    /// <summary>Decimal numbers.</summary>
    DecimalNumbers,
    /// <summary>Uppercase Roman numbers.</summary>
    UpperRoman,
    /// <summary>Lowercase Roman numbers.</summary>
    LowerRoman,
    /// <summary>Uppercase alphabetic labels.</summary>
    UpperAlpha,
    /// <summary>Lowercase alphabetic labels.</summary>
    LowerAlpha,
    /// <summary>Other ordered labels. Requires PDF 2.0.</summary>
    Ordered,
    /// <summary>Other unordered labels. Requires PDF 2.0.</summary>
    Unordered,
    /// <summary>Description-list labels. Requires PDF 2.0.</summary>
    Description
}
