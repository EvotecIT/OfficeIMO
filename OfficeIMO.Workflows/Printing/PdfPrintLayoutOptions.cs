namespace OfficeIMO.Workflows;

/// <summary>Position of a source page within its print-sheet slot, including cropped overflow.</summary>
public enum PdfPrintAlignment {
    /// <summary>Centered horizontally and vertically.</summary>
    Center,
    /// <summary>Top left.</summary>
    TopLeft,
    /// <summary>Top center.</summary>
    Top,
    /// <summary>Top right.</summary>
    TopRight,
    /// <summary>Left center.</summary>
    Left,
    /// <summary>Right center.</summary>
    Right,
    /// <summary>Bottom left.</summary>
    BottomLeft,
    /// <summary>Bottom center.</summary>
    Bottom,
    /// <summary>Bottom right.</summary>
    BottomRight
}

/// <summary>Filter applied to selected original source-page numbers before sheet assembly.</summary>
public enum PdfPrintPageSubset {
    /// <summary>Retain every selected source page.</summary>
    All,
    /// <summary>Retain selected odd-numbered source pages.</summary>
    Odd,
    /// <summary>Retain selected even-numbered source pages.</summary>
    Even
}

/// <summary>Color treatment of prepared print pixels, independent of driver color capabilities.</summary>
public enum PdfPrintColorMode {
    /// <summary>Retain source colors; final device output depends on the printer.</summary>
    Color,
    /// <summary>Convert reviewed and delivered sheet pixels to grayscale.</summary>
    Grayscale
}
