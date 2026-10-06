namespace OfficeIMO.Pdf;

/// <summary>Units and expansion behavior for paragraph line spacing.</summary>
public enum PdfLineSpacingRule {
    /// <summary>Line advance is proportional to font size.</summary>
    Multiple,
    /// <summary>Line advance is fixed in points.</summary>
    Exact,
    /// <summary>Line advance is a point minimum and expands for larger text or inline elements.</summary>
    AtLeast
}
