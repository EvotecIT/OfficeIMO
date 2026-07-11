namespace OfficeIMO.Pdf;

/// <summary>
/// Semantic structure role for tagged text placed on an absolute page canvas.
/// </summary>
public enum PdfCanvasTextStructureRole {
    /// <summary>Paragraph content.</summary>
    Paragraph = 0,
    /// <summary>Level-one heading.</summary>
    Heading1 = 1,
    /// <summary>Level-two heading.</summary>
    Heading2 = 2,
    /// <summary>Level-three heading.</summary>
    Heading3 = 3,
    /// <summary>Level-four heading.</summary>
    Heading4 = 4,
    /// <summary>Level-five heading.</summary>
    Heading5 = 5,
    /// <summary>Level-six heading.</summary>
    Heading6 = 6,
    /// <summary>Generic inline span.</summary>
    Span = 7
}
