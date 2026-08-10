namespace OfficeIMO.OpenDocument;

/// <summary>Horizontal alignment for an ODP paragraph.</summary>
public enum OdpParagraphAlignment {
    /// <summary>Aligns content to the logical start edge.</summary>
    Start = 0,
    /// <summary>Centers content.</summary>
    Center = 1,
    /// <summary>Aligns content to the logical end edge.</summary>
    End = 2,
    /// <summary>Justifies content on both edges.</summary>
    Justify = 3,
    /// <summary>Aligns content to the physical left edge.</summary>
    Left = 4,
    /// <summary>Aligns content to the physical right edge.</summary>
    Right = 5
}
