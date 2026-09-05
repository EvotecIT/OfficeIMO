namespace OfficeIMO.Pdf;

/// <summary>Block-level semantic roles for generated tagged PDF content.</summary>
public enum PdfSemanticRole {
    /// <summary>A major document part.</summary>
    Part,
    /// <summary>A self-contained article.</summary>
    Article,
    /// <summary>A document section.</summary>
    Section,
    /// <summary>A generic structural division.</summary>
    Division,
    /// <summary>A quotation containing one or more flow blocks.</summary>
    BlockQuote,
    /// <summary>A caption associated with nearby content.</summary>
    Caption,
    /// <summary>A figure whose purpose is described by alternate text.</summary>
    Figure,
    /// <summary>A form region containing one or more interactive fields.</summary>
    Form
}
