namespace OfficeIMO.Pdf;

/// <summary>Expansion preference for one generated PDF outline entry.</summary>
public enum PdfOutlineState {
    /// <summary>Uses the document-wide outline expansion level.</summary>
    Default,
    /// <summary>Shows the entry's descendants when the document is opened.</summary>
    Open,
    /// <summary>Collapses the entry's descendants when the document is opened.</summary>
    Closed
}
