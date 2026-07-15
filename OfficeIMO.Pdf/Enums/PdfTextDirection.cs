namespace OfficeIMO.Pdf;

/// <summary>Base direction supplied to a host-owned PDF text shaping provider.</summary>
public enum PdfTextDirection {
    /// <summary>No strong directional character was found; the provider may resolve direction from its own context.</summary>
    Auto = 0,
    /// <summary>The first strong character establishes a left-to-right base direction.</summary>
    LeftToRight = 1,
    /// <summary>The first strong character establishes a right-to-left base direction.</summary>
    RightToLeft = 2
}
