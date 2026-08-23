namespace OfficeIMO.Pdf;

/// <summary>
/// Controls black generation when RGB vector and text colors are converted to a CMYK print condition.
/// </summary>
public enum PdfBlackPreservationMode {
    /// <summary>Use the ICC profile result without a black-preservation override.</summary>
    None,
    /// <summary>Preserve pure RGB black as 100 percent K.</summary>
    PureBlack,
    /// <summary>Preserve the complete neutral RGB axis as K-only output.</summary>
    NeutralAxis
}
