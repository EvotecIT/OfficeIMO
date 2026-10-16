namespace OfficeIMO.Pdf;

/// <summary>
/// Describes the expected layout impact of a configured embedded font-family substitution.
/// </summary>
public enum PdfFontFamilySubstitutionImpact {
    /// <summary>
    /// The caller or conversion policy selects the substitute with compatibility intent for the document profile.
    /// This classification does not verify exact metrics for the selected font versions, including
    /// vertical metrics, wrapping, or pagination. The generated appearance can still differ from the source font.
    /// </summary>
    Compatible,

    /// <summary>
    /// The substitute can change glyph appearance, text fit, wrapping, or pagination.
    /// </summary>
    LayoutSensitive
}
