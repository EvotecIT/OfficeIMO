namespace OfficeIMO.Pdf;

/// <summary>
/// Describes the expected layout impact of a configured embedded font-family substitution.
/// </summary>
public enum PdfFontFamilySubstitutionImpact {
    /// <summary>
    /// The caller accepts the substitute as layout-compatible for the intended document profile.
    /// The generated appearance can still differ from the source font.
    /// </summary>
    Compatible,

    /// <summary>
    /// The substitute can change glyph appearance, text fit, wrapping, or pagination.
    /// </summary>
    LayoutSensitive
}
