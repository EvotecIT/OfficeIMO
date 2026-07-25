namespace OfficeIMO.Pdf;

/// <summary>
/// Maps a source document font family to an embedded PDF family with an explicit layout-impact contract.
/// </summary>
public sealed class PdfFontFamilySubstitution {
    /// <summary>
    /// Creates a font-family substitution.
    /// </summary>
    /// <param name="sourceFontFamily">Source family declared by the input document.</param>
    /// <param name="targetFontFamily">Registered embedded family used in generated PDF text.</param>
    /// <param name="impact">Expected layout impact of the substitution.</param>
    public PdfFontFamilySubstitution(
        string sourceFontFamily,
        string targetFontFamily,
        PdfFontFamilySubstitutionImpact impact) {
        if (string.IsNullOrWhiteSpace(sourceFontFamily)) {
            throw new ArgumentException("Source font family cannot be empty.", nameof(sourceFontFamily));
        }
        if (string.IsNullOrWhiteSpace(targetFontFamily)) {
            throw new ArgumentException("Target font family cannot be empty.", nameof(targetFontFamily));
        }

        SourceFontFamily = sourceFontFamily.Trim();
        TargetFontFamily = targetFontFamily.Trim();
        if (string.Equals(SourceFontFamily, TargetFontFamily, StringComparison.OrdinalIgnoreCase)) {
            throw new ArgumentException("Source and target font families must be different.", nameof(targetFontFamily));
        }
        if (impact != PdfFontFamilySubstitutionImpact.Compatible &&
            impact != PdfFontFamilySubstitutionImpact.LayoutSensitive) {
            throw new ArgumentOutOfRangeException(nameof(impact), impact, "Unknown font substitution impact.");
        }

        Impact = impact;
    }

    /// <summary>Source family declared by the input document.</summary>
    public string SourceFontFamily { get; }

    /// <summary>Registered embedded family used in generated PDF text.</summary>
    public string TargetFontFamily { get; }

    /// <summary>Expected layout impact of the substitution.</summary>
    public PdfFontFamilySubstitutionImpact Impact { get; }

    internal PdfFontFamilySubstitution Clone() =>
        new(SourceFontFamily, TargetFontFamily, Impact);
}
