namespace OfficeIMO.Pdf;

/// <summary>
/// Options used when updating or flattening parsed AcroForm fields.
/// </summary>
public sealed class PdfFormFillerOptions {
    private PdfEmbeddedFontFamily? _appearanceFontFamily;
    private PdfEmbeddedFontFallbackSet? _appearanceFontFallbacks;
    private PdfConversionReport? _diagnosticsReport;
    private string _diagnosticsConverter = "OfficeIMO.Pdf";

    /// <summary>Maximum XFDF source length accepted before XML materialization.</summary>
    public int MaxXfdfDocumentCharacters { get; set; } = PdfFormDataSet.DefaultMaxXfdfDocumentCharacters;

    /// <summary>Maximum number of fields accepted from one XFDF document.</summary>
    public int MaxXfdfFields { get; set; } = PdfFormDataSet.DefaultMaxXfdfFields;

    /// <summary>Maximum aggregate length of values accepted from one XFDF document.</summary>
    public int MaxXfdfValueCharacters { get; set; } = PdfFormDataSet.DefaultMaxXfdfValueCharacters;

    /// <summary>
    /// When true, keeps the AcroForm /NeedAppearances flag set after filling fields for legacy viewers that ignore normal appearance streams.
    /// </summary>
    /// <remarks>The default is false because OfficeIMO synthesizes normal widget appearances during filling.</remarks>
    public bool KeepNeedAppearances { get; set; }

    /// <summary>
    /// Optional TrueType font family used to synthesize embedded Unicode text appearances when a parsed PDF does not already provide a reusable embedded appearance font.
    /// </summary>
    public PdfEmbeddedFontFamily? AppearanceFontFamily {
        get => _appearanceFontFamily?.Clone();
        set => _appearanceFontFamily = value?.Clone();
    }

    /// <summary>
    /// Optional prioritized fallback set used to synthesize embedded Unicode text appearances when the preferred appearance font cannot cover the field value.
    /// </summary>
    public PdfEmbeddedFontFallbackSet? AppearanceFontFallbacks {
        get => _appearanceFontFallbacks?.Clone();
        set => _appearanceFontFallbacks = value?.Clone();
    }

    internal PdfEmbeddedFontFamily? AppearanceFontFamilySnapshot => _appearanceFontFamily?.Clone();

    internal bool HasAppearanceFontFamily => _appearanceFontFamily != null;

    internal PdfEmbeddedFontFallbackSet? AppearanceFontFallbacksSnapshot => _appearanceFontFallbacks?.Clone();

    internal bool HasAppearanceFontFallbacks => _appearanceFontFallbacks != null;

    internal void AddTextDiagnostics(IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics) {
        if (_diagnosticsReport == null || diagnostics.Count == 0) {
            return;
        }

        _diagnosticsReport.AddTextDiagnostics(diagnostics, _diagnosticsConverter);
    }

    internal void AddTextFallbackPlanDiagnostics(PdfTextFallbackPlan plan) {
        _diagnosticsReport?.AddTextFallbackPlanDiagnostics(plan, _diagnosticsConverter);
    }

    internal void AddTextShapingDiagnostics(IReadOnlyList<PdfTextShapingDiagnostic> diagnostics) {
        if (_diagnosticsReport == null || diagnostics.Count == 0) {
            return;
        }

        _diagnosticsReport.AddTextShapingDiagnostics(diagnostics, _diagnosticsConverter);
    }

    internal void AddFontDiagnostics(IReadOnlyList<PdfFontEmbeddingDiagnostic> diagnostics) {
        if (_diagnosticsReport == null || diagnostics.Count == 0) {
            return;
        }

        _diagnosticsReport.AddFontDiagnostics(diagnostics, _diagnosticsConverter);
    }

    /// <summary>
    /// Registers a TrueType font family for synthesized text appearances.
    /// </summary>
    /// <param name="fontFamily">Font family to embed into synthesized appearance streams.</param>
    /// <returns>The current options instance for fluent configuration.</returns>
    public PdfFormFillerOptions UseAppearanceFontFamily(PdfEmbeddedFontFamily fontFamily) {
        Guard.NotNull(fontFamily, nameof(fontFamily));
        _appearanceFontFamily = fontFamily.Clone();
        return this;
    }

    /// <summary>
    /// Registers prioritized fallback fonts for synthesized text appearances.
    /// </summary>
    /// <param name="fallbackSet">Fallback set used when the preferred appearance font cannot cover the field value.</param>
    /// <returns>The current options instance for fluent configuration.</returns>
    public PdfFormFillerOptions UseAppearanceFontFallbacks(PdfEmbeddedFontFallbackSet fallbackSet) {
        Guard.NotNull(fallbackSet, nameof(fallbackSet));
        _appearanceFontFallbacks = fallbackSet.Clone();
        return this;
    }

    /// <summary>
    /// Records structured text diagnostics encountered while synthesizing form appearances.
    /// </summary>
    /// <param name="report">Mutable conversion report that receives missing-glyph, shaping, and embedded appearance-font diagnostics.</param>
    /// <param name="converter">Converter or adapter name to place on recorded warnings.</param>
    /// <returns>The current options instance for fluent configuration.</returns>
    /// <remarks>Diagnostics are reported in addition to the existing fail-closed exceptions.</remarks>
    public PdfFormFillerOptions ReportDiagnosticsTo(PdfConversionReport report, string converter = "OfficeIMO.Pdf") {
        Guard.NotNull(report, nameof(report));
        _diagnosticsReport = report;
        _diagnosticsConverter = string.IsNullOrWhiteSpace(converter) ? "OfficeIMO.Pdf" : converter;
        return this;
    }

    /// <summary>
    /// Registers a single regular TrueType or OpenType/CFF font face for synthesized text appearances.
    /// </summary>
    /// <param name="familyName">PDF font family name to expose in synthesized resources.</param>
    /// <param name="regular">Regular TrueType or OpenType/CFF font bytes.</param>
    /// <returns>The current options instance for fluent configuration.</returns>
    public PdfFormFillerOptions UseAppearanceFont(string familyName, byte[] regular) {
        _appearanceFontFamily = new PdfEmbeddedFontFamily(familyName, regular);
        return this;
    }

    /// <summary>
    /// Registers a single regular TrueType or OpenType/CFF font file for synthesized text appearances.
    /// </summary>
    /// <param name="familyName">PDF font family name to expose in synthesized resources.</param>
    /// <param name="regularPath">Path to a regular TrueType or OpenType/CFF font file.</param>
    /// <returns>The current options instance for fluent configuration.</returns>
    public PdfFormFillerOptions UseAppearanceFontFile(string familyName, string regularPath) {
        Guard.NotNullOrWhiteSpace(regularPath, nameof(regularPath));
        _appearanceFontFamily = new PdfEmbeddedFontFamily(familyName, System.IO.File.ReadAllBytes(regularPath));
        return this;
    }
}
