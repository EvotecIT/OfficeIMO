namespace OfficeIMO.Pdf;

/// <summary>
/// Controls dependency-free generated text shaping applied before PDF glyph writing.
/// </summary>
public enum PdfTextShapingMode {
    /// <summary>Write one font glyph for each Unicode scalar, with no built-in substitutions.</summary>
    UnicodeScalar = 0,

    /// <summary>Apply built-in substitutions for the standard Latin presentation ligatures ff, fi, fl, ffi, and ffl when the embedded font covers them.</summary>
    LatinLigatures = 1
}

public sealed partial class PdfOptions {
    /// <summary>
    /// Dependency-free generated text shaping mode used by embedded TrueType and OpenType/CFF font output.
    /// </summary>
    public PdfTextShapingMode TextShapingMode {
        get => _textShapingMode;
        set {
            if (value != PdfTextShapingMode.UnicodeScalar && value != PdfTextShapingMode.LatinLigatures) {
                throw new ArgumentOutOfRangeException(nameof(value), "Unsupported PDF text shaping mode.");
            }

            _textShapingMode = value;
        }
    }

    /// <summary>
    /// Optional callback used to provide preferred break positions for long unspaced tokens during generated text wrapping.
    /// </summary>
    public PdfTextHyphenationCallback? TextHyphenationCallback {
        get => _textHyphenationCallback;
        set => _textHyphenationCallback = value;
    }

    internal PdfTextHyphenationCallback? TextHyphenationCallbackSnapshot => _textHyphenationCallback;
    internal PdfTextShapingMode TextShapingModeSnapshot => _textShapingMode;

    internal bool HasDiagnosticsReport => _diagnosticsReport != null;

    internal void AddTextDiagnostics(IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics) {
        if (_diagnosticsReport == null || diagnostics.Count == 0) {
            return;
        }

        _diagnosticsReport.AddTextDiagnostics(diagnostics, _diagnosticsConverter);
    }

    internal void AddTextShapingDiagnostics(IReadOnlyList<PdfTextShapingDiagnostic> diagnostics) {
        if (_diagnosticsReport == null || diagnostics.Count == 0) {
            return;
        }

        foreach (PdfTextShapingDiagnostic diagnostic in diagnostics) {
            if (IsCoveredTextShapingDiagnostic(diagnostic)) {
                continue;
            }

            string key = diagnostic.Code + "|" + diagnostic.Source + "|" + diagnostic.Script;
            if ((_reportedTextShapingDiagnostics ??= new HashSet<string>()).Add(key)) {
                _diagnosticsReport.AddTextShapingDiagnostics(new[] { diagnostic }, _diagnosticsConverter);
            }
        }
    }

    internal void AddFontDiagnostics(PdfStandardFont font, IReadOnlyList<PdfFontEmbeddingDiagnostic> diagnostics) {
        if (_diagnosticsReport == null || diagnostics.Count == 0) {
            return;
        }

        foreach (PdfFontEmbeddingDiagnostic diagnostic in diagnostics) {
            string key = font.ToString() + "|" + diagnostic.Code + "|" + diagnostic.FontName + "|" + diagnostic.Format;
            if ((_reportedEmbeddedFontProgramFailures ??= new HashSet<string>()).Add(key)) {
                _diagnosticsReport.AddFontDiagnostics(new[] { diagnostic }, _diagnosticsConverter);
            }
        }
    }

    /// <summary>
    /// Sets or clears the callback used to provide preferred break positions for long unspaced tokens.
    /// </summary>
    /// <param name="callback">Callback returning UTF-16 break indexes for a token, or null to disable the hook.</param>
    public PdfOptions SetTextHyphenation(PdfTextHyphenationCallback? callback) {
        _textHyphenationCallback = callback;
        return this;
    }

    /// <summary>
    /// Sets the dependency-free generated text shaping mode used by embedded font output.
    /// </summary>
    /// <param name="mode">Shaping mode to apply when writing generated text with embedded fonts.</param>
    public PdfOptions SetTextShapingMode(PdfTextShapingMode mode) {
        TextShapingMode = mode;
        return this;
    }

    /// <summary>
    /// Records generated text diagnostics encountered while writing PDF content.
    /// </summary>
    /// <param name="report">Mutable conversion report that receives text encoding, missing-glyph, shaping, and embedded-font diagnostics.</param>
    /// <param name="converter">Converter or adapter name to place on recorded warnings.</param>
    /// <returns>The current options instance for fluent configuration.</returns>
    /// <remarks>Diagnostics are reported in addition to the existing fail-closed exceptions.</remarks>
    public PdfOptions ReportDiagnosticsTo(PdfConversionReport report, string converter = "OfficeIMO.Pdf") {
        Guard.NotNull(report, nameof(report));
        _diagnosticsReport = report;
        _diagnosticsConverter = string.IsNullOrWhiteSpace(converter) ? "OfficeIMO.Pdf" : converter;
        return this;
    }

    private bool IsCoveredTextShapingDiagnostic(PdfTextShapingDiagnostic diagnostic) =>
        _textShapingMode == PdfTextShapingMode.LatinLigatures &&
        string.Equals(diagnostic.Code, "unsupported-font-ligature-substitution", StringComparison.Ordinal);
}
