using OfficeIMO.Drawing;

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
    /// Optional host-provided shaping engine for embedded-font text runs that need script shaping, bidirectional layout, or glyph substitution beyond the built-in dependency-free modes.
    /// </summary>
    public IOfficeTextShapingProvider? TextShapingProvider {
        get => _textShapingProvider;
        set => _textShapingProvider = value;
    }

    /// <summary>
    /// Optional callback used to provide preferred break positions for long unspaced tokens during generated text wrapping.
    /// </summary>
    public PdfTextHyphenationCallback? TextHyphenationCallback {
        get => _textHyphenationCallback;
        set => _textHyphenationCallback = value;
    }

    /// <summary>
    /// Optional callback used to provide preferred non-hyphenating break positions for long unspaced tokens during generated text wrapping.
    /// </summary>
    public Func<string, IReadOnlyList<int>>? TextLineBreakCallback {
        get => _textLineBreakCallback;
        set => _textLineBreakCallback = value;
    }

    internal Func<string, IReadOnlyList<int>>? TextLineBreakCallbackSnapshot => _textLineBreakCallback;
    internal PdfTextHyphenationCallback? TextHyphenationCallbackSnapshot => _textHyphenationCallback;
    internal PdfTextShapingMode TextShapingModeSnapshot => _textShapingMode;
    internal IOfficeTextShapingProvider? TextShapingProviderSnapshot => _textShapingProvider;

    internal bool HasDiagnosticsReport => _diagnosticsReport != null;

    internal void AddTextDiagnostics(IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics) {
        if (_diagnosticsReport == null || diagnostics.Count == 0) {
            return;
        }

        _diagnosticsReport.AddTextDiagnostics(diagnostics, _diagnosticsConverter);
    }

    internal void AddTextShapingDiagnostics(IReadOnlyList<PdfTextShapingDiagnostic> diagnostics) {
        AddTextShapingDiagnostics(diagnostics, text: null, fontName: null, isOpenTypeCff: false);
    }

    internal void AddTextShapingDiagnostics(IReadOnlyList<PdfTextShapingDiagnostic> diagnostics, bool deferProviderCoverable) {
        AddTextShapingDiagnostics(diagnostics, text: null, fontName: null, isOpenTypeCff: false, deferProviderCoverable);
    }

    internal void AddTextShapingDiagnostics(IReadOnlyList<PdfTextShapingDiagnostic> diagnostics, string? text, bool deferProviderCoverable) {
        AddTextShapingDiagnostics(diagnostics, text, fontName: null, isOpenTypeCff: false, deferProviderCoverable);
    }

    internal void AddTextShapingDiagnostics(IReadOnlyList<PdfTextShapingDiagnostic> diagnostics, string? text, string? fontName, bool isOpenTypeCff) {
        AddTextShapingDiagnostics(diagnostics, text, fontName, isOpenTypeCff, deferProviderCoverable: false);
    }

    private void AddTextShapingDiagnostics(IReadOnlyList<PdfTextShapingDiagnostic> diagnostics, string? text, string? fontName, bool isOpenTypeCff, bool deferProviderCoverable) {
        if (_diagnosticsReport == null || diagnostics.Count == 0) {
            return;
        }

        foreach (PdfTextShapingDiagnostic diagnostic in diagnostics) {
            if (deferProviderCoverable && _textShapingProvider != null && IsProviderCoveredDiagnostic(diagnostic)) {
                continue;
            }

            if (IsCoveredTextShapingDiagnostic(diagnostic, text, fontName, isOpenTypeCff)) {
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

    /// <summary>Uses or clears an immutable first-party word hyphenation dictionary.</summary>
    /// <param name="dictionary">Dictionary whose breakpoints should be used, or null to clear hyphenation.</param>
    public PdfOptions UseTextHyphenationDictionary(PdfHyphenationLexicon? dictionary) {
        _textHyphenationCallback = dictionary?.AsCallback();
        return this;
    }

    /// <summary>
    /// Sets or clears the callback used to provide preferred non-hyphenating break positions for long unspaced tokens.
    /// </summary>
    /// <param name="callback">Callback returning UTF-16 break indexes for a token, or null to disable the hook.</param>
    public PdfOptions SetTextLineBreaks(Func<string, IReadOnlyList<int>>? callback) {
        _textLineBreakCallback = callback;
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

    internal void RecordProviderShapedTextRun(string text, string fontName, bool isOpenTypeCff) {
        if (_textShapingProvider == null || string.IsNullOrEmpty(text)) {
            return;
        }

        (_providerShapedTextRuns ??= new HashSet<string>()).Add(BuildProviderShapedTextRunKey(text, fontName, isOpenTypeCff));
    }

    /// <summary>
    /// Sets or clears the host-provided shaping engine used for generated text written with embedded fonts.
    /// </summary>
    /// <param name="provider">Provider that returns shaped glyph runs, or <c>null</c> to use only built-in shaping.</param>
    /// <returns>The current options instance for fluent configuration.</returns>
    public PdfOptions SetTextShapingProvider(IOfficeTextShapingProvider? provider) {
        TextShapingProvider = provider;
        return this;
    }

    /// <summary>
    /// Uses the shared dependency-light Drawing provider for bounded Arabic joining and bidirectional text.
    /// </summary>
    /// <remarks>
    /// The provider declines scripts and font programs outside its proven managed subset, allowing the
    /// normal PDF fallback and shaping diagnostics to remain authoritative.
    /// </remarks>
    public PdfOptions UseManagedTextShaping() {
        TextShapingProvider = OfficeManagedTextShapingProvider.Instance;
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
        _reportedEmbeddedFontProgramFailures?.Clear();
        _reportedTextShapingDiagnostics?.Clear();
        _providerShapedTextRuns?.Clear();
        return this;
    }

    private bool IsCoveredTextShapingDiagnostic(PdfTextShapingDiagnostic diagnostic, string? text = null, string? fontName = null, bool isOpenTypeCff = false) =>
        (_textShapingMode == PdfTextShapingMode.LatinLigatures &&
            diagnostic.IsCoveredByBuiltInShaping &&
            string.Equals(diagnostic.Code, "unsupported-font-ligature-substitution", StringComparison.Ordinal)) ||
        IsCoveredByTextLineBreakCallback(diagnostic, text) ||
        (IsProviderCoveredDiagnostic(diagnostic) &&
            !string.IsNullOrEmpty(text) &&
            _providerShapedTextRuns != null &&
            _providerShapedTextRuns.Contains(BuildProviderShapedTextRunKey(text!, fontName, isOpenTypeCff)));

    private static bool IsProviderCoveredDiagnostic(PdfTextShapingDiagnostic diagnostic) =>
        string.Equals(diagnostic.Code, "unsupported-complex-script-shaping", StringComparison.Ordinal) ||
        string.Equals(diagnostic.Code, "unsupported-bidirectional-text-layout", StringComparison.Ordinal) ||
        string.Equals(diagnostic.Code, "unsupported-font-ligature-substitution", StringComparison.Ordinal) ||
        string.Equals(diagnostic.Code, "unsupported-mark-positioning-or-joiner-shaping", StringComparison.Ordinal) ||
        string.Equals(diagnostic.Code, "unsupported-font-mark-positioning", StringComparison.Ordinal);

    private bool IsCoveredByTextLineBreakCallback(PdfTextShapingDiagnostic diagnostic, string? text) =>
        _textLineBreakCallback != null &&
        !string.IsNullOrEmpty(text) &&
        string.Equals(diagnostic.Code, "unsupported-script-specific-line-breaking", StringComparison.Ordinal) &&
        HasValidTextLineBreakPoint(text!);

    private bool HasValidTextLineBreakPoint(string text) {
        System.Collections.Generic.IReadOnlyList<int>? points = _textLineBreakCallback?.Invoke(text);
        if (points == null || points.Count == 0) {
            return false;
        }

        foreach (int point in points) {
            if (point > 0 &&
                point < text.Length &&
                !(char.IsHighSurrogate(text[point - 1]) && char.IsLowSurrogate(text[point]))) {
                return true;
            }
        }

        return false;
    }

    private static string BuildProviderShapedTextRunKey(string text, string? fontName, bool isOpenTypeCff) =>
        (isOpenTypeCff ? "cff" : "ttf") + "|" + (fontName ?? string.Empty) + "|" + text;
}
