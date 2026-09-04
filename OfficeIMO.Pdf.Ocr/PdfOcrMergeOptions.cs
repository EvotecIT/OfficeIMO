using OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Ocr;

/// <summary>Controls canonical parsing, OCR rendering, confidence filtering, and native-text overlap removal.</summary>
public sealed class PdfOcrMergeOptions {
    /// <summary>
    /// Canonical semantic-read settings, including page selection, layout, stage customization,
    /// and understanding budgets. OCR evidence is processed by this same read pipeline.
    /// </summary>
    public PdfReadOptions ReadOptions { get; set; } = PdfReadOptions.Default;
    /// <summary>Requested language tag or provider-specific expression forwarded to the OCR engine.</summary>
    public string? Language { get; set; }
    /// <summary>Optional source path or logical name attached to OCR requests.</summary>
    public string? SourceName { get; set; }
    /// <summary>Optional stable source identifier attached to OCR requests.</summary>
    public string? SourceId { get; set; }
    /// <summary>Uses line spans when the engine does not return word spans.</summary>
    public bool UseLineSpansWhenWordsUnavailable { get; set; } = true;
    /// <summary>Confidence used when neither a span nor the overall result reports one. Defaults to zero.</summary>
    public double ConfidenceWhenUnavailable { get; set; }
    /// <summary>Provider-specific scalar options forwarded to the OCR engine.</summary>
    public IReadOnlyDictionary<string, string> ProviderOptions { get; set; } =
        new Dictionary<string, string>(StringComparer.Ordinal);
    /// <summary>OCR render DPI.</summary>
    public double Dpi { get; set; } = 150D;
    /// <summary>Minimum accepted provider confidence from 0 through 1.</summary>
    public double MinimumConfidence { get; set; } = 0.5D;
    /// <summary>Overlap ratio at which OCR words duplicating native text are removed.</summary>
    public double NativeTextOverlapThreshold { get; set; } = 0.5D;
    /// <summary>Maximum pages sent to the provider.</summary>
    public int MaxPages { get; set; } = 100;
    /// <summary>Maximum total duration for one provider call, including exclusive-engine wait time. Defaults to two minutes.</summary>
    public TimeSpan ProviderTimeout { get; set; } = TimeSpan.FromMinutes(2);
    /// <summary>Maximum pixels rendered per page.</summary>
    public long MaxPixelsPerPage { get; set; } = 100_000_000L;
    /// <summary>Maximum detailed spans inspected from the provider for one page.</summary>
    public int MaxOcrSpansPerPage { get; set; } = 100_000;
    /// <summary>Maximum OCR words accepted from the provider for one page.</summary>
    public int MaxOcrWordsPerPage { get; set; } = 50_000;
    /// <summary>Maximum aggregate selected OCR span characters inspected for one page.</summary>
    public int MaxOcrTextCharactersPerPage { get; set; } = 4 * 1024 * 1024;
    /// <summary>Maximum aggregate raw block, paragraph, and line identifier characters inspected for one page.</summary>
    public int MaxOcrHierarchyCharactersPerPage { get; set; } = 4 * 1024 * 1024;
    /// <summary>Maximum provider diagnostics accepted for one page.</summary>
    public int MaxDiagnosticsPerPage { get; set; } = 1_000;
    /// <summary>Maximum aggregate provider diagnostic characters accepted for one page.</summary>
    public int MaxDiagnosticCharactersPerPage { get; set; } = 1 * 1024 * 1024;
    /// <summary>Maximum aggregate provider, model, and language metadata characters accepted for one page.</summary>
    public int MaxProviderMetadataCharactersPerPage { get; set; } = 16 * 1024;
    /// <summary>Maximum native text blocks merged with OCR output for one page.</summary>
    public int MaxNativeTextBlocksPerPage { get; set; } = 100_000;
    /// <summary>Maximum native-text overlap comparisons performed for one page.</summary>
    public long MaxNativeTextOverlapComparisonsPerPage { get; set; } = 5_000_000L;
    /// <summary>Maximum characters retained in one merged native/OCR text result.</summary>
    public int MaxMergedTextCharactersPerPage { get; set; } = 8 * 1024 * 1024;
    /// <summary>Creates an independent option snapshot.</summary>
    public PdfOcrMergeOptions Clone() {
        Guard.NotNull(ReadOptions, nameof(ReadOptions));
        return new PdfOcrMergeOptions {
            ReadOptions = ReadOptions.Clone(),
            Language = Language,
            SourceName = SourceName,
            SourceId = SourceId,
            UseLineSpansWhenWordsUnavailable = UseLineSpansWhenWordsUnavailable,
            ConfidenceWhenUnavailable = ConfidenceWhenUnavailable,
            ProviderOptions = ProviderOptions == null
                ? new Dictionary<string, string>(StringComparer.Ordinal)
                : ProviderOptions.ToDictionary(static pair => pair.Key, static pair => pair.Value, StringComparer.Ordinal),
            Dpi = Dpi,
            MinimumConfidence = MinimumConfidence,
            NativeTextOverlapThreshold = NativeTextOverlapThreshold,
            MaxPages = MaxPages,
            ProviderTimeout = ProviderTimeout,
            MaxPixelsPerPage = MaxPixelsPerPage,
            MaxOcrSpansPerPage = MaxOcrSpansPerPage,
            MaxOcrWordsPerPage = MaxOcrWordsPerPage,
            MaxOcrTextCharactersPerPage = MaxOcrTextCharactersPerPage,
            MaxOcrHierarchyCharactersPerPage = MaxOcrHierarchyCharactersPerPage,
            MaxDiagnosticsPerPage = MaxDiagnosticsPerPage,
            MaxDiagnosticCharactersPerPage = MaxDiagnosticCharactersPerPage,
            MaxProviderMetadataCharactersPerPage = MaxProviderMetadataCharactersPerPage,
            MaxNativeTextBlocksPerPage = MaxNativeTextBlocksPerPage,
            MaxNativeTextOverlapComparisonsPerPage = MaxNativeTextOverlapComparisonsPerPage,
            MaxMergedTextCharactersPerPage = MaxMergedTextCharactersPerPage
        };
    }

    internal void Validate() {
        Guard.NotNull(ReadOptions, nameof(ReadOptions));
        PdfReadOptions.Resolve(ReadOptions);
        Guard.Positive(Dpi, nameof(Dpi));
        ValidateRatio(MinimumConfidence, nameof(MinimumConfidence));
        ValidateRatio(ConfidenceWhenUnavailable, nameof(ConfidenceWhenUnavailable));
        ValidateRatio(NativeTextOverlapThreshold, nameof(NativeTextOverlapThreshold));
        Guard.PositiveInteger(MaxPages, nameof(MaxPages));
        if (ProviderTimeout <= TimeSpan.Zero) throw new ArgumentOutOfRangeException(nameof(ProviderTimeout));
        if (MaxPixelsPerPage <= 0) throw new ArgumentOutOfRangeException(nameof(MaxPixelsPerPage));
        Guard.PositiveInteger(MaxOcrSpansPerPage, nameof(MaxOcrSpansPerPage));
        Guard.PositiveInteger(MaxOcrWordsPerPage, nameof(MaxOcrWordsPerPage));
        Guard.PositiveInteger(MaxOcrTextCharactersPerPage, nameof(MaxOcrTextCharactersPerPage));
        Guard.PositiveInteger(MaxOcrHierarchyCharactersPerPage, nameof(MaxOcrHierarchyCharactersPerPage));
        Guard.PositiveInteger(MaxDiagnosticsPerPage, nameof(MaxDiagnosticsPerPage));
        Guard.PositiveInteger(MaxDiagnosticCharactersPerPage, nameof(MaxDiagnosticCharactersPerPage));
        Guard.PositiveInteger(MaxProviderMetadataCharactersPerPage, nameof(MaxProviderMetadataCharactersPerPage));
        Guard.PositiveInteger(MaxNativeTextBlocksPerPage, nameof(MaxNativeTextBlocksPerPage));
        if (MaxNativeTextOverlapComparisonsPerPage <= 0) throw new ArgumentOutOfRangeException(nameof(MaxNativeTextOverlapComparisonsPerPage));
        Guard.PositiveInteger(MaxMergedTextCharactersPerPage, nameof(MaxMergedTextCharactersPerPage));
    }

    private static void ValidateRatio(double value, string name) {
        if (value < 0D || value > 1D || double.IsNaN(value)) throw new ArgumentOutOfRangeException(name);
    }
}
