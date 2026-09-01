using OfficeIMO.Pdf;
using OfficeIMO.Reader.Ocr.Tesseract;

namespace OfficeIMO.Reader.Ocr;

/// <summary>Configures the easy local OCR facade without hiding its runtime or PDF policies.</summary>
public sealed class OfficeOcrOptions {
    /// <summary>
    /// Languages to recognize. Combine values with <c>|</c>; English is the default.
    /// </summary>
    public OfficeOcrLanguage Languages { get; set; } = OfficeOcrLanguage.English;

    /// <summary>
    /// Advanced raw Tesseract language expression for caller-installed custom trained-data models.
    /// When set, this overrides <see cref="Languages"/>. Most callers should leave it unset.
    /// </summary>
    public string? CustomLanguageExpression { get; set; }

    /// <summary>
    /// Controls the final atomic searchable-PDF commit. The safe default rejects an existing destination.
    /// </summary>
    public OfficeConversionFileConflictPolicy OutputConflictPolicy { get; set; } = OfficeConversionFileConflictPolicy.FailIfExists;

    /// <summary>
    /// Advanced Tesseract process, timeout, and resource limits. Its raw <c>Language</c> setting remains supported
    /// for compatibility, but should not be combined with non-default <see cref="Languages"/> or
    /// <see cref="CustomLanguageExpression"/> values.
    /// </summary>
    public TesseractOcrEngineOptions Tesseract { get; set; } = new TesseractOcrEngineOptions();

    /// <summary>Searchable-PDF rendering, filtering, overlap, and resource limits.</summary>
    public PdfOcrMergeOptions Pdf { get; set; } = new PdfOcrMergeOptions();

    /// <summary>Checksum-pinned language-data cache and transport configuration.</summary>
    public TesseractLanguageDataOptions LanguageData { get; set; } = new TesseractLanguageDataOptions();

    /// <summary>
    /// Downloads checksum-pinned OfficeIMO catalog language data when the installed runtime lacks a requested language.
    /// Enabled by default for every typed language and orientation data in the curated catalog.
    /// </summary>
    public bool ProvisionMissingLanguageData { get; set; } = true;
}
