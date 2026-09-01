using OfficeIMO.Pdf;
using OfficeIMO.Reader.Ocr.Tesseract;

namespace OfficeIMO.Reader.Ocr;

/// <summary>Configures the easy local OCR facade without hiding its runtime or PDF policies.</summary>
public sealed class OfficeOcrOptions {
    /// <summary>
    /// Controls the final atomic searchable-PDF commit. The safe default rejects an existing destination.
    /// </summary>
    public OfficeConversionFileConflictPolicy OutputConflictPolicy { get; set; } = OfficeConversionFileConflictPolicy.FailIfExists;

    /// <summary>Tesseract process, language, timeout, and resource limits.</summary>
    public TesseractOcrEngineOptions Tesseract { get; set; } = new TesseractOcrEngineOptions();

    /// <summary>Searchable-PDF rendering, filtering, overlap, and resource limits.</summary>
    public PdfOcrMergeOptions Pdf { get; set; } = new PdfOcrMergeOptions();

    /// <summary>Checksum-pinned language-data cache and transport configuration.</summary>
    public TesseractLanguageDataOptions LanguageData { get; set; } = new TesseractLanguageDataOptions();

    /// <summary>
    /// Downloads checksum-pinned OfficeIMO catalog language data when the installed runtime lacks a requested language.
    /// Enabled by default for the curated <c>eng</c>, <c>pol</c>, and <c>osd</c> catalog.
    /// </summary>
    public bool ProvisionMissingLanguageData { get; set; } = true;
}
