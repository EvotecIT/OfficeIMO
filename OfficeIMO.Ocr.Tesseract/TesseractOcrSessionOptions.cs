namespace OfficeIMO.Ocr.Tesseract;

/// <summary>Configures automatic Tesseract discovery and optional trained-data provisioning.</summary>
public sealed class TesseractOcrSessionOptions {
    /// <summary>Discoverable languages to recognize. Combine values with <c>|</c>; English is the default.</summary>
    public TesseractOcrLanguage Languages { get; set; } = TesseractOcrLanguage.English;

    /// <summary>
    /// Advanced raw Tesseract language expression for caller-installed custom trained-data models.
    /// When set, this overrides <see cref="Languages"/>.
    /// </summary>
    public string? CustomLanguageExpression { get; set; }

    /// <summary>Advanced Tesseract process, timeout, and resource limits.</summary>
    public TesseractOcrEngineOptions Engine { get; set; } = new TesseractOcrEngineOptions();

    /// <summary>Checksum-pinned language-data cache and transport configuration.</summary>
    public TesseractLanguageDataOptions LanguageData { get; set; } = new TesseractLanguageDataOptions();

    /// <summary>Downloads checksum-pinned curated language data when a requested language is missing.</summary>
    public bool ProvisionMissingLanguageData { get; set; } = true;
}
