using OfficeIMO.Reader.Ocr.Tesseract;

namespace OfficeIMO.Reader.Ocr;

/// <summary>Runtime and language evidence captured when an OCR session is created.</summary>
public sealed class OfficeOcrRuntimeEvidence {
    internal OfficeOcrRuntimeEvidence(
        TesseractRuntimeInfo runtime,
        string engineVersion,
        IReadOnlyList<string> languages,
        TesseractLanguageDataResult? provisionedLanguageData) {
        Runtime = runtime;
        EngineVersion = engineVersion;
        Languages = languages;
        ProvisionedLanguageData = provisionedLanguageData;
    }

    /// <summary>Resolved executable and trained-data location.</summary>
    public TesseractRuntimeInfo Runtime { get; }
    /// <summary>First line reported by <c>tesseract --version</c>.</summary>
    public string EngineVersion { get; }
    /// <summary>Languages reported by the final configured engine.</summary>
    public IReadOnlyList<string> Languages { get; }
    /// <summary>Checksum-verified language data provisioned while creating the session, when needed.</summary>
    public TesseractLanguageDataResult? ProvisionedLanguageData { get; }
}
