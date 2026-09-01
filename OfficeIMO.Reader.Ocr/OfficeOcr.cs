using OfficeIMO.Pdf;
using OfficeIMO.Reader.Ocr.Tesseract;

namespace OfficeIMO.Reader.Ocr;

/// <summary>One-call local OCR entry points with reusable session support for repeated work.</summary>
public static class OfficeOcr {
    /// <summary>Discovers Tesseract, validates it, and provisions missing curated language data when enabled.</summary>
    public static async Task<OfficeOcrSession> CreateSessionAsync(
        OfficeOcrOptions? options = null,
        CancellationToken cancellationToken = default) {
        OfficeOcrOptions source = options ?? new OfficeOcrOptions();
        if (source.Tesseract == null) throw new ArgumentException("Tesseract options cannot be null.", nameof(options));
        if (source.Pdf == null) throw new ArgumentException("PDF OCR options cannot be null.", nameof(options));
        if (source.LanguageData == null) throw new ArgumentException("Language-data options cannot be null.", nameof(options));
        TesseractOcrEngineOptions engineOptions = source.Tesseract.Clone();
        TesseractRuntimeInfo runtime = TesseractRuntime.Discover(engineOptions.ExecutablePath);
        engineOptions.ExecutablePath = runtime.ExecutablePath;
        if (string.IsNullOrWhiteSpace(engineOptions.TessdataDirectory) && runtime.TessdataDirectory != null) {
            engineOptions.TessdataDirectory = runtime.TessdataDirectory;
        }

        var engine = new TesseractOcrEngine(engineOptions);
        string version = await engine.GetVersionAsync(cancellationToken).ConfigureAwait(false);
        IReadOnlyList<string> languages = await engine.GetLanguagesAsync(cancellationToken).ConfigureAwait(false);
        string languageExpression = string.IsNullOrWhiteSpace(engineOptions.Language) ? "eng" : engineOptions.Language!;
        string[] requestedLanguages = ParseLanguages(languageExpression);
        TesseractLanguageDataResult? provisioned = null;
        if (requestedLanguages.Any(language => !languages.Contains(language, StringComparer.Ordinal))) {
            if (!source.ProvisionMissingLanguageData || !string.IsNullOrWhiteSpace(source.Tesseract.TessdataDirectory)) {
                throw new InvalidOperationException(
                    "The configured Tesseract runtime does not provide every requested language (" + languageExpression + "). " +
                    "Install the missing trained data, configure TessdataDirectory, or enable ProvisionMissingLanguageData.");
            }
            provisioned = await TesseractLanguageData.EnsureAsync(languageExpression, source.LanguageData, cancellationToken).ConfigureAwait(false);
            engineOptions.TessdataDirectory = provisioned.Directory;
            engine = new TesseractOcrEngine(engineOptions);
            languages = await engine.GetLanguagesAsync(cancellationToken).ConfigureAwait(false);
            if (requestedLanguages.Any(language => !languages.Contains(language, StringComparer.Ordinal))) {
                throw new InvalidOperationException("Tesseract did not report every requested language after checksum-verified provisioning.");
            }
        }

        var evidenceRuntime = new TesseractRuntimeInfo(runtime.ExecutablePath, engineOptions.TessdataDirectory, runtime.Source);
        var evidence = new OfficeOcrRuntimeEvidence(evidenceRuntime, version, languages, provisioned);
        return new OfficeOcrSession(engine, engineOptions.Language, source.Pdf, evidence);
    }

    /// <summary>Recognizes a supported image file with automatic runtime discovery.</summary>
    public static async Task<OfficeOcrEngineResult> ReadTextAsync(
        string imagePath,
        OfficeOcrOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(imagePath)) throw new ArgumentException("Image path cannot be empty.", nameof(imagePath));
        string fullPath = Path.GetFullPath(imagePath);
        OfficeOcrOptions effective = options ?? new OfficeOcrOptions();
        var info = new FileInfo(fullPath);
        if (info.Length > effective.Tesseract.MaxInputBytes) {
            throw new IOException("OCR image exceeds the configured Tesseract MaxInputBytes limit.");
        }
        var session = await CreateSessionAsync(effective, cancellationToken).ConfigureAwait(false);
        byte[] bytes = File.ReadAllBytes(fullPath);
        return await session.RecognizeImageAsync(bytes, MediaTypeFor(fullPath), fullPath, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Creates and atomically saves a searchable PDF in one call.</summary>
    public static async Task<PdfSearchableOcrResult> MakePdfSearchableAsync(
        string inputPath,
        string outputPath,
        OfficeOcrOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(inputPath)) throw new ArgumentException("Input PDF path cannot be empty.", nameof(inputPath));
        if (string.IsNullOrWhiteSpace(outputPath)) throw new ArgumentException("Output PDF path cannot be empty.", nameof(outputPath));
        var session = await CreateSessionAsync(options, cancellationToken).ConfigureAwait(false);
        PdfDocument document = await PdfDocument.LoadAsync(inputPath, cancellationToken: cancellationToken).ConfigureAwait(false);
        PdfSearchableOcrResult result = await session.MakePdfSearchableAsync(document, cancellationToken).ConfigureAwait(false);
        await result.Document.SaveAsync(outputPath, cancellationToken).ConfigureAwait(false);
        return result;
    }

    private static string[] ParseLanguages(string expression) => expression
        .Split('+')
        .Select(static language => language.Trim())
        .Where(static language => language.Length > 0)
        .Distinct(StringComparer.Ordinal)
        .ToArray();

    private static string MediaTypeFor(string path) => Path.GetExtension(path).ToLowerInvariant() switch {
        ".png" => "image/png",
        ".jpg" or ".jpeg" => "image/jpeg",
        ".tif" or ".tiff" => "image/tiff",
        ".bmp" => "image/bmp",
        ".gif" => "image/gif",
        ".webp" => "image/webp",
        ".jp2" or ".j2k" => "image/jp2",
        _ => throw new NotSupportedException("The easy OCR facade supports PNG, JPEG, TIFF, BMP, GIF, WebP, and JPEG 2000 image files.")
    };
}
