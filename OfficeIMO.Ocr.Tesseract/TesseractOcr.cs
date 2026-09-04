using OfficeIMO.Ocr;

namespace OfficeIMO.Ocr.Tesseract;

/// <summary>Automatic local Tesseract discovery with reusable session support.</summary>
public static class TesseractOcr {
    /// <summary>Discovers Tesseract, validates it, and provisions missing curated language data when enabled.</summary>
    public static async Task<TesseractOcrSession> CreateSessionAsync(
        TesseractOcrSessionOptions? options = null,
        CancellationToken cancellationToken = default) {
        TesseractOcrSessionOptions source = options ?? new TesseractOcrSessionOptions();
        if (source.Engine == null) throw new ArgumentException("Tesseract engine options cannot be null.", nameof(options));
        if (source.LanguageData == null) throw new ArgumentException("Language-data options cannot be null.", nameof(options));
        TesseractOcrEngineOptions engineOptions = source.Engine.Clone();
        string languageExpression = ResolveLanguageExpression(source);
        engineOptions.Language = languageExpression;
        TesseractRuntimeInfo runtime = TesseractRuntime.Discover(engineOptions.ExecutablePath);
        engineOptions.ExecutablePath = runtime.ExecutablePath;
        if (string.IsNullOrWhiteSpace(engineOptions.TessdataDirectory) && runtime.TessdataDirectory != null) {
            engineOptions.TessdataDirectory = runtime.TessdataDirectory;
        }

        var engine = new TesseractOcrEngine(engineOptions);
        string version = await engine.GetVersionAsync(cancellationToken).ConfigureAwait(false);
        IReadOnlyList<string> languages = await engine.GetLanguagesAsync(cancellationToken).ConfigureAwait(false);
        string[] requestedLanguages = ResolveRequiredLanguageData(languageExpression, engineOptions.PageSegmentationMode);
        string requiredLanguageExpression = string.Join("+", requestedLanguages);
        TesseractLanguageDataResult? provisioned = null;
        if (requestedLanguages.Any(language => !languages.Contains(language, StringComparer.Ordinal))) {
            if (!source.ProvisionMissingLanguageData || !string.IsNullOrWhiteSpace(source.Engine.TessdataDirectory)) {
                throw new InvalidOperationException(
                    "The configured Tesseract runtime does not provide every required language-data file (" + requiredLanguageExpression + "). " +
                    "Install the missing trained data, configure TessdataDirectory, or enable ProvisionMissingLanguageData.");
            }
            provisioned = await TesseractLanguageData.EnsureAsync(requiredLanguageExpression, source.LanguageData, cancellationToken).ConfigureAwait(false);
            engineOptions.TessdataDirectory = provisioned.Directory;
            engine = new TesseractOcrEngine(engineOptions);
            languages = await engine.GetLanguagesAsync(cancellationToken).ConfigureAwait(false);
            if (requestedLanguages.Any(language => !languages.Contains(language, StringComparer.Ordinal))) {
                throw new InvalidOperationException("Tesseract did not report every requested language after checksum-verified provisioning.");
            }
        }

        var evidenceRuntime = new TesseractRuntimeInfo(runtime.ExecutablePath, engineOptions.TessdataDirectory, runtime.Source);
        var evidence = new TesseractOcrRuntimeEvidence(evidenceRuntime, version, languages, provisioned);
        return new TesseractOcrSession(engine, engineOptions.Language, evidence);
    }

    /// <summary>Recognizes a supported image file with automatic runtime discovery.</summary>
    public static async Task<OcrResult> RecognizeFileAsync(
        string imagePath,
        TesseractOcrSessionOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(imagePath)) throw new ArgumentException("Image path cannot be empty.", nameof(imagePath));
        string fullPath = Path.GetFullPath(imagePath);
        TesseractOcrSessionOptions effective = options ?? new TesseractOcrSessionOptions();
        if (effective.Engine == null) throw new ArgumentException("Tesseract engine options cannot be null.", nameof(options));
        var info = new FileInfo(fullPath);
        if (info.Length > effective.Engine.MaxInputBytes) {
            throw new IOException("OCR image exceeds the configured Tesseract MaxInputBytes limit.");
        }
        TesseractOcrSession session = await CreateSessionAsync(effective, cancellationToken).ConfigureAwait(false);
        byte[] bytes = File.ReadAllBytes(fullPath);
        return await session.RecognizeAsync(bytes, MediaTypeFor(fullPath), fullPath, cancellationToken).ConfigureAwait(false);
    }

    internal static string[] ResolveRequiredLanguageData(string languageExpression, int? pageSegmentationMode) {
        string[] languages = languageExpression
            .Split('+')
            .Select(static language => language.Trim())
            .Where(static language => language.Length > 0)
            .Distinct(StringComparer.Ordinal)
            .ToArray();
        if (pageSegmentationMode is 0 or 1 or 12 && !languages.Contains("osd", StringComparer.Ordinal)) {
            return languages.Concat(new[] { "osd" }).ToArray();
        }
        return languages;
    }

    internal static string ResolveLanguageExpression(TesseractOcrSessionOptions options) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        if (options.Engine == null) throw new ArgumentException("Tesseract engine options cannot be null.", nameof(options));
        string? custom = string.IsNullOrWhiteSpace(options.CustomLanguageExpression) ? null : options.CustomLanguageExpression!.Trim();
        string? engineDefault = string.IsNullOrWhiteSpace(options.Engine.Language) ? null : options.Engine.Language!.Trim();
        string typed = options.Languages.ToTesseractExpression();
        bool hasEngineOverride = engineDefault != null && !string.Equals(engineDefault, "eng", StringComparison.Ordinal);
        bool hasTypedOverride = options.Languages != TesseractOcrLanguage.English;

        if (custom != null) {
            if (hasTypedOverride) throw new ArgumentException("Use Languages or CustomLanguageExpression, not both.", nameof(options));
            if (hasEngineOverride) throw new ArgumentException("Use CustomLanguageExpression or Engine.Language, not both.", nameof(options));
            return custom;
        }
        if (hasTypedOverride) {
            if (hasEngineOverride) throw new ArgumentException("Use Languages or Engine.Language, not both.", nameof(options));
            return typed;
        }
        return engineDefault ?? typed;
    }

    private static string MediaTypeFor(string path) => Path.GetExtension(path).ToLowerInvariant() switch {
        ".png" => "image/png",
        ".jpg" or ".jpeg" => "image/jpeg",
        ".tif" or ".tiff" => "image/tiff",
        ".bmp" => "image/bmp",
        ".gif" => "image/gif",
        ".webp" => "image/webp",
        ".jp2" or ".j2k" => "image/jp2",
        _ => throw new NotSupportedException("Tesseract OCR supports PNG, JPEG, TIFF, BMP, GIF, WebP, and JPEG 2000 image files.")
    };
}
