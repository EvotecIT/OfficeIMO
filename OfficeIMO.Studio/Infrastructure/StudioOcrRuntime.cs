using OfficeIMO.Ocr;
using OfficeIMO.Ocr.Tesseract;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Infrastructure;

internal sealed record StudioOcrRuntimeInfo(string ExecutablePath, string Version, IReadOnlyList<string> Languages);

/// <summary>Shares the user's OCR executable choice across Studio recognition workflows.</summary>
internal sealed class StudioOcrRuntime(StudioPreferencesService preferences) {
    internal async Task<StudioOcrRuntimeInfo> InspectAsync(string? executablePath, CancellationToken cancellationToken) {
        if (StudioOcrProvider.UnavailableReason is { } reason) throw new NotSupportedException(reason);
        TesseractRuntimeInfo runtime = TesseractRuntime.Discover(executablePath);
        var engine = TesseractOcrEngine.CreateDefault(new() {
            ExecutablePath = runtime.ExecutablePath, TessdataDirectory = runtime.TessdataDirectory,
            Timeout = TimeSpan.FromSeconds(10), MaxProcessOutputCharacters = 16 * 1024
        });
        string version = await engine.GetVersionAsync(cancellationToken).ConfigureAwait(false);
        if (!version.StartsWith("tesseract ", StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("The selected executable did not identify itself as Tesseract.");
        IReadOnlyList<string> languages = await engine.GetLanguagesAsync(cancellationToken).ConfigureAwait(false);
        return new(runtime.ExecutablePath, version, languages);
    }

    internal Task<TesseractOcrSession> CreateSessionAsync(TesseractOcrSessionOptions options, CancellationToken cancellationToken) {
        ArgumentNullException.ThrowIfNull(options);
        TesseractOcrEngineOptions engine = options.Engine.Clone();
        engine.ExecutablePath = preferences.Current.OcrExecutablePath ?? engine.ExecutablePath;
        return StudioOcrProvider.CreateSessionAsync(new() {
            Languages = options.Languages, CustomLanguageExpression = options.CustomLanguageExpression,
            Engine = engine, LanguageData = options.LanguageData,
            ProvisionMissingLanguageData = options.ProvisionMissingLanguageData
        }, cancellationToken);
    }

    internal async Task<IOcrEngine> CreateEngineAsync(TesseractOcrLanguage languages, bool provision, CancellationToken token) =>
        (await CreateSessionAsync(new() { Languages = languages, ProvisionMissingLanguageData = provision }, token)
            .ConfigureAwait(false)).Engine;
}
