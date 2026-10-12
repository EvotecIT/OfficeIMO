namespace OfficeIMO.Ocr.Tesseract;

/// <summary>Registers the installed Tesseract CLI engine with the explicit optional-provider catalog.</summary>
/// <remarks>Supported scalar keys are executable, tessdata, language, temporaryDirectory, engineMode,
/// pageSegmentationMode, dpi, and timeoutSeconds. The host must trust filesystem and executable settings;
/// no executable, model, or package is downloaded by this provider.</remarks>
public sealed class TesseractOcrEngineProvider : IOcrEngineProvider {
    /// <inheritdoc />
    public string Id => "tesseract-cli";
    /// <inheritdoc />
    public string DisplayName => "Installed Tesseract CLI";
    /// <inheritdoc />
    public OcrEngineCapabilities Capabilities => new TesseractOcrEngine().Capabilities;

    /// <inheritdoc />
    public IOcrEngine Create(IReadOnlyDictionary<string, string> options) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        var settings = new TesseractOcrEngineOptions();
        foreach (var option in options) {
            if (string.IsNullOrWhiteSpace(option.Value) || option.Value.Length > 4096)
                throw new ArgumentException("Tesseract provider values must contain 1 through 4096 characters.", nameof(options));
            switch (option.Key) {
                case "executable": settings.ExecutablePath = option.Value; break;
                case "tessdata": settings.TessdataDirectory = option.Value; break;
                case "language": settings.Language = option.Value; break;
                case "temporaryDirectory": settings.TemporaryDirectory = option.Value; break;
                case "engineMode": settings.EngineMode = Number(option.Value, option.Key, 0, 3); break;
                case "pageSegmentationMode": settings.PageSegmentationMode = Number(option.Value, option.Key, 0, 13); break;
                case "dpi": settings.Dpi = Number(option.Value, option.Key, 36, 1200); break;
                case "timeoutSeconds": settings.Timeout = TimeSpan.FromSeconds(Number(option.Value, option.Key, 1, 600)); break;
                default: throw new ArgumentException("Unknown Tesseract provider option '" + option.Key + "'.", nameof(options));
            }
        }
        return TesseractOcrEngine.CreateDefault(settings);
    }

    private static int Number(string value, string key, int minimum, int maximum) {
        if (!int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out int number) || number < minimum || number > maximum)
            throw new ArgumentException("Tesseract option '" + key + "' is outside its supported numeric range.");
        return number;
    }
}
