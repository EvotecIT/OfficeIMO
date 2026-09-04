namespace OfficeIMO.Ocr.Tesseract;

public sealed partial class TesseractOcrEngine {
    /// <summary>Creates an engine using explicit, environment, PATH, and known-location runtime discovery.</summary>
    public static TesseractOcrEngine CreateDefault(TesseractOcrEngineOptions? options = null) {
        TesseractOcrEngineOptions effective = (options ?? new TesseractOcrEngineOptions()).Clone();
        TesseractRuntimeInfo runtime = TesseractRuntime.Discover(effective.ExecutablePath);
        effective.ExecutablePath = runtime.ExecutablePath;
        if (string.IsNullOrWhiteSpace(effective.TessdataDirectory) && runtime.TessdataDirectory != null) {
            effective.TessdataDirectory = runtime.TessdataDirectory;
        }
        return new TesseractOcrEngine(effective);
    }
}
