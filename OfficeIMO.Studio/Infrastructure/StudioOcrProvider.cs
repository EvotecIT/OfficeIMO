using OfficeIMO.Ocr.Tesseract;

namespace OfficeIMO.Studio.Infrastructure;

/// <summary>Studio's OCR provider boundary for the installed distribution channel.</summary>
internal static class StudioOcrProvider {
    internal static string? UnavailableReason => StudioDistributionPolicy.ExternalToolsAllowed ? null
        : OperatingSystem.IsIOS()
            ? "Text recognition requires an OCR provider that runs on iPhone and iPad. The desktop Tesseract provider cannot run on iOS. Scan preparation remains available."
            : "OCR requires an external Tesseract installation and is unavailable in the Mac App Store edition. Use the direct desktop edition for OCR.";

    internal static Task<TesseractOcrSession> CreateSessionAsync(
        TesseractOcrSessionOptions options, CancellationToken cancellationToken) {
        if (UnavailableReason is { } reason) throw new NotSupportedException(reason);
        return TesseractOcr.CreateSessionAsync(options, cancellationToken);
    }
}
