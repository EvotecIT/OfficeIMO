using OfficeIMO.Ocr.Tesseract;

namespace OfficeIMO.Studio.Infrastructure;

/// <summary>Studio's OCR provider boundary for the installed distribution channel.</summary>
internal static class StudioOcrProvider {
    internal static Task<TesseractOcrSession> CreateSessionAsync(
        TesseractOcrSessionOptions options, CancellationToken cancellationToken) {
        if (!StudioDistributionPolicy.ExternalToolsAllowed)
            throw new NotSupportedException("OCR requires an external Tesseract installation and is unavailable in the Mac App Store edition. Use the direct desktop edition for OCR.");
        return TesseractOcr.CreateSessionAsync(options, cancellationToken);
    }
}
