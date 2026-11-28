using OfficeIMO.Ocr.Tesseract;

namespace OfficeIMO.Studio.Infrastructure;

/// <summary>Studio's OCR provider boundary: the installed Tesseract runtime.</summary>
internal static class StudioOcrProvider {
    internal static Task<TesseractOcrSession> CreateSessionAsync(
        TesseractOcrSessionOptions options, CancellationToken cancellationToken) =>
        TesseractOcr.CreateSessionAsync(options, cancellationToken);
}
