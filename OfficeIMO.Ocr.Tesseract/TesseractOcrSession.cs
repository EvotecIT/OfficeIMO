using OfficeIMO.Ocr;

namespace OfficeIMO.Ocr.Tesseract;

/// <summary>Reusable, fully resolved Tesseract OCR session for raster images from any document format.</summary>
public sealed class TesseractOcrSession {
    private readonly TesseractOcrEngine _engine;
    private readonly string? _language;

    internal TesseractOcrSession(
        TesseractOcrEngine engine,
        string? language,
        TesseractOcrRuntimeEvidence runtime) {
        _engine = engine ?? throw new ArgumentNullException(nameof(engine));
        _language = string.IsNullOrWhiteSpace(language) ? null : language!.Trim();
        Runtime = runtime;
    }

    /// <summary>Resolved runtime and trained-data evidence.</summary>
    public TesseractOcrRuntimeEvidence Runtime { get; }

    /// <summary>The reusable engine-neutral OCR engine.</summary>
    public IOcrEngine Engine => _engine;

    /// <summary>Recognizes one supported raster payload.</summary>
    public Task<OcrResult> RecognizeAsync(
        byte[] image,
        string mediaType,
        string? sourceName = null,
        CancellationToken cancellationToken = default) {
        if (image == null) throw new ArgumentNullException(nameof(image));
        if (image.Length == 0) throw new ArgumentException("OCR image payload cannot be empty.", nameof(image));
        if (string.IsNullOrWhiteSpace(mediaType)) throw new ArgumentException("OCR image media type cannot be empty.", nameof(mediaType));
        return _engine.RecognizeAsync(new OcrRequest {
            Payload = (byte[])image.Clone(),
            MediaType = mediaType.Trim(),
            FileName = sourceName == null ? null : Path.GetFileName(sourceName),
            SourceName = sourceName,
            CandidateId = "image",
            CandidateKind = "image",
            Language = _language
        }, cancellationToken);
    }
}
