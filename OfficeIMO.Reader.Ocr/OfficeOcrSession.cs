using OfficeIMO.Pdf;
using OfficeIMO.Reader.Ocr.Tesseract;
using OfficeIMO.Reader.Pdf;

namespace OfficeIMO.Reader.Ocr;

/// <summary>Reusable, fully resolved OCR session for images and PDFs.</summary>
public sealed class OfficeOcrSession {
    private readonly IOfficeOcrEngine _engine;
    private readonly string? _language;
    private readonly PdfOcrMergeOptions _pdfOptions;

    internal OfficeOcrSession(
        IOfficeOcrEngine engine,
        string? language,
        PdfOcrMergeOptions pdfOptions,
        OfficeOcrRuntimeEvidence runtime) {
        _engine = engine ?? throw new ArgumentNullException(nameof(engine));
        _language = string.IsNullOrWhiteSpace(language) ? null : language!.Trim();
        _pdfOptions = pdfOptions.Clone();
        Runtime = runtime;
    }

    /// <summary>Resolved runtime and trained-data evidence.</summary>
    public OfficeOcrRuntimeEvidence Runtime { get; }

    /// <summary>The reusable engine-neutral OCR engine.</summary>
    public IOfficeOcrEngine Engine => _engine;

    /// <summary>Recognizes one supported raster payload.</summary>
    public async Task<OfficeOcrEngineResult> RecognizeImageAsync(
        byte[] image,
        string mediaType,
        string? sourceName = null,
        CancellationToken cancellationToken = default) {
        if (image == null) throw new ArgumentNullException(nameof(image));
        if (image.Length == 0) throw new ArgumentException("OCR image payload cannot be empty.", nameof(image));
        if (string.IsNullOrWhiteSpace(mediaType)) throw new ArgumentException("OCR image media type cannot be empty.", nameof(mediaType));
        string hash = OfficeDocumentAssetHash.ComputeSha256Hex(image);
        var location = new ReaderLocation { Path = sourceName, SourceBlockKind = "image" };
        var asset = new OfficeDocumentAsset {
            Id = "ocr-image",
            Kind = "image",
            MediaType = mediaType.Trim(),
            Extension = ExtensionFor(mediaType),
            FileName = sourceName,
            LengthBytes = image.LongLength,
            PayloadHash = hash,
            PayloadBytes = (byte[])image.Clone(),
            Location = location
        };
        var request = new OfficeOcrEngineRequest {
            Payload = (byte[])image.Clone(),
            Asset = asset,
            Candidate = new OfficeDocumentOcrCandidate {
                Id = "ocr-image",
                Kind = "image",
                Reason = "Direct image OCR request.",
                Confidence = 1D,
                AssetId = asset.Id,
                Location = location
            },
            Language = _language,
            Source = new OfficeDocumentSource {
                Path = sourceName,
                SourceId = "ocr-image",
                SourceHash = hash,
                LengthBytes = image.LongLength
            }
        };
        return await _engine.RecognizeAsync(request, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Recognizes and merges selected PDF pages without rewriting the input artifact.</summary>
    public Task<PdfOcrMergeResult> ReadPdfAsync(PdfDocument document, CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return document.Ocr.ReadAsync(CreatePdfProvider(), _pdfOptions, cancellationToken);
    }

    /// <summary>Creates a searchable PDF with invisible text aligned to accepted OCR word geometry.</summary>
    public Task<PdfSearchableOcrResult> MakePdfSearchableAsync(PdfDocument document, CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return document.Ocr.MakeSearchableAsync(CreatePdfProvider(), _pdfOptions, cancellationToken);
    }

    private OfficeOcrEnginePdfProvider CreatePdfProvider() => new OfficeOcrEnginePdfProvider(
        _engine,
        new OfficeOcrEnginePdfProviderOptions { Language = _language });

    private static string? ExtensionFor(string mediaType) => mediaType.Trim().ToLowerInvariant() switch {
        "image/png" => ".png",
        "image/jpeg" or "image/jpg" => ".jpg",
        "image/tiff" => ".tiff",
        "image/bmp" => ".bmp",
        "image/gif" => ".gif",
        "image/webp" => ".webp",
        "image/jp2" => ".jp2",
        _ => null
    };
}
