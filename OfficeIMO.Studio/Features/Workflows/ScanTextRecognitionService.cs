using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Ocr.Tesseract;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;

namespace OfficeIMO.Studio.Features.Workflows;

internal interface IScanTextRecognitionService {
    Task<PdfSearchableOcrReview> PrepareAsync(byte[] source, SearchablePdfOcrOptions options, CancellationToken cancellationToken);
}

internal sealed class ScanTextRecognitionService : IScanTextRecognitionService {
    private readonly StudioOcrRuntime? _runtime;
    internal ScanTextRecognitionService(StudioOcrRuntime? runtime = null) => _runtime = runtime;
    public async Task<PdfSearchableOcrReview> PrepareAsync(byte[] source, SearchablePdfOcrOptions options, CancellationToken cancellationToken) {
        var sessionOptions = new TesseractOcrSessionOptions {
            Languages = options.Languages, ProvisionMissingLanguageData = options.ProvisionMissingLanguageData
        };
        var session = await (_runtime is null ? StudioOcrProvider.CreateSessionAsync(sessionOptions, cancellationToken)
            : _runtime.CreateSessionAsync(sessionOptions, cancellationToken)).ConfigureAwait(false);
        var settings = options.Pdf.Clone();
        settings.Language = options.Languages.ToTesseractExpression();
        return await PdfDocument.Load(source).PrepareSearchableOcrAsync(session.Engine, settings, cancellationToken).ConfigureAwait(false);
    }
}
