using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Ocr.Tesseract;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;

namespace OfficeIMO.Studio.Features.Workflows;

internal interface IScanTextRecognitionService {
    // The review can retain the supplied private storage snapshot for subsequent page reads.
    Task<PdfSearchableOcrReview> PrepareAsync(byte[] ownedSource, SearchablePdfOcrOptions options, CancellationToken cancellationToken);
}

internal sealed class ScanTextRecognitionService : IScanTextRecognitionService {
    private readonly StudioOcrRuntime? _runtime;
    internal ScanTextRecognitionService(StudioOcrRuntime? runtime = null) => _runtime = runtime;
    public async Task<PdfSearchableOcrReview> PrepareAsync(byte[] ownedSource, SearchablePdfOcrOptions options, CancellationToken cancellationToken) {
        var sessionOptions = new TesseractOcrSessionOptions {
            Languages = options.Languages, ProvisionMissingLanguageData = options.ProvisionMissingLanguageData
        };
        var session = await (_runtime is null ? StudioOcrProvider.CreateSessionAsync(sessionOptions, cancellationToken)
            : _runtime.CreateSessionAsync(sessionOptions, cancellationToken)).ConfigureAwait(false);
        var settings = options.Pdf.Clone();
        settings.Language = options.Languages.ToTesseractExpression();
        using PdfDocument document = PdfDocument.LoadOwned(ownedSource, cancellationToken: cancellationToken);
        return await document.PrepareSearchableOcrAsync(session.Engine, settings, cancellationToken).ConfigureAwait(false);
    }
}
