using OfficeIMO.Ocr.Tesseract;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;

namespace OfficeIMO.Studio.Features.Workflows;

internal interface IScanTextRecognitionService {
    Task<PdfSearchableOcrReview> PrepareAsync(byte[] source, SearchablePdfOcrOptions options, CancellationToken cancellationToken);
}

internal sealed class ScanTextRecognitionService : IScanTextRecognitionService {
    public async Task<PdfSearchableOcrReview> PrepareAsync(byte[] source, SearchablePdfOcrOptions options, CancellationToken cancellationToken) {
        var session = await TesseractOcr.CreateSessionAsync(new TesseractOcrSessionOptions {
            Languages = options.Languages, ProvisionMissingLanguageData = options.ProvisionMissingLanguageData
        }, cancellationToken).ConfigureAwait(false);
        var settings = options.Pdf.Clone();
        settings.Language = options.Languages.ToTesseractExpression();
        return await PdfDocument.Load(source).PrepareSearchableOcrAsync(session.Engine, settings, cancellationToken).ConfigureAwait(false);
    }
}
