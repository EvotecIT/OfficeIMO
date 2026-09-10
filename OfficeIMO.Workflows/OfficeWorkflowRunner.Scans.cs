using System.Security.Cryptography;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;

namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    private static OperationArtifact CleanScan(ValidatedRequest request, List<OfficeWorkflowDiagnostic> diagnostics, CancellationToken token) {
        var settings = request.ScanCleanup ?? throw new InvalidOperationException("Scan settings are missing.");
        byte[] input = ReadInput(request.InputPath, request.Limits, token);
        if (settings.ExpectedSourceSha256 != null && !string.Equals(settings.ExpectedSourceSha256,
            System.Convert.ToHexString(SHA256.HashData(input)), StringComparison.OrdinalIgnoreCase))
            throw new IOException("The source changed since scan review. Preview the current source before saving.");
        PdfDocument source = PdfDocument.Load(input, request.PdfLoadOptions);
        int pageCount = source.Inspect(request.PdfLoadOptions, token).PageCount;
        var preparation = settings.Preparation;
        int[] pages = preparation.GetSelectedPages(pageCount);
        if (pages.Length == 0 || pages.Length > preparation.MaxPages) throw new InvalidOperationException("Scan page selection exceeds its page limit.");
        var documents = new List<PdfDocument>();
        long retainedBytes = 0;
        foreach (int page in pages) {
            token.ThrowIfCancellationRequested();
            PdfScanPreview preview = source.PreviewScanAsync(page, preparation, token).GetAwaiter().GetResult();
            PdfDocument output = preview.CreateImagePdf(token);
            retainedBytes = checked(retainedBytes + output.ToBytes().LongLength);
            if (retainedBytes > request.Limits.MaximumOutputBytes) throw new IOException("Prepared scan pages exceed the output byte budget.");
            documents.Add(output);
            foreach (string message in preview.Diagnostics)
                diagnostics.Add(new OfficeWorkflowDiagnostic("ScanPreparation", $"Page {page}: {message}", OfficeWorkflowDiagnosticSeverity.Warning, "scan"));
        }
        PdfDocument result = documents.Count == 1 ? documents[0] : PdfDocument.Merge(documents, token);
        byte[] bytes = result.ToBytes();
        if (PdfDocument.Load(bytes).Inspect(options: null, cancellationToken: token).PageCount != pages.Length)
            throw new IOException("The prepared scan page count could not be verified.");
        diagnostics.Add(new OfficeWorkflowDiagnostic("RasterScanCopy", "Saved rendered page appearances only; native text and interactive content are omitted.", OfficeWorkflowDiagnosticSeverity.Warning, "scan"));
        return new OperationArtifact(bytes, $"Created a prepared raster copy containing {pages.Length} page(s).", null);
    }
}