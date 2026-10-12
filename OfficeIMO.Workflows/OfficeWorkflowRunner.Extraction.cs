using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    private static OperationArtifact ExtractPages(ValidatedRequest request, CancellationToken token) {
        byte[] input = ReadInput(request.InputPath, request.Limits, token);
        token.ThrowIfCancellationRequested();
        var document = PdfDocument.Load(input, request.PdfLoadOptions);
        int[] pages = request.PageSelector?.Resolve(document.Inspect(request.PdfLoadOptions, token).PageCount, request.MaximumExtractedPages).ToArray()
            ?? request.PageNumbers!;
        var extracted = document.Pages.Extract(pages, request.Limits.MaximumOutputBytes);
        token.ThrowIfCancellationRequested();
        using var output = new OfficeWorkflowBoundedMemoryStream(request.Limits.MaximumOutputBytes);
        extracted.SaveAsync(output, token).GetAwaiter().GetResult();
        token.ThrowIfCancellationRequested();
        byte[] bytes = output.ToArray();
        if (PdfDocument.Load(bytes, request.OutputPdfLoadOptions).Inspect().PageCount != pages.Length)
            throw new InvalidDataException("The saved extraction does not contain the requested number of pages.");
        return new OperationArtifact(bytes, $"Extracted {pages.Length:N0} pages into a separate PDF.", null);
    }
}
