using System.Diagnostics;
using OfficeIMO.Drawing;
using OfficeIMO.Internal;

namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    private static async Task<PdfPageImageExportResult> PublishProviderImagesAsync(ValidatedImageExportRequest request,
        OfficeImageExportBatchSaveResult saved, int[] pageNumbers, long inputBytes, Stopwatch stopwatch,
        List<OfficeWorkflowDiagnostic> diagnostics, CancellationToken token) {
        var batch = await PublishProviderFilesAsync(saved.Files.Select(file => file.Path).ToArray(),
            request.DirectoryOutput!, request.Limits.MaximumOutputBytes, request.PublicationGuard, diagnostics, token).ConfigureAwait(false);
        var files = batch.Files.Select((file, index) => new PdfPageImageFile(pageNumbers[index], file.Path,
            saved.Files[index].Format, saved.Files[index].Width, saved.Files[index].Height, file.SizeBytes)).ToArray();
        long outputBytes = files.Sum(file => file.SizeBytes);
        OfficeWorkflowStatus status = batch.Status;
        string? failure = batch.Failure;
        var recoveries = batch.Recoveries;
        string summary = status == OfficeWorkflowStatus.Completed
            ? $"Exported and verified {files.Length:N0} page images in the provider folder."
            : $"Verified {files.Length:N0} of {saved.Files.Count:N0} page images. Previously verified files remain in the destination. {failure}";
        return new(request.Id, status,
            status is OfficeWorkflowStatus.Completed or OfficeWorkflowStatus.Cancelled ? OfficeWorkflowFailureKind.None : OfficeWorkflowFailureKind.OutputFailed,
            files.Length > 0 ? request.OutputDirectory : null, inputBytes, outputBytes, stopwatch.Elapsed, summary, files, diagnostics, recoveries);
    }

}
