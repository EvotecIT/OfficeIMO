using OfficeIMO.Internal;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    /// <summary>Splits consecutive pages through the PDF engine and publishes reopened, bounded outputs.</summary>
    public async Task<PdfSplitWorkflowResult> SplitPdfAsync(PdfSplitWorkflowRequest request,
        IProgress<OfficeWorkflowProgress>? progress = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(request);
        string id = request.Id;
        var diagnostics = new List<OfficeWorkflowDiagnostic>();
        var inputs = new WorkflowInputSnapshots();
        string? staging = null;
        try {
            if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("Request id cannot be empty.");
            int pagesPerPart = request.PagesPerDocument;
            int maximumParts = request.MaximumParts;
            if (pagesPerPart < 1 || maximumParts < 1) throw new ArgumentOutOfRangeException(nameof(request), "Page and part limits must be positive.");
            var policy = request.ConflictPolicy;
            if (!Enum.IsDefined(policy)) throw new ArgumentOutOfRangeException(nameof(request.ConflictPolicy));
            var limits = (request.Limits ?? throw new ArgumentException("Workflow limits are required.")).CloneAndValidate();
            var directory = request.DirectoryOutput;
            var sourceStream = request.InputStream;
            string source = ValidateInputLocation(request.InputPath, sourceStream);
            string name = sourceStream?.Name ?? Path.GetFileName(source);
            EnsurePdfExtension(name);
            string output = directory is null
                ? Path.TrimEndingDirectorySeparator(ValidateLocalOutput(request.OutputDirectory))
                : OfficeStorageIdentity.Normalize(request.OutputDirectory);
            if (directory is not null && policy != OfficeWorkflowConflictPolicy.Replace)
                throw new ArgumentException("Provider folder publication requires explicit Replace semantics.");
            if (directory is null && string.IsNullOrEmpty(Path.GetDirectoryName(output)))
                throw new ArgumentException("Output directory cannot be a filesystem root.");
            var loadOptions = CreatePdfLoadOptions(request.PdfPassword, limits.MaximumInputBytes);
            var hostGuard = request.PublicationGuard;
            sourceStream ??= new OfficeWorkflowStreamInput(name,
                token => { token.ThrowIfCancellationRequested(); return Task.FromResult<Stream>(File.OpenRead(source)); });
            Report(progress, id, "input", "Reading the source PDF", 0.05);
            string captured = await inputs.CaptureOneAsync(source, sourceStream, limits.MaximumInputBytes, cancellationToken).ConfigureAwait(false);
            var guard = inputs.Guard(hostGuard, limits.MaximumInputBytes, [source]);
            PdfDocument document = await PdfDocument.LoadAsync(captured, loadOptions, cancellationToken).ConfigureAwait(false);
            int pageCount = document.Inspect(loadOptions, cancellationToken).PageCount;
            if (pageCount < 1) throw new InvalidDataException("The source PDF has no pages.");
            PdfSplitPlan plan = PdfSplitPlan.Create(pageCount, pagesPerPart, maximumParts);
            int count = plan.Parts.Count;
            cancellationToken.ThrowIfCancellationRequested();
            if (directory is null) {
                string parent = Path.GetDirectoryName(output)!;
                Directory.CreateDirectory(parent);
                staging = Path.Combine(parent, ".officeimo-split." + Guid.NewGuid().ToString("N") + ".tmp");
                Directory.CreateDirectory(staging);
            } else staging = OfficeIMO.Core.Internal.OfficeTemporaryDirectory.Create("officeimo-split-");
            var staged = new List<PdfSplitFile>(count);
            long totalBytes = 0;
            // Use a stable portable name; provider display names are not filesystem-safe on every host.
            for (int index = 0; index < count; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                PdfSplitPart planned = plan.Parts[index];
                int first = planned.FirstSourcePage;
                int expected = planned.PageCount;
                string path = Path.Combine(staging, planned.Name);
                Report(progress, id, "split", $"Preparing part {index + 1} of {count}", 0.1 + 0.65 * index / count);
                cancellationToken.ThrowIfCancellationRequested();
                // Keep the PDF engine's split policy while retaining only one part at a time.
                long remainingOutputBytes = limits.MaximumOutputBytes - totalBytes;
                if (remainingOutputBytes <= 0) throw new InvalidOperationException("The split outputs exceed the aggregate output byte limit.");
                PdfDocument part = document.Pages.Split(PdfPageRange.From(first, first + expected - 1), remainingOutputBytes);
                await using (var file = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                await using (var bounded = new OfficeWorkflowBoundedWriteStream(file, remainingOutputBytes, leaveOpen: true)) {
                    await part.SaveAsync(bounded, cancellationToken).ConfigureAwait(false);
                }
                long size = new FileInfo(path).Length;
                totalBytes = checked(totalBytes + size);
                if (totalBytes > limits.MaximumOutputBytes) throw new InvalidOperationException("The split outputs exceed the aggregate output byte limit.");
                var reopened = await PdfDocument.LoadAsync(path, cancellationToken: cancellationToken).ConfigureAwait(false);
                if (reopened.Inspect().PageCount != expected) throw new InvalidDataException("A staged split PDF has an unexpected page count.");
                staged.Add(new(path, first, expected, size));
            }
            diagnostics.Add(new("SplitPdfsReopened", "Every staged PDF was reopened and its page count verified.", stage: "validate-output"));
            // Guards retain independent source fingerprints after private input cleanup.
            inputs.Dispose();
            Report(progress, id, "publish", "Publishing validated PDF parts", 0.9);
            if (directory is not null) {
                var batch = await PublishProviderFilesAsync(staged.Select(file => file.Path).ToArray(), directory,
                    limits.MaximumOutputBytes, guard, diagnostics, cancellationToken).ConfigureAwait(false);
                var files = batch.Files.Select((file, index) => staged[index] with { Path = file.Path, SizeBytes = file.SizeBytes });
                string summary = batch.Status == OfficeWorkflowStatus.Completed
                    ? $"Created and verified {batch.Files.Count} PDF parts."
                    : $"Verified {batch.Files.Count} of {count} PDF parts. Previously verified files remain in the destination. {batch.Failure}";
                return new(id, batch.Status, summary, files, diagnostics, batch.Recoveries);
            }
            string published = await PublishDirectoryAsync(staging, output, policy, diagnostics, cancellationToken, guard).ConfigureAwait(false);
            staging = null;
            Report(progress, id, "complete", "Split PDFs are ready", 1);
            return new(id, OfficeWorkflowStatus.Completed, $"Created {count} PDF parts.",
                staged.Select(file => file with { Path = Path.Combine(published, Path.GetFileName(file.Path)) }), diagnostics);
        } catch (Exception error) when (error is not OutOfMemoryException and not StackOverflowException) {
            var status = error is OperationCanceledException && cancellationToken.IsCancellationRequested
                ? OfficeWorkflowStatus.Cancelled : OfficeWorkflowStatus.Failed;
            diagnostics.Add(new("PdfSplitFailed", error.Message, status == OfficeWorkflowStatus.Cancelled
                ? OfficeWorkflowDiagnosticSeverity.Information : OfficeWorkflowDiagnosticSeverity.Error,
                details: CreateFailureDetails(error)));
            return new(id, status, error.Message, [], diagnostics);
        } finally {
            inputs.Cleanup(diagnostics);
            if (staging is not null) TryDeleteDirectory(staging);
        }
    }
}
