using OfficeIMO.Core.Internal;
using OfficeIMO.Internal;
using OfficeIMO.Ocr;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;

namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    /// <summary>Recognizes an immutable source snapshot, reopens the generated PDF, and publishes after source and host checks.</summary>
    /// <remarks>The caller owns the OCR engine lifetime. Local inputs are fingerprinted and checked for physical replacement;
    /// provider inputs are reopened before publication. Provider outputs require an explicit direct-write contract.</remarks>
    public async Task<PdfSearchableWorkflowResult> MakePdfSearchableAsync(PdfSearchableWorkflowRequest request,
        IOcrEngine engine, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(request);
        ArgumentNullException.ThrowIfNull(engine);
        var diagnostics = new List<OfficeWorkflowDiagnostic>();
        var inputs = new WorkflowInputSnapshots();
        string? stagingPath = null;
        string? providerDirectory = null;
        int words = 0;
        IReadOnlyList<int> pages = Array.Empty<int>();
        string? providerName = null;
        try {
            if (string.IsNullOrWhiteSpace(request.Id)) throw new ArgumentException("Request id cannot be empty.");
            if (!Enum.IsDefined(request.ConflictPolicy)) throw new ArgumentException("Choose a supported output conflict policy.");
            var policy = request.ConflictPolicy;
            var limits = request.Limits.CloneAndValidate();
            var options = request.Ocr.Clone();
            var password = request.PdfPassword;
            var inputStream = request.InputStream;
            var outputStream = request.OutputStream;
            string input = ValidateInputLocation(request.InputPath, inputStream);
            string output = outputStream is null ? ValidateLocalOutput(request.OutputPath) : OfficeStorageIdentity.Normalize(request.OutputPath);
            EnsurePdfExtension(inputStream?.Name ?? input);
            EnsurePdfExtension(outputStream?.Name ?? output);
            if (outputStream is not null && policy != OfficeWorkflowConflictPolicy.Replace)
                throw new ArgumentException("A provider output requires the Replace policy after direct-write confirmation.");
            if (OfficeStorageIdentity.AreEquivalent(input, output)) throw new IOException("Choose an output different from the source PDF.");
            IOfficeWorkflowPublicationGuard sourceGuard = new ProviderSourcePublicationGuard(request.PublicationGuard, [input]);
            if (OfficeStorageIdentity.GetLocalPath(input) is { } localInput) {
                string identity = OfficePathIdentity.GetPhysicalIdentityKey(localInput);
                sourceGuard = new OcrLocalSourceGuard(sourceGuard, localInput, identity);
                inputStream ??= new OfficeWorkflowStreamInput(Path.GetFileName(localInput), token => {
                    token.ThrowIfCancellationRequested();
                    if (OfficePathIdentity.GetPhysicalIdentityKey(localInput) != identity) throw new IOException("The source PDF was replaced during OCR.");
                    return Task.FromResult<Stream>(new FileStream(localInput, FileMode.Open, FileAccess.Read, FileShare.Read,
                        81920, FileOptions.Asynchronous | FileOptions.SequentialScan));
                });
            }
            string snapshot = await inputs.CaptureOneAsync(input, inputStream, limits.MaximumInputBytes, cancellationToken).ConfigureAwait(false);
            IOfficeWorkflowPublicationGuard? guard = inputs.Guard(sourceGuard, limits.MaximumInputBytes);
            var loadOptions = CreatePdfLoadOptions(password, limits.MaximumInputBytes);
            PdfDocument source = await PdfDocument.LoadAsync(snapshot, loadOptions, cancellationToken).ConfigureAwait(false);
            options.SourceName ??= inputStream!.Name;
            PdfSearchableOcrResult recognized = await source.MakeSearchableAsync(engine, options, cancellationToken).ConfigureAwait(false);
            words = recognized.AddedWordCount;
            pages = recognized.ModifiedPages;
            providerName = recognized.Ocr.Pages.Select(page => page.Provider).FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
            string directory = outputStream is null ? Path.GetDirectoryName(output)!
                : providerDirectory = OfficeTemporaryDirectory.Create("officeimo-ocr-output-");
            Directory.CreateDirectory(directory);
            stagingPath = Path.Combine(directory, ".ocr-" + Guid.NewGuid().ToString("N") + ".tmp");
            await using (var file = new FileStream(stagingPath, FileMode.CreateNew, FileAccess.Write, FileShare.None, 81920, FileOptions.Asynchronous))
            await using (var bounded = new OfficeWorkflowBoundedWriteStream(file, limits.MaximumOutputBytes, leaveOpen: false)) {
                await recognized.Document.SaveAsync(bounded, cancellationToken).ConfigureAwait(false);
            }
            var outputLoadOptions = CreatePdfLoadOptions(password, limits.MaximumOutputBytes);
            PdfDocument reopened = await PdfDocument.LoadAsync(stagingPath, outputLoadOptions, cancellationToken).ConfigureAwait(false);
            if (reopened.Inspect(outputLoadOptions, cancellationToken).PageCount != source.Inspect(loadOptions, cancellationToken).PageCount)
                throw new InvalidDataException("The searchable PDF did not preserve the source page count.");
            diagnostics.Add(new OfficeWorkflowDiagnostic("SearchablePdfReopened", "The staged searchable PDF was reopened and its page count verified.", stage: "validate-output"));
            inputs.Dispose();
            if (outputStream is not null) {
                var outcome = await PublishProviderArtifactAsync(stagingPath, output, outputStream, limits.MaximumOutputBytes, guard, () => {
                    File.Delete(stagingPath);
                    stagingPath = null;
                    Directory.Delete(providerDirectory!, recursive: false);
                    providerDirectory = null;
                }, diagnostics, cancellationToken).ConfigureAwait(false);
                return new PdfSearchableWorkflowResult(outcome.Status, outcome.Status == OfficeWorkflowStatus.Completed ? output : null,
                    outcome.Summary, words, pages, providerName, diagnostics, outcome.Recovery);
            }
            string published = await PublishAsync(stagingPath, output, policy, guard, cancellationToken).ConfigureAwait(false);
            stagingPath = null;
            return new PdfSearchableWorkflowResult(OfficeWorkflowStatus.Completed, published, "Searchable PDF created.", words, pages, providerName, diagnostics);
        } catch (OperationCanceledException error) when (cancellationToken.IsCancellationRequested) {
            ReportInputStagingCleanupFailure(error, diagnostics);
            inputs.Cleanup(diagnostics);
            return new PdfSearchableWorkflowResult(OfficeWorkflowStatus.Cancelled, null, "OCR cancelled before publication.", words, pages, providerName, diagnostics);
        } catch (Exception error) when (error is not OutOfMemoryException and not StackOverflowException) {
            ReportInputStagingCleanupFailure(error, diagnostics);
            inputs.Cleanup(diagnostics);
            diagnostics.Add(new OfficeWorkflowDiagnostic("SearchablePdfFailed", error.Message, OfficeWorkflowDiagnosticSeverity.Error));
            return new PdfSearchableWorkflowResult(OfficeWorkflowStatus.Failed, null, error.Message, words, pages, providerName, diagnostics);
        } finally {
            inputs.Cleanup(diagnostics);
            if (stagingPath is not null) TryDelete(stagingPath);
            if (providerDirectory is not null) TryDeleteDirectory(providerDirectory);
        }
    }

    private sealed class OcrLocalSourceGuard(IOfficeWorkflowPublicationGuard host, string source, string identity) : IOfficeWorkflowPublicationGuard {
        public async ValueTask<bool> CanPublishAsync(string path, bool isDirectory, CancellationToken token) {
            if (!await host.CanPublishAsync(path, isDirectory, token).ConfigureAwait(false)) return false;
            if (OfficePathIdentity.GetPhysicalIdentityKey(source) != identity) throw new IOException("The source PDF was replaced during OCR.");
            return !File.Exists(path) || OfficePathIdentity.GetPhysicalIdentityKey(path) != identity;
        }
    }
}
