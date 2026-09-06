using System.Text;
using OfficeIMO.Core.Internal;
using OfficeIMO.Internal;
using OfficeIMO.Ocr;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Image;

namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    /// <summary>Recognizes an image through Reader.Ocr, reviews the text, and publishes a verified UTF-8 artifact.</summary>
    /// <remarks>The caller owns the engine. Sources are snapshotted and revalidated before publication;
    /// provider outputs use the same explicit direct-write and recovery contract as document workflows.</remarks>
    public async Task<ImageOcrWorkflowResult> RecognizeImageAsync(ImageOcrWorkflowRequest request,
        IOcrEngine engine, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(request);
        ArgumentNullException.ThrowIfNull(engine);
        var diagnostics = new List<OfficeWorkflowDiagnostic>();
        var inputs = new WorkflowInputSnapshots();
        string? staging = null;
        string? temporaryDirectory = null;
        int characters = 0;
        try {
            if (!Enum.IsDefined(request.ConflictPolicy)) throw new ArgumentException("Choose a supported output conflict policy.");
            var policy = request.ConflictPolicy;
            var limits = request.Limits.CloneAndValidate();
            var options = request.Ocr.Clone();
            var callback = request.ReviewAsync;
            var inputStream = request.InputStream;
            var outputStream = request.OutputStream;
            var hostGuard = request.PublicationGuard;
            string input = ValidateInputLocation(request.InputPath, inputStream);
            string output = outputStream is null ? ValidateLocalOutput(request.OutputPath) : OfficeStorageIdentity.Normalize(request.OutputPath);
            string sourceName = inputStream?.Name ?? Path.GetFileName(input);
            if (!string.Equals(Path.GetExtension(outputStream?.Name ?? output), ".txt", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("Choose a .txt output filename.");
            if (outputStream is not null && policy != OfficeWorkflowConflictPolicy.Replace)
                throw new ArgumentException("A provider output requires the Replace policy after direct-write confirmation.");
            if (string.Equals(input, output, StringComparison.Ordinal)) throw new IOException("Choose an output different from the source image.");
            inputStream ??= new OfficeWorkflowStreamInput(sourceName, token => {
                token.ThrowIfCancellationRequested();
                return Task.FromResult<Stream>(new FileStream(input, FileMode.Open, FileAccess.Read, FileShare.Read,
                    81920, FileOptions.Asynchronous | FileOptions.SequentialScan));
            });
            string snapshot = await inputs.CaptureOneAsync(input, inputStream, limits.MaximumInputBytes, cancellationToken).ConfigureAwait(false);
            var guard = inputs.Guard(hostGuard, limits.MaximumInputBytes, [input], outputStream);
            var reader = new OfficeDocumentReaderBuilder().AddImageHandler().Build();
            OfficeDocumentReadResult document;
            await using (var file = File.OpenRead(snapshot)) {
                document = await reader.ReadDocumentAsync(file, sourceName,
                    new ReaderOptions { MaxInputBytes = limits.MaximumInputBytes }, cancellationToken).ConfigureAwait(false);
            }
            OfficeDocumentOcrExecutionResult recognition = await document.ApplyOcrAsync(engine, options, cancellationToken).ConfigureAwait(false);
            foreach (var diagnostic in recognition.Diagnostics) {
                diagnostics.Add(new OfficeWorkflowDiagnostic(diagnostic.Code, diagnostic.Message,
                    diagnostic.Severity == OfficeDocumentDiagnosticSeverity.Error ? OfficeWorkflowDiagnosticSeverity.Error
                        : OfficeWorkflowDiagnosticSeverity.Warning, stage: "recognize-image"));
            }
            if (recognition.Report.FailedCandidateCount > 0 || recognition.Report.SkippedCandidateCount > 0 || recognition.Report.AttemptedCandidateCount == 0)
                throw new InvalidDataException("Image recognition did not complete. Inspect the OCR diagnostics before retrying.");
            string text = string.Join(Environment.NewLine, recognition.Recognitions.Select(item => item.Result.Text ?? string.Empty));
            if (callback is not null) {
                var review = new ImageOcrWorkflowReview(sourceName, document.Assets.Single().PayloadBytes!, recognition, text);
                text = await callback(review, cancellationToken).WaitAsync(cancellationToken).ConfigureAwait(false)
                    ?? throw new InvalidOperationException("OCR review did not return text.");
            }
            cancellationToken.ThrowIfCancellationRequested();
            var encoding = new UTF8Encoding(false, true);
            if (encoding.GetByteCount(text) > limits.MaximumOutputBytes) throw new InvalidDataException("Reviewed text exceeds the output byte limit.");
            byte[] bytes = encoding.GetBytes(text);
            characters = text.Length;
            string directory = outputStream is null ? Path.GetDirectoryName(output)!
                : temporaryDirectory = OfficeTemporaryDirectory.Create("officeimo-image-ocr-output-");
            Directory.CreateDirectory(directory);
            staging = Path.Combine(directory, ".image-ocr-" + Guid.NewGuid().ToString("N") + ".tmp");
            await using (var file = OfficeTemporaryFile.CreateAtPath(staging, 81920, FileOptions.Asynchronous)) {
                await file.WriteAsync(bytes, cancellationToken).ConfigureAwait(false);
            }
            byte[] reopened = await File.ReadAllBytesAsync(staging, cancellationToken).ConfigureAwait(false);
            if (!bytes.AsSpan().SequenceEqual(reopened)) throw new InvalidDataException("The staged text could not be verified.");
            inputs.Dispose();
            if (outputStream is not null) {
                var outcome = await PublishProviderArtifactAsync(staging, output, outputStream, limits.MaximumOutputBytes, guard, () => {
                    File.Delete(staging); staging = null;
                    Directory.Delete(temporaryDirectory!, recursive: false); temporaryDirectory = null;
                }, diagnostics, cancellationToken).ConfigureAwait(false);
                return new(outcome.Status, outcome.Status == OfficeWorkflowStatus.Completed ? outcome.PublishedLocation : null,
                    outcome.Summary, characters, diagnostics, outcome.Recovery);
            }
            string published = await PublishAsync(staging, output, policy, guard, cancellationToken).ConfigureAwait(false);
            staging = null;
            return new(OfficeWorkflowStatus.Completed, published, "Recognized text saved.", characters, diagnostics);
        } catch (OperationCanceledException error) when (cancellationToken.IsCancellationRequested) {
            ReportInputStagingCleanupFailure(error, diagnostics);
            inputs.Cleanup(diagnostics);
            return new(OfficeWorkflowStatus.Cancelled, null, "Image OCR cancelled before publication.", characters, diagnostics);
        } catch (Exception error) when (error is not OutOfMemoryException and not StackOverflowException) {
            ReportInputStagingCleanupFailure(error, diagnostics);
            inputs.Cleanup(diagnostics);
            diagnostics.Add(new OfficeWorkflowDiagnostic("ImageOcrFailed", error.Message, OfficeWorkflowDiagnosticSeverity.Error));
            return new(OfficeWorkflowStatus.Failed, null, error.Message, characters, diagnostics);
        } finally {
            inputs.Cleanup(diagnostics);
            if (staging is not null) TryDelete(staging);
            if (temporaryDirectory is not null) TryDeleteDirectory(temporaryDirectory);
        }
    }
}
