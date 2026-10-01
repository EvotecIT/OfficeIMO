using OfficeIMO.Core.Internal;
using OfficeIMO.Internal;
using OfficeIMO.Invoicing;
using OfficeIMO.Invoicing.Validation;

namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner : IOfficeInvoiceWorkflowRunner {
    /// <summary>Executes shared invoice operations through bounded source snapshots and the existing local/provider publication owner.</summary>
    public async Task<OfficeInvoiceStorageWorkflowResult> RunInvoiceAsync(OfficeInvoiceStorageWorkflowRequest request,
        InvoiceValidator? validator = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(request);
        var diagnostics = new List<OfficeWorkflowDiagnostic>();
        var inputs = new WorkflowInputSnapshots();
        OfficeInvoiceWorkflowResult? workflow = null;
        string? staging = null, providerDirectory = null;
        try {
            cancellationToken.ThrowIfCancellationRequested();
            // Capture host settings before invoking a permission-scoped or asynchronous input factory.
            var operation = request.Operation; var target = request.Target; var release = request.ValidationRelease;
            var edits = request.SourceEdits; var policy = request.ConflictPolicy;
            var inputStream = request.InputStream; var outputStream = request.OutputStream;
            var hostGuard = request.PublicationGuard;
            long maximumInput = request.MaximumInputBytes, maximumOutput = request.MaximumOutputBytes;
            var layout = request.Layout.Clone(); var pdfOptions = request.PdfOptions?.Clone();
            OfficeInvoiceWorkflowRequest.ValidateContract(operation, target, release, edits);
            if (operation != OfficeInvoiceWorkflowOperation.EditSource && edits != null)
                throw new ArgumentException("Source replacements require EditSource.");
            if (!Enum.IsDefined(policy)) throw new ArgumentOutOfRangeException(nameof(request.ConflictPolicy));
            if (maximumInput is <= 0 or > InvoiceProfileDeclaration.MaximumXmlBytes || maximumOutput <= 0)
                throw new ArgumentException("Invoice input must be bounded to 16 MiB and output requires a positive byte limit.");
            bool writes = operation is OfficeInvoiceWorkflowOperation.Convert or OfficeInvoiceWorkflowOperation.EditSource or
                OfficeInvoiceWorkflowOperation.RenderPresentationPdf or OfficeInvoiceWorkflowOperation.RenderHybridPdf;
            string input = ValidateInputLocation(request.InputPath, inputStream);
            string? output = request.OutputPath == null ? null : outputStream == null
                ? ValidateLocalOutput(request.OutputPath) : OfficeStorageIdentity.Normalize(request.OutputPath);
            if (writes != (output != null) || !writes && outputStream != null)
                throw new ArgumentException("Only writing invoice operations require and accept an output destination.");
            if (outputStream != null && policy != OfficeWorkflowConflictPolicy.Replace)
                throw new ArgumentException("Provider output requires Replace after direct-write confirmation.");
            if (writes) {
                string extension = operation is OfficeInvoiceWorkflowOperation.Convert or OfficeInvoiceWorkflowOperation.EditSource ? ".xml" : ".pdf";
                if (!string.Equals(Path.GetExtension(outputStream?.Name ?? output), extension, StringComparison.OrdinalIgnoreCase))
                    throw new ArgumentException("The output filename requires " + extension + " for the selected invoice operation.");
                if (OfficeStorageIdentity.AreEquivalent(input, output!)) throw new IOException("Invoice output must be separate from its source.");
                if (outputStream == null && policy == OfficeWorkflowConflictPolicy.Fail && (File.Exists(output) || Directory.Exists(output)))
                    throw new IOException("Invoice output already exists.");
            }
            inputStream ??= new OfficeWorkflowStreamInput(Path.GetFileName(input), token => {
                token.ThrowIfCancellationRequested();
                return Task.FromResult<Stream>(new FileStream(input, FileMode.Open, FileAccess.Read, FileShare.Read, 81920,
                    FileOptions.Asynchronous | FileOptions.SequentialScan));
            });
            string snapshot = await inputs.CaptureOneAsync(input, inputStream, maximumInput, cancellationToken).ConfigureAwait(false);
            var guard = inputs.Guard(hostGuard, maximumInput, [input], outputStream);
            byte[] xml = await File.ReadAllBytesAsync(snapshot, cancellationToken).ConfigureAwait(false);
            var memory = operation == OfficeInvoiceWorkflowOperation.EditSource
                ? OfficeInvoiceWorkflowRequest.ForSourceEdit(xml, edits!, release, inputStream.Name)
                : new OfficeInvoiceWorkflowRequest(xml, operation, target, release, layout, pdfOptions, inputStream.Name);
            workflow = (await OfficeInvoiceBufferWorkflow.RunBatchAsync([memory], new() {
                MaximumInputBytes = maximumInput, MaximumOutputBytes = maximumOutput
            }, validator, cancellationToken).ConfigureAwait(false))[0];
            if (!workflow.Succeeded) return Result(OfficeWorkflowStatus.Failed, null, "Invoice execution failed; no artifact was published.");
            if (!writes) return Result(OfficeWorkflowStatus.Completed, null, "Invoice operation completed. Check the separate model and standards findings.");
            cancellationToken.ThrowIfCancellationRequested();
            string directory = outputStream == null ? Path.GetDirectoryName(output!)!
                : providerDirectory = OfficeTemporaryDirectory.Create("officeimo-invoice-output-");
            Directory.CreateDirectory(directory);
            staging = Path.Combine(directory, ".invoice-" + Guid.NewGuid().ToString("N") + ".tmp");
            await File.WriteAllBytesAsync(staging, workflow.ToOutputBytes()!, cancellationToken).ConfigureAwait(false);
            inputs.Dispose();
            if (outputStream != null) {
                var outcome = await PublishProviderArtifactAsync(staging, output!, outputStream, maximumOutput, guard, () => {
                    File.Delete(staging); staging = null;
                    Directory.Delete(providerDirectory!, recursive: false); providerDirectory = null;
                }, diagnostics, cancellationToken).ConfigureAwait(false);
                return new(outcome.Status, workflow, outcome.Status == OfficeWorkflowStatus.Completed ? outcome.PublishedLocation : null,
                    outcome.Summary, diagnostics, outcome.Recovery);
            }
            string published = await PublishAsync(staging, output!, policy, guard, cancellationToken).ConfigureAwait(false);
            staging = null;
            diagnostics.Add(new OfficeWorkflowDiagnostic("AtomicPublication", "The invoice artifact was published with one filesystem move.", stage: "publish"));
            return Result(OfficeWorkflowStatus.Completed, published, "Invoice artifact created.");
        } catch (OperationCanceledException error) when (cancellationToken.IsCancellationRequested) {
            ReportInputStagingCleanupFailure(error, diagnostics);
            return Result(OfficeWorkflowStatus.Cancelled, null, "Invoice operation cancelled before publication.");
        } catch (Exception error) when (error is not OutOfMemoryException and not StackOverflowException) {
            ReportInputStagingCleanupFailure(error, diagnostics);
            diagnostics.Add(new OfficeWorkflowDiagnostic("InvoiceStorageWorkflowFailed", error.Message, OfficeWorkflowDiagnosticSeverity.Error));
            return Result(OfficeWorkflowStatus.Failed, null, error.Message);
        } finally {
            inputs.Cleanup(diagnostics);
            if (staging != null) TryDelete(staging);
            if (providerDirectory != null) TryDeleteDirectory(providerDirectory);
        }
        OfficeInvoiceStorageWorkflowResult Result(OfficeWorkflowStatus status, string? path, string summary) =>
            new(status, workflow, path, summary, diagnostics);
    }
}
