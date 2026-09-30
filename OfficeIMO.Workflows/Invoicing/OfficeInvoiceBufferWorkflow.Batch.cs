using OfficeIMO.Invoicing;
using OfficeIMO.Invoicing.Validation;

namespace OfficeIMO.Workflows;

public static partial class OfficeInvoiceBufferWorkflow {
    /// <summary>Preflights bounded immutable requests, then processes in input order. No item is executed when preflight exceeds a batch limit.</summary>
    public static async Task<IReadOnlyList<OfficeInvoiceWorkflowResult>> RunBatchAsync(IEnumerable<OfficeInvoiceWorkflowRequest> requests,
        OfficeInvoiceWorkflowBatchOptions? options = null, InvoiceValidator? validator = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(requests);
        OfficeInvoiceWorkflowBatchOptions limits = (options ?? new OfficeInvoiceWorkflowBatchOptions()).Snapshot();
        cancellationToken.ThrowIfCancellationRequested();
        var captured = new List<OfficeInvoiceWorkflowRequest>();
        long inputBytes = 0;
        foreach (OfficeInvoiceWorkflowRequest request in requests) {
            cancellationToken.ThrowIfCancellationRequested();
            if (request == null) throw new ArgumentException("Invoice batches cannot contain null requests.", nameof(requests));
            if (captured.Count == limits.MaximumRequests) throw new InvalidDataException("Invoice batch exceeds its request limit.");
            if (request.InputByteLength > limits.MaximumInputBytes - inputBytes) throw new InvalidDataException("Invoice batch exceeds its combined input-byte limit.");
            inputBytes += request.InputByteLength;
            captured.Add(request);
        }
        var results = new List<OfficeInvoiceWorkflowResult>(captured.Count);
        long outputBytes = 0;
        foreach (OfficeInvoiceWorkflowRequest request in captured) {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeInvoiceWorkflowResult result = await RunAsync(request, validator, cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            if (result.RetainedOutputBytes > limits.MaximumOutputBytes - outputBytes) {
                var diagnostics = new OperationDiagnostics();
                diagnostics.AddRange(result.Diagnostics);
                diagnostics.Add(new InvoiceDiagnostic("INV-WORKFLOW-BATCH-OUTPUT", "This artifact would exceed the batch's combined output-byte limit; no bytes from this item are returned.", "Output"));
                result = new OfficeInvoiceWorkflowResult(request, result.InputSha256, false, result.Source, result.ModelValidation, diagnostics.ToList(), result.StandardsValidation);
            } else outputBytes += result.RetainedOutputBytes;
            results.Add(result);
            if (!result.Succeeded && !limits.ContinueOnFailure) break;
        }
        return results.AsReadOnly();
    }
}
