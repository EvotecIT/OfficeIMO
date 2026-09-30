using OfficeIMO.Core.Internal;
using OfficeIMO.Invoicing;
using OfficeIMO.Invoicing.Pdf;
using OfficeIMO.Invoicing.Validation;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows;

/// <summary>A prospective invoice destination conflicts with existing files, input paths or another batch output.</summary>
public sealed class OfficeInvoiceOutputException : IOException {
    /// <summary>Creates an output preflight failure with an explanatory message.</summary>
    public OfficeInvoiceOutputException(string message) : base(message) { }
}

/// <summary>Local-file invoice request. Existing destinations are never overwritten.</summary>
public sealed class OfficeInvoiceFileWorkflowRequest {
    internal readonly InvoicePdfLayoutOptions Layout;
    internal readonly PdfOptions? PdfOptions;
    /// <summary>Captures explicit source, operation, target and rendering settings. Writing requires an output path.</summary>
    public OfficeInvoiceFileWorkflowRequest(string inputPath, OfficeInvoiceWorkflowOperation operation = OfficeInvoiceWorkflowOperation.Inspect,
        string? outputPath = null, InvoiceXmlOptions? target = null, InvoiceSpecificationRelease? validationRelease = null,
        InvoicePdfLayoutOptions? layout = null, PdfOptions? pdfOptions = null) {
        InputPath = Path.GetFullPath(inputPath);
        OutputPath = outputPath == null ? null : Path.GetFullPath(outputPath);
        OfficeInvoiceWorkflowRequest.ValidateContract(operation, target, validationRelease);
        bool writes = operation is OfficeInvoiceWorkflowOperation.Convert or OfficeInvoiceWorkflowOperation.RenderHybridPdf or OfficeInvoiceWorkflowOperation.RenderPresentationPdf;
        if (writes != (OutputPath != null)) throw new ArgumentException("Only writing operations require and accept an output path.", nameof(outputPath));
        Operation = operation; Target = target; ValidationRelease = validationRelease;
        Layout = (layout ?? new()).Clone(); PdfOptions = pdfOptions?.Clone();
    }
    /// <summary>Absolute input path.</summary>
    public string InputPath { get; }
    /// <summary>Absolute output path, when writing.</summary>
    public string? OutputPath { get; }
    /// <summary>Operation to perform.</summary>
    public OfficeInvoiceWorkflowOperation Operation { get; }
    /// <summary>Selected output/inspection contract.</summary>
    public InvoiceXmlOptions? Target { get; }
    /// <summary>Requested exact-byte standards release.</summary>
    public InvoiceSpecificationRelease? ValidationRelease { get; }
}

/// <summary>Invoice execution and local-file publication evidence.</summary>
public sealed record OfficeInvoiceFileWorkflowResult(
    string InputPath, string? OutputPath, OfficeInvoiceWorkflowResult Workflow, bool Published, string? PublicationError) {
    /// <summary>True when execution succeeded and any requested artifact was published.</summary>
    public bool Succeeded => Workflow.Succeeded && PublicationError == null && (OutputPath == null || Published);
}

/// <summary>Thin local-file adapter over the memory workflow and shared atomic file commit owner.</summary>
public static class OfficeInvoiceFileWorkflow {
    /// <summary>Reads bounded XML and atomically creates a new output file after successful execution.</summary>
    public static async Task<OfficeInvoiceFileWorkflowResult> RunAsync(OfficeInvoiceFileWorkflowRequest request,
        InvoiceValidator? validator = null, CancellationToken cancellationToken = default) =>
        (await RunBatchAsync([request], validator: validator, cancellationToken: cancellationToken).ConfigureAwait(false))[0];

    /// <summary>Preflights and captures all bounded inputs before sequential execution. Outputs cannot overwrite existing files, sources or another planned output. Publication is per item, not a batch transaction.</summary>
    public static async Task<IReadOnlyList<OfficeInvoiceFileWorkflowResult>> RunBatchAsync(IEnumerable<OfficeInvoiceFileWorkflowRequest> requests,
        OfficeInvoiceWorkflowBatchOptions? options = null, InvoiceValidator? validator = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(requests);
        cancellationToken.ThrowIfCancellationRequested();
        var limits = (options ?? new()).Snapshot();
        var captured = new List<OfficeInvoiceFileWorkflowRequest>();
        var outputs = new SortedSet<string>(StringComparer.Ordinal);
        foreach (var request in requests) {
            cancellationToken.ThrowIfCancellationRequested();
            ArgumentNullException.ThrowIfNull(request);
            if (captured.Count >= limits.MaximumRequests) throw new ArgumentException("Invoice batch exceeds the request limit.", nameof(requests));
            if (request.OutputPath != null) {
                string identity = OfficeWorkflowPathIdentity.Normalize(request.OutputPath);
                if (outputs.Contains(identity) || OfficeWorkflowPathIdentity.TryFindAncestorOrDescendant(identity, outputs, out _))
                    throw new OfficeInvoiceOutputException("Invoice batch outputs collide.");
                outputs.Add(identity);
                if (File.Exists(request.OutputPath) || Directory.Exists(request.OutputPath)) throw new OfficeInvoiceOutputException("Invoice output already exists: " + request.OutputPath);
            }
            captured.Add(request);
        }
        var inputs = new SortedSet<string>(captured.Select(r => OfficeWorkflowPathIdentity.Normalize(r.InputPath)), StringComparer.Ordinal);
        foreach (string output in outputs) {
            cancellationToken.ThrowIfCancellationRequested();
            if (inputs.Contains(output) || OfficeWorkflowPathIdentity.TryFindAncestorOrDescendant(output, inputs, out _))
                throw new OfficeInvoiceOutputException("Invoice output overlaps a source path: " + output);
        }
        var memory = new List<OfficeInvoiceWorkflowRequest>();
        long inputBytes = 0;
        foreach (var request in captured) {
            cancellationToken.ThrowIfCancellationRequested();
            using var stream = new FileStream(request.InputPath, FileMode.Open, FileAccess.Read, FileShare.Read);
            long length = stream.Length;
            if (length == 0 || length > InvoiceProfileDeclaration.MaximumXmlBytes || length > limits.MaximumInputBytes - inputBytes)
                throw new InvalidDataException("Invoice input exceeds the individual or combined input limit.");
            byte[] xml = new byte[(int)length];
            await stream.ReadExactlyAsync(xml.AsMemory(), cancellationToken).ConfigureAwait(false);
            if (stream.ReadByte() != -1) throw new InvalidDataException("Invoice input changed while being captured.");
            inputBytes += length;
            memory.Add(new(xml, request.Operation, request.Target, request.ValidationRelease, request.Layout, request.PdfOptions, request.InputPath));
        }
        var executions = await OfficeInvoiceBufferWorkflow.RunBatchAsync(memory, limits, validator, cancellationToken).ConfigureAwait(false);
        var results = new List<OfficeInvoiceFileWorkflowResult>();
        for (int i = 0; i < executions.Count; i++) {
            var request = captured[i]; var execution = executions[i];
            bool published = false; string? error = null;
            if (execution.Succeeded && request.OutputPath != null) {
                cancellationToken.ThrowIfCancellationRequested();
                try {
                    await OfficeFileCommit.WriteAllBytesAsync(request.OutputPath, execution.ToOutputBytes()!,
                        OfficeFileCommit.ConflictPolicy.FailIfExists, cancellationToken).ConfigureAwait(false);
                    published = true;
                } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) { error = exception.Message; }
            }
            results.Add(new(request.InputPath, request.OutputPath, execution, published, error));
            if (error != null && !limits.ContinueOnFailure) break;
        }
        return results.AsReadOnly();
    }
}
