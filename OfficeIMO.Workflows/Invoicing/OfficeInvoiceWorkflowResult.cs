using OfficeIMO.Invoicing;
using OfficeIMO.Invoicing.Validation;

namespace OfficeIMO.Workflows;

/// <summary>Operation evidence kept separate from standards compliance, bound to captured input and exact returned artifacts.</summary>
public sealed class OfficeInvoiceWorkflowResult {
    private readonly byte[]? _output;
    private readonly byte[]? _outputXml;
    internal OfficeInvoiceWorkflowResult(OfficeInvoiceWorkflowRequest request, string hash, bool succeeded,
        InvoiceReadResult? source, InvoiceModelValidationResult? model, IReadOnlyList<InvoiceDiagnostic> diagnostics,
        InvoiceValidationReport? standards = null, byte[]? output = null, byte[]? outputXml = null) {
        Operation = request.Operation; InputName = request.InputName; InputSha256 = hash; InputByteLength = request.InputByteLength;
        Succeeded = succeeded; Source = source; ModelValidation = model; Diagnostics = diagnostics;
        StandardsValidation = standards; _output = output; _outputXml = outputXml;
    }
    /// <summary>Operation performed.</summary>
    public OfficeInvoiceWorkflowOperation Operation { get; }
    /// <summary>Caller-supplied input identifier.</summary>
    public string? InputName { get; }
    /// <summary>SHA-256 of the exact captured input.</summary>
    public string InputSha256 { get; }
    /// <summary>Length of captured input bytes.</summary>
    public int InputByteLength { get; }
    /// <summary>True when the selected operation completed. Inspection completion does not mean model or standards validity.</summary>
    public bool Succeeded { get; }
    /// <summary>Parsed source and its mapping findings. Its editable model belongs to this result, not the request or caller.</summary>
    public InvoiceReadResult? Source { get; }
    /// <summary>Semantic validation of the source, including recognized aggregate-only profiles.</summary>
    public InvoiceModelValidationResult? ModelValidation { get; }
    /// <summary>Bounded model, mapping, target and execution findings.</summary>
    public IReadOnlyList<InvoiceDiagnostic> Diagnostics { get; }
    /// <summary>Exact-byte standards report, or null when not requested/configured. For writing operations it validates output XML before output is returned.</summary>
    public InvoiceValidationReport? StandardsValidation { get; }
    /// <summary>Explicit standards schema status; NotRun when no standards report exists.</summary>
    public InvoiceValidationStatus SchemaStatus => StandardsValidation?.SchemaStatus ?? InvoiceValidationStatus.NotRun;
    /// <summary>Explicit business-rule status; NotRun when no standards report exists.</summary>
    public InvoiceValidationStatus BusinessRulesStatus => StandardsValidation?.BusinessRulesStatus ?? InvoiceValidationStatus.NotRun;
    /// <summary>Length of the generated primary artifact, or zero.</summary>
    public int OutputByteLength => _output?.Length ?? 0;
    /// <summary>Returns a defensive copy of converted XML or rendered PDF. Failure has no output bytes.</summary>
    public byte[]? ToOutputBytes() => _output == null ? null : (byte[])_output.Clone();
    /// <summary>Returns the XML used for conversion or PDF capture. These exact bytes are validated when standards validation is requested.</summary>
    public byte[]? ToOutputXmlBytes() => _outputXml == null ? null : (byte[])_outputXml.Clone();
    internal long RetainedOutputBytes => (_output?.LongLength ?? 0) + (ReferenceEquals(_output, _outputXml) ? 0 : _outputXml?.LongLength ?? 0);
}
