using OfficeIMO.Invoicing;
using OfficeIMO.Invoicing.Pdf;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows;

/// <summary>Invoice operations composed from the owning XML, standards and PDF libraries.</summary>
public enum OfficeInvoiceWorkflowOperation {
    /// <summary>Inspect the declaration, mapped model, monetary checks and optional target without writing.</summary>
    Inspect,
    /// <summary>Validate the source model and, when configured explicitly, its exact bytes against standards rules.</summary>
    Validate,
    /// <summary>Convert only when all observed source data and the selected target can be represented.</summary>
    Convert,
    /// <summary>Render a separate presentation PDF from a qualified CII projection.</summary>
    RenderPresentationPdf,
    /// <summary>Render a hybrid PDF containing the same captured CII XML as its visible invoice.</summary>
    RenderHybridPdf,
    /// <summary>Replace selected existing source headers while retaining unmapped XML content.</summary>
    EditSource
}

/// <summary>Captured memory-only request. The workflow never follows invoice URLs or writes files.</summary>
public sealed class OfficeInvoiceWorkflowRequest {
    private readonly byte[] _xml;
    private readonly InvoicePdfLayoutOptions _layout;
    private readonly PdfOptions? _pdfOptions;

    /// <summary>Captures bounded XML and independent rendering settings. Conversion and rendering require an explicit target; standards validation requires an explicit release.</summary>
    public OfficeInvoiceWorkflowRequest(byte[] xml, OfficeInvoiceWorkflowOperation operation = OfficeInvoiceWorkflowOperation.Inspect,
        InvoiceXmlOptions? target = null, InvoiceSpecificationRelease? validationRelease = null,
        InvoicePdfLayoutOptions? layout = null, PdfOptions? pdfOptions = null, string? inputName = null)
        : this(xml, operation, target, validationRelease, layout, pdfOptions, inputName, null) { }

    private OfficeInvoiceWorkflowRequest(byte[] xml, OfficeInvoiceWorkflowOperation operation,
        InvoiceXmlOptions? target, InvoiceSpecificationRelease? validationRelease,
        InvoicePdfLayoutOptions? layout, PdfOptions? pdfOptions, string? inputName, InvoiceSourceEdits? sourceEdits) {
        ArgumentNullException.ThrowIfNull(xml);
        if (xml.Length == 0 || xml.Length > InvoiceProfileDeclaration.MaximumXmlBytes)
            throw new ArgumentException("Invoice XML must contain between one byte and 16 MiB.", nameof(xml));
        ValidateContract(operation, target, validationRelease, sourceEdits);
        if (inputName?.Length > 4096) throw new ArgumentException("Input name exceeds 4,096 characters.", nameof(inputName));
        _xml = (byte[])xml.Clone();
        _layout = (layout ?? new InvoicePdfLayoutOptions()).Clone();
        _pdfOptions = pdfOptions?.Clone();
        Operation = operation; Target = target; ValidationRelease = validationRelease; InputName = inputName; SourceEdits = sourceEdits;
    }
    /// <summary>Captures preservation-aware header replacements. Completion does not establish model validity; requested standards checks validate exact edited XML and must pass before output is returned.</summary>
    public static OfficeInvoiceWorkflowRequest ForSourceEdit(byte[] xml, InvoiceSourceEdits edits,
        InvoiceSpecificationRelease? validationRelease = null, string? inputName = null) {
        ArgumentNullException.ThrowIfNull(edits);
        return new(xml, OfficeInvoiceWorkflowOperation.EditSource, null, validationRelease, null, null, inputName, edits);
    }

    internal static void ValidateContract(OfficeInvoiceWorkflowOperation operation, InvoiceXmlOptions? target, InvoiceSpecificationRelease? validationRelease, InvoiceSourceEdits? sourceEdits = null) {
        if (!Enum.IsDefined(operation)) throw new ArgumentOutOfRangeException(nameof(operation));
        if (operation == OfficeInvoiceWorkflowOperation.EditSource && (sourceEdits == null || target != null))
            throw new ArgumentException("Source editing requires captured edits and retains its original syntax/profile. Use ForSourceEdit.", nameof(sourceEdits));
        if (operation is OfficeInvoiceWorkflowOperation.Convert or OfficeInvoiceWorkflowOperation.RenderPresentationPdf or OfficeInvoiceWorkflowOperation.RenderHybridPdf && target == null)
            throw new ArgumentException("Writing invoice output requires an explicit target contract.", nameof(target));
        if (operation is OfficeInvoiceWorkflowOperation.RenderPresentationPdf or OfficeInvoiceWorkflowOperation.RenderHybridPdf && target!.Syntax != InvoiceSyntax.Cii)
            throw new ArgumentException("Invoice PDF rendering requires a CII target contract.", nameof(target));
        if (validationRelease.HasValue && !Enum.IsDefined(validationRelease.Value)) throw new ArgumentOutOfRangeException(nameof(validationRelease));
        if (operation is OfficeInvoiceWorkflowOperation.Convert or OfficeInvoiceWorkflowOperation.RenderPresentationPdf or OfficeInvoiceWorkflowOperation.RenderHybridPdf &&
            validationRelease.HasValue && validationRelease != target!.Release)
            throw new ArgumentException("Output standards validation must use the selected target release.", nameof(validationRelease));
    }
    /// <summary>Selected operation.</summary>
    public OfficeInvoiceWorkflowOperation Operation { get; }
    /// <summary>Explicit output or inspection target, if supplied.</summary>
    public InvoiceXmlOptions? Target { get; }
    /// <summary>Explicit standards release. Null requests model checks without standards validation.</summary>
    public InvoiceSpecificationRelease? ValidationRelease { get; }
    /// <summary>Immutable source replacements, supplied only by ForSourceEdit.</summary>
    public InvoiceSourceEdits? SourceEdits { get; }
    /// <summary>Caller-supplied name for batch identification; never interpreted as a path.</summary>
    public string? InputName { get; }
    /// <summary>Length of captured input bytes.</summary>
    public int InputByteLength => _xml.Length;
    /// <summary>Returns an independent copy of the captured XML.</summary>
    public byte[] ToInputBytes() => (byte[])_xml.Clone();
    internal byte[] Xml => _xml;
    internal InvoicePdfLayoutOptions Layout => _layout;
    internal PdfOptions? PdfOptions => _pdfOptions;
}

/// <summary>Bounds the retained input and output of sequential invoice batch processing.</summary>
public sealed class OfficeInvoiceWorkflowBatchOptions {
    /// <summary>Maximum requests, from one to 10,000. Defaults to 256.</summary>
    public int MaximumRequests { get; set; } = 256;
    /// <summary>Maximum combined input bytes; defaults to 64 MiB.</summary>
    public long MaximumInputBytes { get; set; } = 64L * 1024 * 1024;
    /// <summary>Maximum combined returned artifact bytes; defaults to 64 MiB.</summary>
    public long MaximumOutputBytes { get; set; } = 64L * 1024 * 1024;
    /// <summary>Whether a failed item permits the next item to run. Defaults to true.</summary>
    public bool ContinueOnFailure { get; set; } = true;

    internal OfficeInvoiceWorkflowBatchOptions Snapshot() {
        if (MaximumRequests is < 1 or > 10000) throw new ArgumentOutOfRangeException(nameof(MaximumRequests));
        if (MaximumInputBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumInputBytes));
        if (MaximumOutputBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumOutputBytes));
        return new OfficeInvoiceWorkflowBatchOptions { MaximumRequests = MaximumRequests, MaximumInputBytes = MaximumInputBytes,
            MaximumOutputBytes = MaximumOutputBytes, ContinueOnFailure = ContinueOnFailure };
    }
}
