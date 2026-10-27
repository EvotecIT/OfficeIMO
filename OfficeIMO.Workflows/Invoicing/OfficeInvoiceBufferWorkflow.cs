using OfficeIMO.Internal;
using System.Security.Cryptography;
using System.Xml;
using OfficeIMO.Invoicing;
using OfficeIMO.Invoicing.Pdf;
#if NET8_0_OR_GREATER
using OfficeIMO.Invoicing.Validation;
#endif

namespace OfficeIMO.Workflows;

/// <summary>Memory-only invoice workflows. Hosts own input acquisition and safe publication; format owners own semantics and rendering.</summary>
public static partial class OfficeInvoiceBufferWorkflow {
    /// <summary>Runs captured input through shared owners. A requested standards stage fails closed when its validator is unavailable. Cancellation returns no artifact.</summary>
    public static Task<OfficeInvoiceWorkflowResult> RunAsync(OfficeInvoiceWorkflowRequest request,

#if NET8_0_OR_GREATER
        InvoiceValidator? validator = null,
#endif
        CancellationToken cancellationToken = default) {
        if (request == null) throw new ArgumentNullException(nameof(request));
        cancellationToken.ThrowIfCancellationRequested();
        return Task.Run(() => ProcessAsync(request,
#if NET8_0_OR_GREATER
            validator,
#endif
            cancellationToken), cancellationToken);
    }

    private static async Task<OfficeInvoiceWorkflowResult> ProcessAsync(OfficeInvoiceWorkflowRequest request,

#if NET8_0_OR_GREATER
        InvoiceValidator? validator,
#endif
        CancellationToken token) {
        if (request.Operation == OfficeInvoiceWorkflowOperation.EditSource)
            return await ProcessSourceEditAsync(request,
#if NET8_0_OR_GREATER
            validator,
#endif
            token).ConfigureAwait(false);
        string hash = OfficeRuntimeCompatibility.ToHexString(OfficeRuntimeCompatibility.HashSha256(request.Xml));
        var diagnostics = new OperationDiagnostics();
        InvoiceReadResult? source = null;
        InvoiceModelValidationResult? model = null;
#if NET8_0_OR_GREATER
        InvoiceValidationReport? standards = null;
#endif
        try {
            token.ThrowIfCancellationRequested();
            source = InvoiceParser.Read(request.Xml);
            model = InvoiceModelValidator.ValidateSource(source);
            diagnostics.AddRange(source.UnmappedData);
            diagnostics.AddRange(model.Diagnostics);
            if (request.Target != null) diagnostics.AddRange(InvoiceSerializer.InspectTarget(source.Invoice, request.Target));
            token.ThrowIfCancellationRequested();
            bool writes = request.Operation is OfficeInvoiceWorkflowOperation.Convert or OfficeInvoiceWorkflowOperation.RenderPresentationPdf or OfficeInvoiceWorkflowOperation.RenderHybridPdf;
            if (writes && diagnostics.HasErrors) return Result(false);

            byte[]? xml = null;
            PdfInvoiceDocument? pdf = null;
            if (request.Operation == OfficeInvoiceWorkflowOperation.Convert)
                xml = source.Write(request.Target!);
            else if (writes) {
                pdf = PdfInvoiceDocument.Create(source.Invoice, request.Target!, request.Layout);
                xml = pdf.ToXmlBytes();
            }
            token.ThrowIfCancellationRequested();
            if (request.ValidationRelease.HasValue) {
#if NET8_0_OR_GREATER
                if (validator == null) {
                    diagnostics.Add(new InvoiceDiagnostic("INV-WORKFLOW-VALIDATOR", "Standards validation was requested, but no configured validator was supplied.", "Standards"));
                    return Result(false);
                }
                standards = await validator.ValidateAsync(xml ?? request.Xml, request.ValidationRelease.Value, token).ConfigureAwait(false);
                diagnostics.AddRange(standards.Diagnostics);
                if (!standards.IsValid) {
                    bool completed = standards.SchemaStatus == InvoiceValidationStatus.Invalid ||
                        standards.SchemaStatus == InvoiceValidationStatus.Passed && standards.BusinessRulesStatus == InvoiceValidationStatus.Invalid;
                    diagnostics.Add(new InvoiceDiagnostic("INV-WORKFLOW-STANDARDS", completed
                        ? "Standards validation completed and found invalid invoice data. No output artifact is returned."
                        : "The requested standards stages did not complete successfully. No output artifact is returned.", "Standards"));
                    if (request.Operation == OfficeInvoiceWorkflowOperation.Inspect && completed) return Result(true);
                    return Result(false);
                }

#else
                diagnostics.Add(new InvoiceDiagnostic("INV-WORKFLOW-VALIDATOR", "Standards validation requires .NET 8 or later. No output artifact is returned.", "Standards"));
                return Result(false);
#endif
            } else {
                diagnostics.Add(new InvoiceDiagnostic("INV-WORKFLOW-STANDARDS-NOT-RUN", "Only model and mapping checks ran. Standards schema and business rules were not requested.", "Standards", InvoiceDiagnosticSeverity.Information));
            }
            token.ThrowIfCancellationRequested();
            if (request.Operation == OfficeInvoiceWorkflowOperation.Inspect) return Result(true);
            if (request.Operation == OfficeInvoiceWorkflowOperation.Validate) return Result(!diagnostics.HasErrors);
            byte[] output = request.Operation switch {
                OfficeInvoiceWorkflowOperation.Convert => xml!,
                OfficeInvoiceWorkflowOperation.RenderPresentationPdf => pdf!.ToPresentationPdfBytes(request.PdfOptions, token),
                OfficeInvoiceWorkflowOperation.RenderHybridPdf => pdf!.ToPdfBytes(request.PdfOptions, token),
                _ => throw new InvalidOperationException("Unsupported invoice workflow operation.")
            };
            token.ThrowIfCancellationRequested();
            return Result(true, output, xml);
        } catch (Exception exception) when (exception is InvalidDataException or XmlException or ArgumentException or
            NotSupportedException or InvalidOperationException or OverflowException or IOException) {
            diagnostics.Add(new InvoiceDiagnostic("INV-WORKFLOW-OPERATION", exception.Message, source == null ? "Source" : "Operation"));
            return Result(false);
        }

        OfficeInvoiceWorkflowResult Result(bool succeeded, byte[]? output = null, byte[]? outputXml = null) =>
            new(request, hash, succeeded, source, model, diagnostics.ToList(),
#if NET8_0_OR_GREATER
                standards,
#endif
                 output, outputXml);
    }

    private sealed class OperationDiagnostics {
        private readonly List<InvoiceDiagnostic> _items = new();
        private readonly HashSet<(string Code, string Location, string Message, InvoiceDiagnosticSeverity Severity)> _keys = new();
        private int _omitted;
        private InvoiceDiagnosticSeverity _omittedSeverity;
        internal bool HasErrors { get; private set; }
        internal void AddRange(IEnumerable<InvoiceDiagnostic> diagnostics) { foreach (InvoiceDiagnostic diagnostic in diagnostics) Add(diagnostic); }
        internal void Add(InvoiceDiagnostic diagnostic) {
            if (diagnostic.Severity == InvoiceDiagnosticSeverity.Error) HasErrors = true;
            var key = (diagnostic.Code, diagnostic.Location, diagnostic.Message, diagnostic.Severity);
            if (_keys.Contains(key)) return;
            if (_items.Count < 999) { _keys.Add(key); _items.Add(diagnostic); }
            else { _omitted++; if (diagnostic.Severity > _omittedSeverity) _omittedSeverity = diagnostic.Severity; }
        }
        internal IReadOnlyList<InvoiceDiagnostic> ToList() {
            var result = new List<InvoiceDiagnostic>(_items);
            if (_omitted != 0) result.Add(new InvoiceDiagnostic("INV-WORKFLOW-DIAGNOSTICS-LIMIT", _omitted + " additional findings were omitted; the summary retains their highest severity.", "Diagnostics", _omittedSeverity));
            return result.AsReadOnly();
        }
    }
}
