using System.Security.Cryptography;
using System.Xml;
using OfficeIMO.Invoicing;
using OfficeIMO.Invoicing.Validation;

namespace OfficeIMO.Workflows;

public static partial class OfficeInvoiceBufferWorkflow {
    private static async Task<OfficeInvoiceWorkflowResult> ProcessSourceEditAsync(OfficeInvoiceWorkflowRequest request,
        InvoiceValidator? validator, CancellationToken token) {
        string hash = Convert.ToHexString(SHA256.HashData(request.Xml));
        var diagnostics = new OperationDiagnostics();
        InvoiceReadResult? source = null;
        InvoiceModelValidationResult? model = null;
        InvoiceValidationReport? standards = null;
        try {
            token.ThrowIfCancellationRequested();
            var edit = InvoiceSourceEditor.Apply(InvoiceSourceDocument.Load(request.Xml), request.SourceEdits!);
            diagnostics.AddRange(edit.Diagnostics);
            if (!edit.Succeeded) return Result(false);
            byte[] xml = edit.Document!.ToBytes();
            token.ThrowIfCancellationRequested();
            // Mapping is evidence only: a safe header edit must not discard extensions or require semantic rewriting.
            try {
                source = InvoiceParser.Read(xml);
                model = InvoiceModelValidator.ValidateSource(source);
                diagnostics.AddRange(source.UnmappedData);
                diagnostics.AddRange(model.Diagnostics);
            } catch (Exception exception) when (exception is InvalidDataException or XmlException or ArgumentException or
                NotSupportedException or InvalidOperationException or OverflowException) {
                diagnostics.Add(new InvoiceDiagnostic("INV-WORKFLOW-MODEL-UNAVAILABLE", exception.Message, "Model"));
            }
            token.ThrowIfCancellationRequested();
            if (request.ValidationRelease.HasValue) {
                if (validator == null) {
                    diagnostics.Add(new InvoiceDiagnostic("INV-WORKFLOW-VALIDATOR", "Standards validation was requested, but no configured validator was supplied.", "Standards"));
                    return Result(false);
                }
                standards = await validator.ValidateAsync(xml, request.ValidationRelease.Value, token).ConfigureAwait(false);
                diagnostics.AddRange(standards.Diagnostics);
                if (!standards.IsValid) {
                    diagnostics.Add(new InvoiceDiagnostic("INV-WORKFLOW-STANDARDS", "The edited XML did not pass all requested standards stages. No output artifact is returned.", "Standards"));
                    return Result(false);
                }
            } else {
                diagnostics.Add(new InvoiceDiagnostic("INV-WORKFLOW-STANDARDS-NOT-RUN", "The source edit retains XML content; standards schema and business rules were not requested.", "Standards", InvoiceDiagnosticSeverity.Information));
            }
            token.ThrowIfCancellationRequested();
            return Result(true, xml);
        } catch (Exception exception) when (exception is InvalidDataException or XmlException or ArgumentException or
            NotSupportedException or InvalidOperationException or OverflowException or IOException) {
            diagnostics.Add(new InvoiceDiagnostic("INV-WORKFLOW-OPERATION", exception.Message, "Source"));
            return Result(false);
        }

        OfficeInvoiceWorkflowResult Result(bool succeeded, byte[]? xml = null) =>
            new(request, hash, succeeded, source, model, diagnostics.ToList(), standards, xml, xml);
    }
}
