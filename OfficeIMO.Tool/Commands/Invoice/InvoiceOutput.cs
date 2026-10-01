using System.Text;
using System.Text.Json;
using OfficeIMO.Invoicing;
using OfficeIMO.Workflows;

namespace OfficeIMO.Tool.Commands.Invoice;

internal static class InvoiceOutput {
    internal static async Task WriteAsync(TextWriter output, IReadOnlyList<OfficeInvoiceFileWorkflowResult> results, bool batch) {
        using var bytes = new MemoryStream();
        using (var json = new Utf8JsonWriter(bytes, new JsonWriterOptions { Indented = true })) {
            if (batch) { json.WriteStartObject(); json.WriteString("schema", "officeimo.invoice.batch.v1"); json.WriteStartArray("results"); }
            foreach (var result in results) WriteResult(json, result);
            if (batch) { json.WriteEndArray(); json.WriteEndObject(); }
        }
        await output.WriteLineAsync(Encoding.UTF8.GetString(bytes.ToArray())).ConfigureAwait(false);
    }
    private static void WriteResult(Utf8JsonWriter json, OfficeInvoiceFileWorkflowResult file) {
        var result = file.Workflow;
        json.WriteStartObject(); json.WriteString("schema", "officeimo.invoice.result.v1");
        json.WriteString("operation", result.Operation.ToString()); json.WriteBoolean("succeeded", file.Succeeded);
        json.WriteString("inputPath", file.InputPath); json.WriteString("inputSha256", result.InputSha256);
        json.WriteNumber("inputByteLength", result.InputByteLength); json.WriteString("outputPath", file.OutputPath);
        json.WriteBoolean("published", file.Published); json.WriteString("publicationError", file.PublicationError);
        json.WriteNumber("outputByteLength", result.OutputByteLength);
        if (result.Source is { } source) {
            json.WriteStartObject("source"); json.WriteString("syntax", source.Declaration.Syntax.ToString());
            json.WriteString("profile", source.Declaration.Profile?.ToString()); json.WriteBoolean("completeMapping", source.HasCompleteMapping);
            json.WriteString("documentId", source.Invoice.Number); json.WriteString("currency", source.Invoice.Currency);
            json.WriteNumber("lineCount", source.Invoice.Lines.Count); json.WriteEndObject();
        } else json.WriteNull("source");
        if (result.ModelValidation is { } model) {
            json.WriteStartObject("model"); json.WriteBoolean("isValid", model.IsValid);
            if (model.Calculation is { } calculation) {
                json.WriteNumber("taxExclusiveAmount", calculation.TaxExclusiveTotal); json.WriteNumber("taxAmount", calculation.TaxTotal);
                json.WriteNumber("taxInclusiveAmount", calculation.TaxInclusiveTotal); json.WriteNumber("payableAmount", calculation.PayableAmount);
            }
            json.WriteEndObject();
        } else json.WriteNull("model");
        json.WriteStartObject("standards"); json.WriteString("schemaStatus", result.SchemaStatus.ToString());
        json.WriteString("businessRulesStatus", result.BusinessRulesStatus.ToString());
        json.WriteString("validatedSha256", result.StandardsValidation?.Sha256);
        json.WriteString("release", result.StandardsValidation?.Release.ToString());
        json.WriteEndObject(); json.WriteStartArray("diagnostics");
        foreach (InvoiceDiagnostic diagnostic in result.Diagnostics) {
            json.WriteStartObject(); json.WriteString("code", diagnostic.Code); json.WriteString("severity", diagnostic.Severity.ToString());
            json.WriteString("location", diagnostic.Location); json.WriteString("message", diagnostic.Message); json.WriteEndObject();
        }
        json.WriteEndArray(); json.WriteEndObject();
    }
}
