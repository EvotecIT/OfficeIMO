using OfficeIMO.Invoicing.Validation;
using OfficeIMO.Pdf;

namespace OfficeIMO.Invoicing.Benchmarks;

internal static class InvoiceBenchmarkCorrectness {
    internal static void Write(byte[] xml) => Read(InvoiceParser.Read(xml));

    internal static void Read(InvoiceReadResult result) {
        if (!result.HasCompleteMapping || result.Invoice.Lines.Count != InvoiceBenchmarkCorpus.LineCount ||
            !string.Equals(result.Invoice.Lines[^1].Description?.Split(' ').Last(), InvoiceBenchmarkCorpus.Marker, StringComparison.Ordinal))
            throw new InvalidDataException("The XML workload did not preserve the equivalent invoice corpus.");
    }

    internal static void Rules(InvoiceValidationReport report) {
        if (report.SchemaStatus != InvoiceValidationStatus.Passed || report.BusinessRulesStatus != InvoiceValidationStatus.Passed || !report.IsValid)
            throw new InvalidDataException("The rules workload did not pass the pinned Factur-X schema and business rules: " + string.Join("; ", report.Diagnostics.Select(item => item.Code + " " + item.Message)));
    }

    internal static void Pdf(byte[] pdf, byte[] xml) {
        PdfDocument document = PdfDocument.Load(pdf);
        if (!document.Attachments.Extract().Single().Bytes.SequenceEqual(xml))
            throw new InvalidDataException("The PDF workload did not embed the exact equivalent XML corpus.");
        string text = PdfReadDocument.Open(pdf).ExtractText();
        if (!text.Contains(InvoiceBenchmarkCorpus.Marker, StringComparison.Ordinal) || !text.Contains("日本語", StringComparison.Ordinal) || !text.Contains("العربية", StringComparison.Ordinal))
            throw new InvalidDataException("The PDF workload did not preserve representative visible multilingual text.");
    }
}
