using BenchmarkDotNet.Attributes;
using OfficeIMO.Invoicing.Pdf;
using OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>Measures OfficeIMO's complete typed invoice-to-electronic-PDF workflow without competitor ranking.</summary>
[MemoryDiagnoser]
public class PdfTypedInvoiceWorkflowBenchmarks {
    private InvoiceComparisonScenario _scenario = null!;

    [GlobalSetup]
    public void Setup() {
        BenchmarkAffinityGuard.Validate();
        _scenario = InvoiceComparisonScenario.Create();
        PdfInvoiceDocument snapshot = _scenario.CreateTypedSnapshot();
        byte[] pdf = snapshot.ToPdfBytes(_scenario.PdfOptions);
        PdfInvoiceComparisonValidation.Validate(pdf, _scenario, nameof(TypedElectronicInvoicePdf));
        PdfExtractedAttachment attachment = AssertSingleAttachment(pdf);
        if (!attachment.Bytes.SequenceEqual(snapshot.ToXmlBytes()))
            throw new InvalidDataException("The typed workflow PDF did not preserve the generated CII attachment bytes.");
    }

    [Benchmark]
    public byte[] TypedElectronicInvoicePdf() => Generate();

    private byte[] Generate() => _scenario.CreateTypedSnapshot().ToPdfBytes(_scenario.PdfOptions);

    private static PdfExtractedAttachment AssertSingleAttachment(byte[] pdf) {
        IReadOnlyList<PdfExtractedAttachment> attachments = PdfDocument.Load(pdf).Attachments.Extract();
        if (attachments.Count != 1) throw new InvalidDataException($"Expected one CII attachment, found {attachments.Count}.");
        return attachments[0];
    }
}
