using BenchmarkDotNet.Attributes;
using OfficeIMO.Invoicing.Pdf;
using OfficeIMO.Invoicing.Validation;

namespace OfficeIMO.Invoicing.Benchmarks;

[MemoryDiagnoser]
public class InvoiceXmlWriteBenchmarks {
    private InvoiceBenchmarkCorpus _corpus = null!;

    [GlobalSetup]
    public void Setup() {
        _corpus = InvoiceBenchmarkCorpus.Create();
        byte[] xml = WriteXml();
        InvoiceBenchmarkCorrectness.Record(nameof(WriteXml), "written XML is byte-identical to the canonical corpus and parses back completely", () => InvoiceBenchmarkCorrectness.Write(xml, _corpus));
    }

    [Benchmark]
    [BenchmarkCategory("xml-write")]
    [BenchmarkCategory("library:OfficeIMO", "case:WriteXml")]
    public byte[] WriteXml() => InvoiceSerializer.Write(_corpus.Invoice, InvoiceBenchmarkCorpus.Contract);
}

[MemoryDiagnoser]
public class InvoiceXmlReadBenchmarks {
    private InvoiceBenchmarkCorpus _corpus = null!;

    [GlobalSetup]
    public void Setup() {
        _corpus = InvoiceBenchmarkCorpus.Create();
        InvoiceReadResult result = ReadXml();
        InvoiceBenchmarkCorrectness.Record(nameof(ReadXml), "complete mapping; the parsed invoice writes back byte-identical to the corpus", () => InvoiceBenchmarkCorrectness.Read(result, _corpus));
    }

    [Benchmark]
    [BenchmarkCategory("xml-read")]
    [BenchmarkCategory("library:OfficeIMO", "case:ReadXml")]
    public InvoiceReadResult ReadXml() => InvoiceParser.Read(_corpus.Xml);
}

[MemoryDiagnoser]
public class InvoiceRulesValidationBenchmarks {
    private InvoiceBenchmarkCorpus _corpus = null!;
    private InvoiceValidator _validator = null!;

    [GlobalSetup]
    public async Task Setup() {
        _corpus = InvoiceBenchmarkCorpus.Create();
        _validator = InvoiceBenchmarkAuthority.CreateValidator();
        InvoiceValidationReport report = await ValidateRules().ConfigureAwait(false);
        InvoiceBenchmarkCorrectness.Record(nameof(ValidateRules), "pinned Factur-X schema and business rules pass", () => InvoiceBenchmarkCorrectness.Rules(report));
    }

    [Benchmark]
    [BenchmarkCategory("rules-validation")]
    [BenchmarkCategory("library:OfficeIMO", "case:ValidateRules")]
    public Task<InvoiceValidationReport> ValidateRules() => _validator.ValidateAsync(
        _corpus.Xml,
        InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2);
}

[MemoryDiagnoser]
public class InvoicePdfGenerationBenchmarks {
    private InvoiceBenchmarkCorpus _corpus = null!;
    private PdfInvoiceDocument _snapshot = null!;

    [GlobalSetup]
    public void Setup() {
        _corpus = InvoiceBenchmarkCorpus.Create();
        _snapshot = _corpus.CapturePdf();
        byte[] pdf = GeneratePdf();
        InvoiceBenchmarkCorrectness.Record(nameof(GeneratePdf), "PDF reopened: the single attachment is the exact corpus XML and the visible text keeps the marker, Japanese and Arabic samples", () => InvoiceBenchmarkCorrectness.Pdf(pdf, _corpus.Xml));
    }

    [Benchmark]
    [BenchmarkCategory("pdf-generation")]
    [BenchmarkCategory("library:OfficeIMO", "case:GeneratePdf")]
    public byte[] GeneratePdf() => _snapshot.ToPdfBytes(_corpus.PdfOptions);
}
