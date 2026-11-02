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
        InvoiceBenchmarkCorrectness.Write(WriteXml(), _corpus);
    }

    [Benchmark]
    [BenchmarkCategory("xml-write")]
    public byte[] WriteXml() => InvoiceSerializer.Write(_corpus.Invoice, InvoiceBenchmarkCorpus.Contract);
}

[MemoryDiagnoser]
public class InvoiceXmlReadBenchmarks {
    private InvoiceBenchmarkCorpus _corpus = null!;

    [GlobalSetup]
    public void Setup() {
        _corpus = InvoiceBenchmarkCorpus.Create();
        InvoiceBenchmarkCorrectness.Read(ReadXml(), _corpus);
    }

    [Benchmark]
    [BenchmarkCategory("xml-read")]
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
        InvoiceBenchmarkCorrectness.Rules(await ValidateRules().ConfigureAwait(false));
    }

    [Benchmark]
    [BenchmarkCategory("rules-validation")]
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
        InvoiceBenchmarkCorrectness.Pdf(GeneratePdf(), _corpus.Xml);
    }

    [Benchmark]
    [BenchmarkCategory("pdf-generation")]
    public byte[] GeneratePdf() => _snapshot.ToPdfBytes(_corpus.PdfOptions);
}
