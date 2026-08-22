using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>
/// Compares cold-start and warm throughput for exact 21 KiB HTML payloads.
/// Use BenchmarkDotNet's Dry job for process-isolated cold-start evidence.
/// </summary>
[MemoryDiagnoser]
[RankColumn]
public class PdfHtmlPayloadBenchmarks {
    private PdfHtmlPayloadScenario _scenario = null!;
    private byte[]? _officeImoResult;
    private byte[]? _peachPdfResult;

    [Params(PdfHtmlPayloadKind.PlainText, PdfHtmlPayloadKind.Table, PdfHtmlPayloadKind.Multilingual)]
    public PdfHtmlPayloadKind Payload { get; set; }

    [GlobalSetup]
    public void Setup() => _scenario = PdfHtmlPayloadScenario.Create(Payload);

    [Benchmark(Baseline = true)]
    public byte[] OfficeIMO() =>
        _officeImoResult = OfficeImoPdfGenerator.GenerateHtml(_scenario.Html);

    [Benchmark]
    public byte[] PeachPDF() =>
        _peachPdfResult = PeachPdfGenerator.Generate(_scenario.Html);

    [GlobalCleanup(Target = nameof(OfficeIMO))]
    public void ValidateOfficeIMO() => Validate(nameof(OfficeIMO), _officeImoResult);

    [GlobalCleanup(Target = nameof(PeachPDF))]
    public void ValidatePeachPDF() => Validate(nameof(PeachPDF), _peachPdfResult);

    private void Validate(string engine, byte[]? bytes) {
        if (bytes == null) {
            throw new InvalidDataException($"{engine} did not return a result for {_scenario.Kind}.");
        }

        PdfReadObservation observation = PdfBenchmarkValidation.ReadWithPdfPig(bytes);
        if (observation.PageCount < 2 || observation.PageCount > 20) {
            throw new InvalidDataException(
                $"{engine} produced an implausible {observation.PageCount} pages for the paginated 21 KiB {_scenario.Kind} workload.");
        }

        string normalized = observation.NormalizedText;
        foreach (string required in _scenario.RequiredText) {
            string fragment = PdfBenchmarkValidation.Normalize(required);
            if (!normalized.Contains(fragment, StringComparison.Ordinal)) {
                throw new InvalidDataException($"{engine} did not preserve required text '{required}' for {_scenario.Kind}.");
            }
        }

        Console.WriteLine(
            $"HTML_PDF_EVIDENCE engine={engine} payload={_scenario.Kind} htmlBytes={PdfHtmlPayloadScenario.TargetUtf8Bytes} " +
            $"pdfBytes={bytes.Length} pages={observation.PageCount} textLength={observation.TextLength}");
    }
}
