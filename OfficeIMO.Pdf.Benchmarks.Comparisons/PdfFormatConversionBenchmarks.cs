using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>
/// Measures complete source parsing, PDF layout, and serialization for the primary
/// fixed-layout conversion routes. Each result is reopened and checked for the
/// complete semantic record manifest before the lane is accepted.
/// </summary>
[MemoryDiagnoser]
[RankColumn]
public class PdfFormatConversionBenchmarks {
    private PdfFormatConversionScenario _scenario = null!;
    private byte[]? _result;

    [Params(
        PdfFormatConversionKind.Docx,
        PdfFormatConversionKind.Xlsx,
        PdfFormatConversionKind.Pptx,
        PdfFormatConversionKind.Html,
        PdfFormatConversionKind.Markdown,
        PdfFormatConversionKind.Rtf)]
    public PdfFormatConversionKind Format { get; set; }

    [GlobalSetup]
    public void Setup() => _scenario = PdfFormatConversionScenario.Create(Format);

    [Benchmark]
    [BenchmarkCategory("library:OfficeIMO", "case:ConvertToPdfBytes")]
    public byte[] ConvertToPdfBytes() => _result = _scenario.ConvertToPdf();

    // Cleanup validates the PDF the last measured invocation produced, outside timing, and records the outcome.
    [GlobalCleanup]
    public void Validate() =>
        PdfBenchmarkValidation.Record(
            "OfficeIMO",
            nameof(ConvertToPdfBytes),
            $"Format={Format}; measured PDF reopened with PdfPig: 1-80 pages and every record of the source manifest",
            ValidateResult);

    private void ValidateResult() {
        if (_result == null) {
            throw new InvalidDataException($"{Format} did not return a PDF result.");
        }

        PdfReadObservation observation = PdfBenchmarkValidation.ReadWithPdfPig(_result);
        if (observation.PageCount < 1 || observation.PageCount > 80) {
            throw new InvalidDataException(
                $"{Format} produced an implausible {observation.PageCount} pages for the conversion workload.");
        }

        foreach (string marker in _scenario.RequiredText) {
            string normalized = PdfBenchmarkValidation.Normalize(marker);
            if (!observation.NormalizedText.Contains(normalized, StringComparison.Ordinal)) {
                throw new InvalidDataException($"{Format} did not preserve required text '{marker}'.");
            }
        }

        Console.WriteLine(
            $"FORMAT_PDF_EVIDENCE format={Format} sourceBytes={_scenario.SourceBytes.Length} " +
            $"pdfBytes={_result.Length} pages={observation.PageCount} textLength={observation.TextLength}");
    }
}
