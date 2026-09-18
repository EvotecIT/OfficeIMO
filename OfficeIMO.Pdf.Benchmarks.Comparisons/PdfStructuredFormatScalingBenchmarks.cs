using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>
/// Measures complete parse, projection, layout, and PDF serialization across
/// adjacent and larger structured workloads. Validation reopens the artifact and
/// requires every record plus a format-specific nested/structured marker.
/// </summary>
[MemoryDiagnoser]
[RankColumn]
public class PdfStructuredFormatScalingBenchmarks {
    private PdfFormatConversionScenario _scenario = null!;
    private byte[]? _result;

    [Params(
        PdfFormatConversionKind.Docx,
        PdfFormatConversionKind.Xlsx,
        PdfFormatConversionKind.Html,
        PdfFormatConversionKind.Markdown)]
    public PdfFormatConversionKind Format { get; set; }

    [Params(5, 6, 30, 120)]
    public int RecordCount { get; set; }

    [GlobalSetup]
    public void Setup() => _scenario = PdfFormatConversionScenario.Create(
        Format,
        recordCount: RecordCount,
        profile: PdfFormatConversionProfile.Structured);

    [Benchmark]
    public byte[] ConvertToPdfBytes() => _result = _scenario.ConvertToPdf();

    [GlobalCleanup]
    public void Validate() {
        if (_result == null) {
            throw new InvalidDataException($"{Format}/{RecordCount} did not return a PDF result.");
        }

        PdfReadObservation observation = PdfBenchmarkValidation.ReadWithPdfPig(_result);
        if (observation.PageCount < 1 || observation.PageCount > 100) {
            throw new InvalidDataException(
                $"{Format}/{RecordCount} produced an implausible {observation.PageCount} pages.");
        }

        foreach (string marker in _scenario.RequiredText) {
            string normalized = PdfBenchmarkValidation.Normalize(marker);
            if (!observation.NormalizedText.Contains(normalized, StringComparison.Ordinal)) {
                throw new InvalidDataException(
                    $"{Format}/{RecordCount} did not preserve required text '{marker}'.");
            }
        }

        Console.WriteLine(
            $"STRUCTURED_FORMAT_PDF_EVIDENCE format={Format} records={RecordCount} " +
            $"sourceBytes={_scenario.SourceBytes.Length} pdfBytes={_result.Length} " +
            $"pages={observation.PageCount} textLength={observation.TextLength}");
    }
}
