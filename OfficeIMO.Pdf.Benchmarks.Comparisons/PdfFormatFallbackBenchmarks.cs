using BenchmarkDotNet.Attributes;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>
/// Diagnoses the cost of automatic installed-font discovery and embedding for
/// representative source formats. Output is reopened after every benchmark case
/// so a cheaper configuration is never accepted by timing alone.
/// </summary>
[MemoryDiagnoser]
[RankColumn]
public class PdfFormatFallbackBenchmarks {
    private PdfFormatConversionScenario _scenario = null!;
    private byte[]? _result;

    [Params(
        PdfFormatConversionKind.Docx,
        PdfFormatConversionKind.Xlsx,
        PdfFormatConversionKind.Pptx,
        PdfFormatConversionKind.Markdown)]
    public PdfFormatConversionKind Format { get; set; }

    [Params(false, true)]
    public bool AutomaticFallbacks { get; set; }

    [GlobalSetup]
    public void Setup() => _scenario = PdfFormatConversionScenario.Create(
        Format,
        AutomaticFallbacks
            ? PdfCore.PdfTextFallbackFeatures.Default
            : PdfCore.PdfTextFallbackFeatures.None);

    [Benchmark]
    public byte[] ConvertToPdf() => _result = _scenario.ConvertToPdf();

    [GlobalCleanup]
    public void Validate() {
        if (_result == null) {
            throw new InvalidDataException($"{Format} did not return a PDF result.");
        }

        PdfReadObservation observation = PdfBenchmarkValidation.ReadWithPdfPig(_result);
        foreach (string marker in _scenario.RequiredText) {
            if (!observation.NormalizedText.Contains(
                    PdfBenchmarkValidation.Normalize(marker),
                    StringComparison.Ordinal)) {
                throw new InvalidDataException(
                    $"{Format} with AutomaticFallbacks={AutomaticFallbacks} lost required text '{marker}'.");
            }
        }

        Console.WriteLine(
            $"FORMAT_FALLBACK_EVIDENCE format={Format} automaticFallbacks={AutomaticFallbacks} " +
            $"sourceBytes={_scenario.SourceBytes.Length} pdfBytes={_result.Length} pages={observation.PageCount}");
    }
}
