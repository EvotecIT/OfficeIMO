using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>
/// Measures complete source parsing, projection, PDF layout, and serialization for
/// the remaining advertised conversion routes. Cleanup reopens each artifact and
/// checks the complete semantic record manifest.
/// </summary>
[MemoryDiagnoser]
[RankColumn]
public class PdfExtendedFormatConversionBenchmarks {
    private PdfExtendedFormatConversionScenario _scenario = null!;
    private byte[]? _result;

    [Params(
        PdfExtendedFormatConversionKind.AsciiDoc,
        PdfExtendedFormatConversionKind.Latex,
        PdfExtendedFormatConversionKind.Mhtml,
        PdfExtendedFormatConversionKind.OneNote,
        PdfExtendedFormatConversionKind.Odt,
        PdfExtendedFormatConversionKind.Ods,
        PdfExtendedFormatConversionKind.Odp,
        PdfExtendedFormatConversionKind.Visio)]
    public PdfExtendedFormatConversionKind Format { get; set; }

    [GlobalSetup]
    public void Setup() => _scenario = PdfExtendedFormatConversionScenario.Create(Format);

    [Benchmark]
    public byte[] ConvertToPdfBytes() => _result = _scenario.ConvertToPdfBytes();

    [GlobalCleanup]
    public void Validate() {
        if (_result == null) {
            throw new InvalidDataException($"{Format} did not return a PDF result.");
        }

        PdfReadObservation observation = PdfBenchmarkValidation.ReadWithPdfPig(_result);
        if (observation.PageCount < 1 || observation.PageCount > 100) {
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
            $"EXTENDED_FORMAT_PDF_EVIDENCE format={Format} sourceBytes={_scenario.SourceBytes.Length} " +
            $"pdfBytes={_result.Length} pages={observation.PageCount} textLength={observation.TextLength}");
    }
}
