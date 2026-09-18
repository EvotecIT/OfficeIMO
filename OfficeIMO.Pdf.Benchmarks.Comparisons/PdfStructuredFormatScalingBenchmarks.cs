using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Configs;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>
/// Measures complete parse, projection, layout, and PDF serialization across
/// adjacent and larger structured workloads. Each format is a separate logical
/// group because the source structures are intentionally format-native rather
/// than equivalent cross-format layouts. Validation reopens every artifact and
/// requires every record plus its format-specific nested/structured marker.
/// </summary>
[MemoryDiagnoser]
[GroupBenchmarksBy(BenchmarkLogicalGroupRule.ByCategory)]
[CategoriesColumn]
[RankColumn]
public class PdfStructuredFormatScalingBenchmarks {
    private PdfFormatConversionScenario _docx = null!;
    private PdfFormatConversionScenario _xlsx = null!;
    private PdfFormatConversionScenario _html = null!;
    private PdfFormatConversionScenario _markdown = null!;
    private byte[]? _docxResult;
    private byte[]? _xlsxResult;
    private byte[]? _htmlResult;
    private byte[]? _markdownResult;
    private PdfFormatConversionKind? _executedFormat;

    [Params(5, 6, 30, 120)]
    public int RecordCount { get; set; }

    [GlobalSetup]
    public void Setup() {
        _docx = CreateScenario(PdfFormatConversionKind.Docx);
        _xlsx = CreateScenario(PdfFormatConversionKind.Xlsx);
        _html = CreateScenario(PdfFormatConversionKind.Html);
        _markdown = CreateScenario(PdfFormatConversionKind.Markdown);
    }

    [Benchmark]
    [BenchmarkCategory("DOCX")]
    public byte[] ConvertDocxToPdfBytes() {
        _executedFormat = PdfFormatConversionKind.Docx;
        return _docxResult = _docx.ConvertToPdf();
    }

    [Benchmark]
    [BenchmarkCategory("XLSX")]
    public byte[] ConvertXlsxToPdfBytes() {
        _executedFormat = PdfFormatConversionKind.Xlsx;
        return _xlsxResult = _xlsx.ConvertToPdf();
    }

    [Benchmark]
    [BenchmarkCategory("HTML")]
    public byte[] ConvertHtmlToPdfBytes() {
        _executedFormat = PdfFormatConversionKind.Html;
        return _htmlResult = _html.ConvertToPdf();
    }

    [Benchmark]
    [BenchmarkCategory("Markdown")]
    public byte[] ConvertMarkdownToPdfBytes() {
        _executedFormat = PdfFormatConversionKind.Markdown;
        return _markdownResult = _markdown.ConvertToPdf();
    }

    [GlobalCleanup]
    public void Validate() {
        switch (_executedFormat) {
            case PdfFormatConversionKind.Docx:
                ValidateResult(_docx, _docxResult);
                break;
            case PdfFormatConversionKind.Xlsx:
                ValidateResult(_xlsx, _xlsxResult);
                break;
            case PdfFormatConversionKind.Html:
                ValidateResult(_html, _htmlResult);
                break;
            case PdfFormatConversionKind.Markdown:
                ValidateResult(_markdown, _markdownResult);
                break;
            default:
                throw new InvalidDataException(
                    $"Structured format scaling did not execute a supported converter for {RecordCount} records.");
        }
    }

    private PdfFormatConversionScenario CreateScenario(PdfFormatConversionKind format) =>
        PdfFormatConversionScenario.Create(
            format,
            recordCount: RecordCount,
            profile: PdfFormatConversionProfile.Structured);

    private void ValidateResult(PdfFormatConversionScenario scenario, byte[]? result) {
        if (result == null) {
            throw new InvalidDataException(
                $"{scenario.Kind}/{RecordCount} did not return a PDF result.");
        }

        PdfReadObservation observation = PdfBenchmarkValidation.ReadWithPdfPig(result);
        if (observation.PageCount < 1 || observation.PageCount > 100) {
            throw new InvalidDataException(
                $"{scenario.Kind}/{RecordCount} produced an implausible {observation.PageCount} pages.");
        }

        foreach (string marker in scenario.RequiredText) {
            string normalized = PdfBenchmarkValidation.Normalize(marker);
            if (!observation.NormalizedText.Contains(normalized, StringComparison.Ordinal)) {
                throw new InvalidDataException(
                    $"{scenario.Kind}/{RecordCount} did not preserve required text '{marker}'.");
            }
        }

        Console.WriteLine(
            $"STRUCTURED_FORMAT_PDF_EVIDENCE format={scenario.Kind} records={RecordCount} " +
            $"sourceBytes={scenario.SourceBytes.Length} pdfBytes={result.Length} " +
            $"pages={observation.PageCount} textLength={observation.TextLength}");
    }
}
