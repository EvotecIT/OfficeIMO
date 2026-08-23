using BenchmarkDotNet.Attributes;
using HtmlTinkerX;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>
/// Compares complete in-memory HTML parsing, paged layout, and PDF serialization
/// from one identical HTML document.
/// </summary>
[MemoryDiagnoser]
[RankColumn]
public class PdfHtmlBenchmarks {
    private PdfBenchmarkScenario _scenario = null!;
    private string _html = null!;
    private HtmlBrowserSession? _browserSession;
    private byte[]? _officeImoResult;
    private byte[]? _peachPdfResult;
    private byte[]? _iTextPdfHtmlResult;
    private byte[]? _chromiumResult;

    [Params(PdfBenchmarkScale.Easy, PdfBenchmarkScale.Medium, PdfBenchmarkScale.High)]
    public PdfBenchmarkScale Scale { get; set; }

    [GlobalSetup(Targets = new[] { nameof(OfficeIMO), nameof(PeachPDF), nameof(ITextPdfHtml) })]
    public void SetupManaged() {
        InitializeScenario();
    }

    [GlobalSetup(Target = nameof(Chromium))]
    public async Task SetupChromium() {
        InitializeScenario();
        _browserSession = await HtmlPdfComparisonRenderers.OpenChromiumSessionAsync().ConfigureAwait(false);
    }

    private void InitializeScenario() {
        _scenario = PdfBenchmarkScenario.Get(Scale);
        _html = PdfHtmlScenarioBuilder.Create(_scenario);
    }

    [Benchmark(Baseline = true)]
    public byte[] OfficeIMO() =>
        _officeImoResult = HtmlPdfComparisonRenderers.RenderManaged(HtmlPdfComparisonEngine.OfficeIMO, _html);

    [Benchmark]
    public byte[] PeachPDF() =>
        _peachPdfResult = HtmlPdfComparisonRenderers.RenderManaged(HtmlPdfComparisonEngine.PeachPDF, _html);

    [Benchmark]
    public byte[] ITextPdfHtml() =>
        _iTextPdfHtmlResult = HtmlPdfComparisonRenderers.RenderManaged(HtmlPdfComparisonEngine.ITextPdfHtml, _html);

    [Benchmark]
    public async Task<byte[]> Chromium() {
        HtmlBrowserSession session = _browserSession
            ?? throw new InvalidOperationException("The HtmlTinkerX Chromium session was not initialized.");
        _chromiumResult = await HtmlPdfComparisonRenderers.RenderChromiumAsync(session, _html).ConfigureAwait(false);
        return _chromiumResult;
    }

    [GlobalCleanup(Target = nameof(OfficeIMO))]
    public void ValidateOfficeIMO() => Validate(nameof(OfficeIMO), _officeImoResult);

    [GlobalCleanup(Target = nameof(PeachPDF))]
    public void ValidatePeachPDF() => Validate(nameof(PeachPDF), _peachPdfResult);

    [GlobalCleanup(Target = nameof(ITextPdfHtml))]
    public void ValidateITextPdfHtml() => Validate(nameof(ITextPdfHtml), _iTextPdfHtmlResult);

    [GlobalCleanup(Target = nameof(Chromium))]
    public async Task ValidateChromium() {
        try {
            Validate(nameof(Chromium), _chromiumResult);
        } finally {
            if (_browserSession != null) {
                await _browserSession.DisposeAsync().ConfigureAwait(false);
                _browserSession = null;
            }
        }
    }

    private void Validate(string engine, byte[]? bytes) {
        if (bytes == null) {
            throw new InvalidDataException($"{engine} did not return a PDF for {_scenario.Scale}.");
        }

        PdfReadObservation observation = PdfBenchmarkValidation.ValidateGenerated(bytes, _scenario, engine);
        PdfBenchmarkValidation.ValidateTaggedStructure(bytes, engine, _scenario.PageCount);
        Console.WriteLine(
            $"HTML_PDF_EVIDENCE engine={engine} scale={_scenario.Scale} htmlBytes={System.Text.Encoding.UTF8.GetByteCount(_html)} " +
            $"pdfBytes={bytes.Length} pages={observation.PageCount} textLength={observation.TextLength}");
    }
}
