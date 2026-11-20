using BenchmarkDotNet.Attributes;
using HtmlTinkerX;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>
/// Compares complete in-memory HTML parsing, paged layout, and PDF serialization
/// from one identical HTML document. Chromium opens a fresh browser session before every iteration, which makes
/// BenchmarkDotNet time one invocation per iteration; every library runs with that same job so the results compare
/// as one configuration.
/// </summary>
[MemoryDiagnoser]
[RankColumn]
[InvocationCount(1)]
public class PdfHtmlBenchmarks {
    private PdfBenchmarkScenario _scenario = null!;
    private string _html = null!;
    private HtmlBrowserSession? _browserSession;
    private byte[]? _officeImoResult;
    private byte[]? _peachPdfResult;
    private byte[]? _pdfPeckerResult;
    private byte[]? _iTextPdfHtmlResult;
    private byte[]? _chromiumResult;

    [Params(PdfBenchmarkScale.Easy, PdfBenchmarkScale.Medium, PdfBenchmarkScale.High)]
    public PdfBenchmarkScale Scale { get; set; }

    [GlobalSetup(Targets = new[] { nameof(OfficeIMO), nameof(PeachPDF), nameof(PdfPecker), nameof(ITextPdfHtml) })]
    public void SetupManaged() {
        InitializeScenario();
    }

    [GlobalSetup(Target = nameof(Chromium))]
    public void SetupChromium() {
        InitializeScenario();
    }

    [IterationSetup(Target = nameof(Chromium))]
    public void SetupChromiumIteration() {
        // Keep browser startup outside the measured operation while isolating
        // each iteration from stale DevTools PDF stream state.
        DisposeChromiumSession();
        _browserSession = HtmlPdfComparisonRenderers.OpenChromiumSessionAsync().GetAwaiter().GetResult();
    }

    [IterationCleanup(Target = nameof(Chromium))]
    public void CleanupChromiumIteration() {
        DisposeChromiumSession();
    }

    private void InitializeScenario() {
        _scenario = PdfBenchmarkScenario.Get(Scale);
        _html = PdfHtmlScenarioBuilder.Create(_scenario);
    }

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("library:OfficeIMO", "case:Default")]
    public byte[] OfficeIMO() =>
        _officeImoResult = HtmlPdfComparisonRenderers.RenderManaged(HtmlPdfComparisonEngine.OfficeIMO, _html);

    [Benchmark]
    [BenchmarkCategory("library:PeachPDF", "case:Default")]
    public byte[] PeachPDF() =>
        _peachPdfResult = HtmlPdfComparisonRenderers.RenderManaged(HtmlPdfComparisonEngine.PeachPDF, _html);

    [Benchmark]
    [BenchmarkCategory("library:PdfPecker", "case:Default")]
    public byte[] PdfPecker() =>
        _pdfPeckerResult = HtmlPdfComparisonRenderers.RenderManaged(HtmlPdfComparisonEngine.PdfPecker, _html);

    [Benchmark]
    [BenchmarkCategory("library:iText pdfHTML", "case:Default")]
    public byte[] ITextPdfHtml() =>
        _iTextPdfHtmlResult = HtmlPdfComparisonRenderers.RenderManaged(HtmlPdfComparisonEngine.ITextPdfHtml, _html);

    [Benchmark]
    [BenchmarkCategory("library:Chromium", "case:Default")]
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

    [GlobalCleanup(Target = nameof(PdfPecker))]
    public void ValidatePdfPecker() => Validate(nameof(PdfPecker), _pdfPeckerResult);

    [GlobalCleanup(Target = nameof(ITextPdfHtml))]
    public void ValidateITextPdfHtml() => Validate(nameof(ITextPdfHtml), _iTextPdfHtmlResult);

    [GlobalCleanup(Target = nameof(Chromium))]
    public void ValidateChromium() {
        try {
            Validate(nameof(Chromium), _chromiumResult);
        } finally {
            DisposeChromiumSession();
        }
    }

    private void DisposeChromiumSession() {
        if (_browserSession == null) {
            return;
        }

        _browserSession.DisposeAsync().AsTask().GetAwaiter().GetResult();
        _browserSession = null;
    }

    // Cleanup validates the PDF the last measured invocation produced, outside timing, and records the outcome.
    private void Validate(string engine, byte[]? bytes) {
        PdfReadObservation observation = default;
        PdfBenchmarkValidation.Record(
            PdfBenchmarkValidation.LibraryId(engine),
            "Default",
            $"Scale={Scale}; measured PDF reopened with PdfPig: A4 pages with the exact visible report text in reading order, tag tree, embedded Carlito",
            () => {
                if (bytes == null) {
                    throw new InvalidDataException($"{engine} did not return a PDF for {_scenario.Scale}.");
                }

                observation = PdfBenchmarkValidation.ValidateHtmlGenerated(bytes, _scenario, engine);
            });
        Console.WriteLine(
            $"HTML_PDF_EVIDENCE engine={engine} scale={_scenario.Scale} htmlBytes={System.Text.Encoding.UTF8.GetByteCount(_html)} " +
            $"pdfBytes={bytes.Length} pages={observation.PageCount} textLength={observation.TextLength}");
    }
}
