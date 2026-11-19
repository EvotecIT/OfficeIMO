using BenchmarkDotNet.Attributes;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
#if NET10_0_OR_GREATER
using OneRender.Text;
#endif
using System.Globalization;
using System.Text;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>
/// Compares cold-start and warm throughput for exact 21 KiB HTML payloads.
/// Use BenchmarkDotNet's Dry job for process-isolated cold-start evidence.
/// </summary>
[MemoryDiagnoser]
[RankColumn]
public class PdfHtmlPayloadBenchmarks {
    private PdfHtmlPayloadScenario _scenario = null!;
    private PdfEmbeddedFontFamily? _sharedFont;
    private byte[]? _officeImoResult;
    private byte[]? _peachPdfResult;
    private byte[]? _pdfPeckerResult;
#if NET10_0_OR_GREATER
    private byte[]? _oneRenderResult;
#endif

    [Params(PdfHtmlPayloadKind.PlainText, PdfHtmlPayloadKind.Table, PdfHtmlPayloadKind.Multilingual)]
    public PdfHtmlPayloadKind Payload { get; set; }

    [GlobalSetup]
    public void Setup() {
        _scenario = PdfHtmlPayloadScenario.Create(Payload);
        _sharedFont = new PdfEmbeddedFontFamily(
            "Carlito",
            PdfBenchmarkAssets.CarlitoRegular,
            PdfBenchmarkAssets.CarlitoBold);
    }

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("library:OfficeIMO", "case:Default")]
    public byte[] OfficeIMO() =>
        _officeImoResult = OfficeImoPdfGenerator.GenerateHtml(_scenario.Html, _sharedFont);

    [Benchmark]
    [BenchmarkCategory("library:PeachPDF", "case:Default")]
    public byte[] PeachPDF() =>
        _peachPdfResult = PeachPdfGenerator.Generate(_scenario.Html, _sharedFont);

    [Benchmark]
    [BenchmarkCategory("library:PdfPecker", "case:Default")]
    public byte[] PdfPecker() =>
        _pdfPeckerResult = PdfPeckerGenerator.Generate(_scenario.Html, _sharedFont);

    // Validation stays in cleanup so the Dry job's first invocation remains a cold start; it checks the PDF the last
    // measured invocation produced and records the outcome.
#if NET10_0_OR_GREATER
    [Benchmark]
    [BenchmarkCategory("library:OneRender", "case:Default")]
    public byte[] OneRender() =>
        _oneRenderResult = OneRenderHtmlGenerator.Generate(_scenario.Html, OneRenderHtmlGenerator.CreateFontProvider(_sharedFont));

    [GlobalCleanup(Target = nameof(OneRender))]
    public void ValidateOneRender() => ValidateAndRecord(nameof(OneRender), _oneRenderResult);
#endif

    [GlobalCleanup(Target = nameof(PdfPecker))]
    public void ValidatePdfPecker() => ValidateAndRecord(nameof(PdfPecker), _pdfPeckerResult);

    [GlobalCleanup(Target = nameof(OfficeIMO))]
    public void ValidateOfficeIMO() => ValidateAndRecord(nameof(OfficeIMO), _officeImoResult);

    [GlobalCleanup(Target = nameof(PeachPDF))]
    public void ValidatePeachPDF() => ValidateAndRecord(nameof(PeachPDF), _peachPdfResult);

    private void ValidateAndRecord(string engine, byte[]? bytes) =>
        PdfBenchmarkValidation.Record(
            engine,
            "Default",
            $"Payload={Payload}; measured PDF reopened with PdfPig: 2-20 pages, every required text fragment, tag tree, embedded Carlito regular and bold",
            () => Validate(engine, bytes));

    /// <summary>Validates the same artifact contract for timings and the correctness-only evidence runner.</summary>
    internal PdfReadObservation Validate(string engine, byte[]? bytes) {
        if (bytes == null) {
            throw new InvalidDataException($"{engine} did not return a result for {_scenario.Kind}.");
        }

        PdfReadObservation observation = PdfBenchmarkValidation.ReadWithPdfPig(bytes);
        if (observation.PageCount < 2 || observation.PageCount > 20) {
            throw new InvalidDataException(
                $"{engine} produced an implausible {observation.PageCount} pages for the paginated 21 KiB {_scenario.Kind} workload.");
        }

        string normalized = _scenario.Kind == PdfHtmlPayloadKind.Multilingual
            ? NormalizeMultilingualText(observation.NormalizedText)
            : observation.NormalizedText;
        foreach (string required in _scenario.RequiredText) {
            string fragment = _scenario.Kind == PdfHtmlPayloadKind.Multilingual
                ? NormalizeMultilingualText(required)
                : PdfBenchmarkValidation.Normalize(required);
            bool preserved = normalized.Contains(fragment, StringComparison.Ordinal);
            if (!preserved && OfficeTextElements.ContainsRightToLeft(required)) {
                string visualOrder = string.Concat(OfficeTextElements.Enumerate(fragment).Reverse());
                preserved = normalized.Contains(visualOrder, StringComparison.Ordinal);
            }

            if (!preserved) {
                throw new InvalidDataException($"{engine} did not preserve required text '{required}' for {_scenario.Kind}.");
            }
        }

        ValidateTaggedStructure(engine, bytes);
        ValidateSharedFont(engine, bytes);

        Console.WriteLine(
            $"HTML_PDF_EVIDENCE engine={engine} payload={_scenario.Kind} htmlBytes={PdfHtmlPayloadScenario.TargetUtf8Bytes} " +
            $"pdfBytes={bytes.Length} pages={observation.PageCount} textLength={observation.TextLength}");
        return observation;
    }

    private static void ValidateTaggedStructure(string engine, byte[] bytes) {
        using var input = new MemoryStream(bytes, writable: false);
        using PdfDocument document = PdfDocument.Load(input);
        PdfDocumentInfo info = document.Inspect();
        PdfTaggedContentInfo? tagged = info.TaggedContent;
        if (!info.HasTaggedContent ||
            tagged == null ||
            tagged.StructureElements.Count == 0 ||
            tagged.MarkedContentReferenceCount == 0) {
            throw new InvalidDataException($"{engine} did not preserve a non-empty tagged structure tree.");
        }
    }

    private void ValidateSharedFont(string engine, byte[] bytes) {
        PdfEmbeddedFontFamily font = _sharedFont
            ?? throw new InvalidOperationException("The shared benchmark font was not initialized.");
        PdfBenchmarkValidation.ValidateSharedFontUse(bytes, engine, font.FamilyName, requireBold: true);
    }

    private static string NormalizeMultilingualText(string value) {
        string logical = OfficeArabicTextShaper.ToLogicalText(PdfBenchmarkValidation.Normalize(value));
        var normalized = new StringBuilder(logical.Length);
        foreach (char character in logical) {
            if (!char.IsWhiteSpace(character) &&
                CharUnicodeInfo.GetUnicodeCategory(character) != UnicodeCategory.Format) {
                normalized.Append(character);
            }
        }
        return normalized.ToString();
    }
}
