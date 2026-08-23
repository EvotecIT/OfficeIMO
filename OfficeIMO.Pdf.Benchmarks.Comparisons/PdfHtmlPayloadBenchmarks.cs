using BenchmarkDotNet.Attributes;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
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
    private PdfEmbeddedFontFamily? _multilingualFont;
    private byte[]? _officeImoResult;
    private byte[]? _peachPdfResult;

    [Params(PdfHtmlPayloadKind.PlainText, PdfHtmlPayloadKind.Table, PdfHtmlPayloadKind.Multilingual)]
    public PdfHtmlPayloadKind Payload { get; set; }

    [GlobalSetup]
    public void Setup() {
        _scenario = PdfHtmlPayloadScenario.Create(Payload);
        if (Payload == PdfHtmlPayloadKind.Multilingual) {
            _multilingualFont = new PdfEmbeddedFontFamily(
                "Carlito",
                PdfBenchmarkAssets.CarlitoRegular);
        }
    }

    [Benchmark(Baseline = true)]
    public byte[] OfficeIMO() =>
        _officeImoResult = OfficeImoPdfGenerator.GenerateHtml(_scenario.Html, _multilingualFont);

    [Benchmark]
    public byte[] PeachPDF() =>
        _peachPdfResult = PeachPdfGenerator.Generate(_scenario.Html, _multilingualFont);

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
        if (_scenario.Kind == PdfHtmlPayloadKind.Multilingual) {
            ValidateSharedMultilingualFont(engine, bytes);
        }

        Console.WriteLine(
            $"HTML_PDF_EVIDENCE engine={engine} payload={_scenario.Kind} htmlBytes={PdfHtmlPayloadScenario.TargetUtf8Bytes} " +
            $"pdfBytes={bytes.Length} pages={observation.PageCount} textLength={observation.TextLength}");
    }

    private static void ValidateTaggedStructure(string engine, byte[] bytes) {
        PdfDocumentInfo info = PdfDocument.Open(bytes).Inspect();
        PdfTaggedContentInfo? tagged = info.TaggedContent;
        if (!info.HasTaggedContent ||
            tagged == null ||
            tagged.StructureElements.Count == 0 ||
            tagged.MarkedContentReferenceCount == 0) {
            throw new InvalidDataException($"{engine} did not preserve a non-empty tagged structure tree.");
        }
    }

    private void ValidateSharedMultilingualFont(string engine, byte[] bytes) {
        PdfEmbeddedFontFamily font = _multilingualFont
            ?? throw new InvalidOperationException("The multilingual benchmark font was not initialized.");
        string expectedFont = NormalizeFontName(font.FamilyName);
        bool embedded = Encoding.Latin1.GetString(bytes).Contains("/FontFile", StringComparison.Ordinal);
        bool usedForGreek = false;
        using (var stream = new MemoryStream(bytes, writable: false)) {
            using UglyToad.PdfPig.PdfDocument document = UglyToad.PdfPig.PdfDocument.Open(stream);
            usedForGreek = document.GetPages()
                .SelectMany(page => page.Letters)
                .Any(letter =>
                    letter.Value.Contains('Ε') &&
                    NormalizeFontName(letter.FontName ?? string.Empty).Contains(expectedFont, StringComparison.Ordinal));
        }

        if (!embedded || !usedForGreek) {
            throw new InvalidDataException(
                $"{engine} did not embed and use the shared '{font.FamilyName}' font for multilingual text.");
        }
    }

    private static string NormalizeFontName(string value) =>
        new(value.Where(char.IsLetterOrDigit).Select(char.ToUpperInvariant).ToArray());

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
