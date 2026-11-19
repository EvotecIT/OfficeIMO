using System.Text;
using System.Security.Cryptography;
using System.Text.Json;
using BenchmarkDotNet.Attributes;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>
/// Measures complete text extraction from prepared, independently sourced large PDFs.
/// Corpus preparation and SHA-256 validation run outside the measured operation.
/// </summary>
[MemoryDiagnoser]
[RankColumn]
public class PdfCorpusReadBenchmarks {
    private static readonly IReadOnlyDictionary<string, (string CorpusId, string FileName)> Documents =
        new Dictionary<string, (string, string)>(StringComparer.Ordinal) {
            ["OfficeIMO-500p"] = ("officeimo-native-large-500-page", "officeimo-native-large-500-page.pdf"),
            ["NIST-492p"] = ("nist-sp800-53-rev5", "nist-sp800-53-rev5.pdf"),
            ["Type3-85p"] = ("verapdf-type3-fonts", "verapdf-type3-fonts.pdf"),
            ["PDFA-258p-12MB"] = ("verapdf-large-pdfa", "verapdf-large-pdfa.pdf")
        };

    private byte[] _pdf = null!;

    public IEnumerable<string> DocumentValues => Documents.Keys;

    [ParamsSource(nameof(DocumentValues))]
    public string Document { get; set; } = string.Empty;

    [GlobalSetup]
    public void Setup() {
        string root = Environment.GetEnvironmentVariable("OFFICEIMO_PDF_CORPUS_ROOT") ?? string.Empty;
        if (string.IsNullOrWhiteSpace(root)) {
            throw new InvalidOperationException(
                "Set OFFICEIMO_PDF_CORPUS_ROOT to the prepared corpus files directory before running this suite.");
        }

        (string corpusId, string fileName) = Documents[Document];
        string path = Path.Combine(Path.GetFullPath(root), fileName);
        if (!File.Exists(path)) {
            throw new FileNotFoundException(
                $"Prepared corpus file '{fileName}' is missing. Run the corpus preparation workflow first.", path);
        }

        _pdf = File.ReadAllBytes(path);
        PdfCorpusEntry entry = LoadManifestEntry(corpusId);
        string sha256 = Convert.ToHexString(SHA256.HashData(_pdf)).ToLowerInvariant();
        if (!string.IsNullOrWhiteSpace(entry.Sha256) &&
            !string.Equals(entry.Sha256, sha256, StringComparison.OrdinalIgnoreCase)) {
            throw new InvalidDataException($"Corpus SHA-256 mismatch for {corpusId}. Expected {entry.Sha256}, observed {sha256}.");
        }
        PdfBenchmarkScenario? generatedScenario = null;
        if (string.Equals(entry.Generator, "officeimo-native-large", StringComparison.Ordinal)) {
            var scenario = new PdfBenchmarkScenario(
                PdfBenchmarkScale.High,
                "Large PDF reader and manipulation corpus",
                PageCount: 500,
                RowsPerPage: 4,
                ParagraphsPerPage: 1);
            string generatedSha256 = Convert.ToHexString(SHA256.HashData(OfficeImoPdfGenerator.Generate(scenario))).ToLowerInvariant();
            if (!string.Equals(generatedSha256, sha256, StringComparison.Ordinal)) {
                throw new InvalidDataException($"Generated corpus identity mismatch for {corpusId}. Expected {generatedSha256}, observed {sha256}.");
            }
            generatedScenario = scenario;
        }

        // Every reader must report the manifest's page count. The generated document's exact text is known, so each
        // reader is checked against it; for the downloaded documents OfficeIMO, iText and OnePDF must recover PdfPig's
        // tokens page by page, and PdfPig and iText must recover each other's tokens, at the manifest's recall threshold.
        PdfCorpusReadPayload pdfPig = ReadPdfPig();
        PdfCorpusReadPayload officeImo = ReadOfficeImo();
        PdfCorpusReadPayload iText = ReadIText();
        string detail = $"Document={Document}; pages={pdfPig.Pages.Count}; minimum token recall {entry.MinimumTokenRecall:0.00}";
        PdfBenchmarkValidation.Record("PdfPig", "Default", detail + "; recall against iText both ways", () => {
            if (entry.ExpectedPages.HasValue && pdfPig.Pages.Count != entry.ExpectedPages.Value) {
                throw new InvalidDataException($"Corpus page-count mismatch for {corpusId}. Expected {entry.ExpectedPages.Value}, observed {pdfPig.Pages.Count}.");
            }
            if (generatedScenario != null) {
                PdfBenchmarkValidation.ValidateRead(pdfPig.Observation, generatedScenario, "PdfPig reading " + Document);
            } else {
                ValidatePayload("iText (against PdfPig)", iText, pdfPig, entry.MinimumTokenRecall);
                ValidatePayload("PdfPig (against iText)", pdfPig, iText, entry.MinimumTokenRecall);
            }
        });
        RecordAgainstPdfPig("OfficeIMO", "Default", officeImo, pdfPig, generatedScenario, entry.MinimumTokenRecall, detail);
        RecordAgainstPdfPig("iText", "Default", iText, pdfPig, generatedScenario, entry.MinimumTokenRecall, detail);
#if NET10_0_OR_GREATER
        RecordAgainstPdfPig("OnePDF", "Default", ReadOnePdf(parallel: false), pdfPig, generatedScenario, entry.MinimumTokenRecall, detail);
        RecordAgainstPdfPig("OnePDF", "Parallel", ReadOnePdf(parallel: true), pdfPig, generatedScenario, entry.MinimumTokenRecall, detail);
#endif
    }

    private void RecordAgainstPdfPig(
        string library,
        string caseName,
        PdfCorpusReadPayload actual,
        PdfCorpusReadPayload pdfPig,
        PdfBenchmarkScenario? generatedScenario,
        double minimumTokenRecall,
        string detail) =>
        PdfBenchmarkValidation.Record(library, caseName, detail + "; recall against PdfPig", () => {
            string lane = caseName == "Default" ? library : library + " " + caseName;
            ValidatePayload(lane, actual, pdfPig, minimumTokenRecall);
            if (generatedScenario != null) {
                PdfBenchmarkValidation.ValidateRead(actual.Observation, generatedScenario, lane + " reading " + Document);
            }
        });

    private static PdfCorpusEntry LoadManifestEntry(string corpusId) {
        string copiedManifestPath = Path.Combine(AppContext.BaseDirectory, "Corpus", "pdf-corpus.json");
        if (File.Exists(copiedManifestPath)) {
            return ReadManifestEntry(copiedManifestPath, corpusId);
        }

        string? directory = AppContext.BaseDirectory;
        while (!string.IsNullOrWhiteSpace(directory)) {
            string manifestPath = Path.Combine(directory, "OfficeIMO.Pdf.Benchmarks.Comparisons", "Corpus", "pdf-corpus.json");
            if (File.Exists(manifestPath)) {
                return ReadManifestEntry(manifestPath, corpusId);
            }
            directory = Directory.GetParent(directory)?.FullName;
        }
        throw new FileNotFoundException("Could not locate the pinned PDF corpus manifest.");
    }

    private static PdfCorpusEntry ReadManifestEntry(string manifestPath, string corpusId) {
        PdfCorpusManifest manifest = JsonSerializer.Deserialize<PdfCorpusManifest>(
            File.ReadAllText(manifestPath),
            new JsonSerializerOptions { PropertyNameCaseInsensitive = true })
            ?? throw new InvalidDataException("PDF corpus manifest is empty.");
        return manifest.Entries.Single(entry => string.Equals(entry.Id, corpusId, StringComparison.Ordinal));
    }

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("library:OfficeIMO", "case:Default")]
    public PdfReadObservation OfficeIMO() => ReadOfficeImo().Observation;

    [Benchmark]
    [BenchmarkCategory("library:PdfPig", "case:Default")]
    public PdfReadObservation PdfPig() => ReadPdfPig().Observation;

    [Benchmark]
    [BenchmarkCategory("library:iText", "case:Default")]
    public PdfReadObservation IText() => ReadIText().Observation;

#if NET10_0_OR_GREATER
    [Benchmark]
    [BenchmarkCategory("library:OnePDF", "case:Default")]
    public PdfReadObservation OnePdf() => ReadOnePdf(parallel: false).Observation;

    /// <summary>OnePDF's default whole-document extraction, which reads up to four pages at once.</summary>
    [Benchmark]
    [BenchmarkCategory("library:OnePDF", "case:Parallel")]
    public PdfReadObservation OnePdfParallel() => ReadOnePdf(parallel: true).Observation;

    private PdfCorpusReadPayload ReadOnePdf(bool parallel) =>
        PdfCorpusReadPayload.Create(OnePdfComparisonEngines.ExtractPages(_pdf, parallel));
#endif

    private PdfCorpusReadPayload ReadOfficeImo() {
        var options = new global::OfficeIMO.Pdf.PdfLoadOptions { IncludeArtifactText = true };
        using var documentStream = new MemoryStream(_pdf, writable: false);
        using global::OfficeIMO.Pdf.PdfDocument document = global::OfficeIMO.Pdf.PdfDocument.Load(documentStream, options);
        string[] pages = document.Pages.Select(static page => page.ExtractText()).ToArray();
        return PdfCorpusReadPayload.Create(pages);
    }

    private PdfCorpusReadPayload ReadPdfPig() {
        using var stream = new MemoryStream(_pdf, writable: false);
        using UglyToad.PdfPig.PdfDocument document = UglyToad.PdfPig.PdfDocument.Open(stream);
        string[] pages = document.GetPages()
            .Select(static page => ContentOrderTextExtractor.GetText(page))
            .ToArray();
        return PdfCorpusReadPayload.Create(pages);
    }

    private PdfCorpusReadPayload ReadIText() {
        using var stream = new MemoryStream(_pdf, writable: false);
        using var reader = new iText.Kernel.Pdf.PdfReader(stream);
        using var document = new iText.Kernel.Pdf.PdfDocument(reader);
        var pages = new string[document.GetNumberOfPages()];
        for (int page = 1; page <= pages.Length; page++) {
            pages[page - 1] = PdfTextExtractor.GetTextFromPage(
                document.GetPage(page),
                new LocationTextExtractionStrategy());
        }

        return PdfCorpusReadPayload.Create(pages);
    }

    private static void ValidatePayload(
        string engine,
        PdfCorpusReadPayload actual,
        PdfCorpusReadPayload oracle,
        double minimumRecall) {
        if (actual.Pages.Count != oracle.Pages.Count) {
            throw new InvalidDataException(
                $"{engine} observed {actual.Pages.Count} pages while PdfPig observed {oracle.Pages.Count}.");
        }

        double recall = PdfCorpusRunner.TokenRecall(oracle.Pages, actual.Pages);
        if (recall < minimumRecall) {
            throw new InvalidDataException(
                $"{engine} token recall {recall:P2} is below the {minimumRecall:P2} parity threshold for the prepared corpus.");
        }
    }

    private sealed record PdfCorpusReadPayload(IReadOnlyList<string> Pages, PdfReadObservation Observation) {
        internal static PdfCorpusReadPayload Create(IReadOnlyList<string> pages) {
            var text = new StringBuilder();
            foreach (string page in pages) {
                text.Append(page);
                text.Append('\n');
            }

            return new PdfCorpusReadPayload(pages, PdfBenchmarkValidation.Observe(pages.Count, text.ToString()));
        }
    }
}
