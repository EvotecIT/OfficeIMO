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
        }
        PdfCorpusReadPayload oracle = ReadPdfPig();
        if (entry.ExpectedPages.HasValue && oracle.Pages.Count != entry.ExpectedPages.Value) {
            throw new InvalidDataException($"Corpus page-count mismatch for {corpusId}. Expected {entry.ExpectedPages.Value}, observed {oracle.Pages.Count}.");
        }
        ValidatePayload("OfficeIMO", ReadOfficeImo(), oracle, entry.MinimumTokenRecall);
        ValidatePayload("iText", ReadIText(), oracle, entry.MinimumTokenRecall);
    }

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
    public PdfReadObservation OfficeIMO() => ReadOfficeImo().Observation;

    [Benchmark]
    public PdfReadObservation PdfPig() => ReadPdfPig().Observation;

    [Benchmark]
    public PdfReadObservation IText() => ReadIText().Observation;

    private PdfCorpusReadPayload ReadOfficeImo() {
        var options = new global::OfficeIMO.Pdf.PdfLoadOptions { IncludeArtifactText = true };
        global::OfficeIMO.Pdf.PdfReadDocument document = global::OfficeIMO.Pdf.PdfReadDocument.Open(_pdf, options);
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
