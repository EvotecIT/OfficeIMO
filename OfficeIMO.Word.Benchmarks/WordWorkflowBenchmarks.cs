using BenchmarkDotNet.Attributes;
using OfficeIMO.Word.Html;

namespace OfficeIMO.Word.Benchmarks;

/// <summary>Measures allocation-sensitive Word 3.1 workflows over deterministic package corpora.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("Word", "Workflows")]
public class WordWorkflowBenchmarks {
    private byte[] _fieldDocument = Array.Empty<byte>();
    private byte[] _mergeDocument = Array.Empty<byte>();
    private byte[] _htmlDocument = Array.Empty<byte>();
    private WordDocument _loadedHtmlDocument = null!;
    private Dictionary<string, string> _mergeValues = null!;
    private string _sourcePath = string.Empty;
    private string _targetPath = string.Empty;
    private string _temporaryDirectory = string.Empty;

    [Params(100, 1000)]
    public int ItemCount { get; set; }

    [GlobalSetup]
    public void Setup() {
        _temporaryDirectory = Path.Combine(Path.GetTempPath(), "OfficeIMO.Word.Benchmarks", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_temporaryDirectory);
        _sourcePath = Path.Combine(_temporaryDirectory, "source.docx");
        _targetPath = Path.Combine(_temporaryDirectory, "target.docx");
        _mergeValues = new Dictionary<string, string>(ItemCount, StringComparer.OrdinalIgnoreCase);

        using (WordDocument fields = WordDocument.Create()) {
            fields.BuiltinDocumentProperties.Creator = "Benchmark";
            for (int index = 0; index < ItemCount; index++) fields.AddParagraph("Author: ").AddField(WordFieldType.Author);
            _fieldDocument = fields.ToBytes();
        }

        using (WordDocument merge = WordDocument.Create()) {
            for (int index = 0; index < ItemCount; index++) {
                string name = "Value" + index;
                merge.AddParagraph(name + ": ").AddField(WordFieldType.MergeField, parameters: new List<string> { name });
                _mergeValues[name] = index.ToString(System.Globalization.CultureInfo.InvariantCulture);
            }
            _mergeDocument = merge.ToBytes();
        }

        using (WordDocument html = WordDocument.Create()) {
            for (int index = 0; index < ItemCount; index++) html.AddParagraph("Paragraph " + index);
            _htmlDocument = html.ToBytes();
        }
        _loadedHtmlDocument = WordDocument.Load(new MemoryStream(_htmlDocument, writable: false));

        using (WordDocument source = WordDocument.Create(_sourcePath)) {
            using WordDocument target = WordDocument.Create(_targetPath);
            for (int index = 0; index < ItemCount; index++) {
                source.AddParagraph("Paragraph " + index);
                target.AddParagraph(index == ItemCount / 2 ? "Changed paragraph" : "Paragraph " + index);
            }
            source.Save();
            target.Save();
        }

        ValidateWorkflowResults();
    }

    [Benchmark]
    public int UpdateFields() {
        using var stream = new MemoryStream(_fieldDocument, writable: false);
        using WordDocument document = WordDocument.Load(stream);
        return document.UpdateFieldsAndGetReport().UpdatedCount;
    }

    [Benchmark]
    public int ExecuteMailMerge() {
        using var stream = new MemoryStream(_mergeDocument, writable: false);
        using WordDocument document = WordDocument.Load(stream);
        return WordMailMerge.ExecuteWithReport(document, _mergeValues).MergedCount;
    }

    [Benchmark]
    public int ExportHtml() {
        using var stream = new MemoryStream(_htmlDocument, writable: false);
        using WordDocument document = WordDocument.Load(stream);
        return document.ToHtmlResult().RequireValue().Length;
    }

    [Benchmark]
    public int LoadDocument() {
        using var stream = new MemoryStream(_htmlDocument, writable: false);
        using WordDocument document = WordDocument.Load(stream);
        return document.Paragraphs.Count;
    }

    [Benchmark]
    public int ExportLoadedHtml() => _loadedHtmlDocument.ToHtmlResult().RequireValue().Length;

    [Benchmark]
    public int CompareStructure() => WordDocumentComparer.CompareStructure(_sourcePath, _targetPath).Findings.Count;

    private void ValidateWorkflowResults() {
        EnsureEqual(nameof(UpdateFields), ItemCount, UpdateFields());
        EnsureEqual(nameof(ExecuteMailMerge), ItemCount, ExecuteMailMerge());
        EnsureEqual(nameof(LoadDocument), ItemCount, LoadDocument());
        EnsurePositive(nameof(ExportHtml), ExportHtml());
        EnsurePositive(nameof(ExportLoadedHtml), ExportLoadedHtml());

        WordComparisonResult comparison = WordDocumentComparer.CompareStructure(_sourcePath, _targetPath);
        if (!comparison.HasChanges ||
            !comparison.Findings.Any(finding =>
                string.Equals(finding.TargetText, "Changed paragraph", StringComparison.Ordinal))) {
            throw new InvalidOperationException("CompareStructure validation did not find the expected changed paragraph.");
        }
    }

    private static void EnsureEqual(string benchmark, int expected, int actual) {
        if (actual != expected) {
            throw new InvalidOperationException(benchmark + " validation returned " + actual + "; expected " + expected + ".");
        }
    }

    private static void EnsurePositive(string benchmark, int actual) {
        if (actual <= 0) {
            throw new InvalidOperationException(benchmark + " validation returned no output.");
        }
    }

    [GlobalCleanup]
    public void Cleanup() {
        _loadedHtmlDocument.Dispose();
        if (Directory.Exists(_temporaryDirectory)) Directory.Delete(_temporaryDirectory, recursive: true);
    }
}
