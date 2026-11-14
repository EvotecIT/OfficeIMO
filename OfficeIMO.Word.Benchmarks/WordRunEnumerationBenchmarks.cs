using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Word.Benchmarks;

/// <summary>
/// Loads a document and enumerates the runs of every paragraph through the public GetRuns, at
/// 1,000 and 4,000 paragraphs. Four times the paragraphs should cost about four times as much.
/// </summary>
[MemoryDiagnoser]
[BenchmarkCategory("Word", "Runs")]
public class WordRunEnumerationBenchmarks {
    private const int FieldInterval = 100;
    private byte[] _document = Array.Empty<byte>();

    [Params(1000, 4000)]
    public int ParagraphCount { get; set; }

    [GlobalSetup]
    public void Setup() {
        using (WordDocument document = WordDocument.Create()) {
            for (int index = 0; index < ParagraphCount; index++) {
                WordParagraph paragraph = document.AddParagraph("Paragraph " + index + " ").AddText("second run");
                if (index % FieldInterval == 0) paragraph.AddField(WordFieldType.Author);
            }
            _document = document.ToBytes();
        }

        // Document.Paragraphs lists one entry per run, and each entry enumerates its paragraph's
        // runs, so two text runs per paragraph give at least four.
        int runs = EnumerateRuns();
        if (runs < ParagraphCount * 4) {
            throw new InvalidOperationException("Enumerating found " + runs + " runs; expected at least " + ParagraphCount * 4 + ".");
        }
    }

    [Benchmark]
    public int EnumerateRuns() {
        using var stream = new MemoryStream(_document, writable: false);
        using WordDocument document = WordDocument.Load(stream);
        int runs = 0;
        foreach (WordParagraph paragraph in document.Paragraphs) {
            foreach (WordParagraph _ in paragraph.GetRuns()) runs++;
        }
        return runs;
    }
}
