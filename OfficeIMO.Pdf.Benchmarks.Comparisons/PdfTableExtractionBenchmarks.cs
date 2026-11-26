#if !NETFRAMEWORK
using BenchmarkDotNet.Attributes;
using Tabula;
using Tabula.Extractors;
using UglyToad.PdfPig;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>
/// Compares extracting the ruled tables of every page from identical bytes: OfficeIMO through its structured
/// read, Tabula through its lattice (ruling-line) extractor. OfficeIMO reads the painted grid, except in
/// PeachPDF's output, which is tagged: there its structured read takes the tables from the structure tree, as
/// it does for any tagged file. Each run must return every row of the source scenario with its exact cells.
/// </summary>
[MemoryDiagnoser]
[RankColumn]
public class PdfTableExtractionBenchmarks {
    private static readonly PdfLoadOptions OfficeImoOptions = new() { IncludeImages = false, IncludeFormWidgets = false };
    private byte[] _pdf = null!;
    private PdfBenchmarkScenario _scenario = null!;

    [Params(PdfBenchmarkScale.Easy, PdfBenchmarkScale.Medium)]
    public PdfBenchmarkScale Scale { get; set; }

    // Producers whose tables carry a complete drawn grid. QuestPDF's scenario draws row rules only, which
    // a ruling-line extractor cannot read as columns, so it would not be equivalent work.
    [Params(PdfBenchmarkProducer.PeachPDF, PdfBenchmarkProducer.MigraDoc, PdfBenchmarkProducer.IText)]
    public PdfBenchmarkProducer Producer { get; set; }

    [GlobalSetup]
    public void Setup() {
        _scenario = PdfBenchmarkScenario.Get(Scale);
        _pdf = PdfPreparedReadInputs.Resolve(Producer, _scenario);
        Validate(nameof(OfficeIMO), OfficeIMO());
        Validate(nameof(Tabula), Tabula());
    }

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("library:OfficeIMO", "case:Default")]
    public IReadOnlyList<string[]> OfficeIMO() {
        using PdfDocument document = PdfDocument.Load(_pdf);
        var rows = new List<string[]>();
        foreach (PdfLogicalTable table in document.Read(OfficeImoOptions).Tables) {
            foreach (IReadOnlyList<string> row in table.Rows) rows.Add(row.ToArray());
        }
        return rows;
    }

    [Benchmark]
    [BenchmarkCategory("library:Tabula", "case:Default")]
    public IReadOnlyList<string[]> Tabula() {
        using UglyToad.PdfPig.PdfDocument document = UglyToad.PdfPig.PdfDocument.Open(_pdf, new ParsingOptions { ClipPaths = true });
        var extractor = new SpreadsheetExtractionAlgorithm();
        var rows = new List<string[]>();
        for (int page = 1; page <= document.NumberOfPages; page++) {
            foreach (Table table in extractor.Extract(ObjectExtractor.Extract(document, page))) {
                foreach (IReadOnlyList<Cell> row in table.Rows) rows.Add(row.Select(static cell => cell.GetText()).ToArray());
            }
        }
        return rows;
    }

    // Every source row of every page must come back as one extracted row with the same cells, in any order.
    // Whitespace inside a cell is collapsed and empty cells at either edge are ignored, as layouts differ there.
    private void Validate(string engine, IReadOnlyList<string[]> extracted) =>
        PdfBenchmarkValidation.Record(
            PdfBenchmarkValidation.LibraryId(engine),
            PdfBenchmarkValidation.CaseId(engine),
            $"Producer={Producer}; Scale={Scale}; every source table row extracted with its exact cells",
            () => {
                var available = new Dictionary<string, int>(StringComparer.Ordinal);
                foreach (string[] row in extracted) {
                    string key = Key(row);
                    available[key] = available.TryGetValue(key, out int count) ? count + 1 : 1;
                }
                for (int page = 1; page <= _scenario.PageCount; page++) {
                    foreach (string[] row in _scenario.TableRows(page)) {
                        string key = Key(row);
                        if (!available.TryGetValue(key, out int count) || count == 0) {
                            throw new InvalidDataException($"{engine} reading {Producer} missed table row '{string.Join(" | ", row)}' on page {page}.");
                        }
                        available[key] = count - 1;
                    }
                }
            });

    private static string Key(IEnumerable<string> cells) {
        string[] normalized = cells.Select(static cell =>
            string.Join(" ", (cell ?? string.Empty).Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries))).ToArray();
        int first = 0, last = normalized.Length;
        while (first < last && normalized[first].Length == 0) first++;
        while (last > first && normalized[last - 1].Length == 0) last--;
        return string.Join("\u001f", normalized, first, last - first);
    }
}
#endif
