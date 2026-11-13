using System.Globalization;
using System.Text;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>
/// Scores page text against the exact logical text of the generated report scenario: title, narrative
/// paragraphs, then table rows with cells in column order. Every producer writes the same model, so the truth
/// is known without labelling. Content error ignores order (missing or extra words); order error is the word
/// edit distance against the logical sequence, so a reader that keeps every word but interleaves table cells
/// scores zero content error and a positive order error.
/// </summary>
internal static class PdfTextQualityRunner {
    private static readonly PdfBenchmarkScale[] Scales = [PdfBenchmarkScale.Easy, PdfBenchmarkScale.Medium];

    internal static int Run(string[] args) {
        var engines = new List<(string Name, Func<byte[], IReadOnlyList<string>> Pages)> {
            ("OfficeIMO", OfficeImoPages),
            ("PdfPig (content order)", PdfPigPages),
            ("iText (location strategy)", ITextPages)
        };
#if NET10_0_OR_GREATER
        engines.Add(("OnePdf (reading order)", OnePdfComparisonEngines.ExtractPages));
#endif
        var report = new StringBuilder();
        report.AppendLine("| Producer | Scale | Reader | Pages | Content error % | Order error % |");
        report.AppendLine("| --- | --- | --- | ---: | ---: | ---: |");
        var totals = engines.ToDictionary(engine => engine.Name, _ => (Content: 0L, Order: 0L, Words: 0L));
        foreach (PdfBenchmarkScale scale in Scales) {
            PdfBenchmarkScenario scenario = PdfBenchmarkScenario.Get(scale);
            foreach (PdfBenchmarkProducer producer in Enum.GetValues<PdfBenchmarkProducer>()) {
                byte[] pdf = PdfPreparedReadInputs.Resolve(producer, scenario);
                foreach ((string name, Func<byte[], IReadOnlyList<string>> read) in engines) {
                    IReadOnlyList<string> pages = read(pdf);
                    if (pages.Count != scenario.PageCount) {
                        throw new InvalidDataException($"{name} returned {pages.Count} pages for {producer}/{scale}.");
                    }
                    long content = 0L, order = 0L, words = 0L;
                    for (int page = 1; page <= scenario.PageCount; page++) {
                        string[] truth = Words(Truth(scenario, page));
                        string[] actual = Words(pages[page - 1]);
                        content += ContentErrors(truth, actual);
                        order += EditDistance(truth, actual);
                        words += truth.Length;
                    }
                    var total = totals[name];
                    totals[name] = (total.Content + content, total.Order + order, total.Words + words);
                    report.AppendLine(string.Create(CultureInfo.InvariantCulture,
                        $"| {producer} | {scale} | {name} | {scenario.PageCount} | {Percent(content, words)} | {Percent(order, words)} |"));
                }
            }
        }
        report.AppendLine();
        report.AppendLine("| Reader | Content error % (all inputs) | Order error % (all inputs) |");
        report.AppendLine("| --- | ---: | ---: |");
        foreach ((string name, (long content, long order, long words)) in totals) {
            report.AppendLine($"| {name} | {Percent(content, words)} | {Percent(order, words)} |");
        }
        Console.Write(report.ToString());
        if (args.Length == 2 && args[0] == "--output") File.WriteAllText(args[1], report.ToString());
        return 0;
    }

    /// <summary>The page's logical text in reading order.</summary>
    internal static string Truth(PdfBenchmarkScenario scenario, int page) {
        var text = new StringBuilder(scenario.PageTitle(page)).Append('\n');
        for (int paragraph = 0; paragraph < scenario.ParagraphsPerPage; paragraph++) {
            text.Append(scenario.Narrative(page, paragraph)).Append('\n');
        }
        foreach (string[] row in scenario.TableRows(page)) text.AppendLine(string.Join(" ", row));
        return text.ToString();
    }

    private static string[] Words(string text) =>
        text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);

    /// <summary>Missing plus extra words, ignoring order.</summary>
    private static long ContentErrors(string[] truth, string[] actual) {
        var counts = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (string word in truth) counts[word] = counts.TryGetValue(word, out int count) ? count + 1 : 1;
        long extra = 0L;
        foreach (string word in actual) {
            if (counts.TryGetValue(word, out int count) && count > 0) counts[word] = count - 1;
            else extra++;
        }
        return extra + counts.Values.Sum();
    }

    private static int EditDistance(string[] truth, string[] actual) {
        var previous = new int[actual.Length + 1];
        var current = new int[actual.Length + 1];
        for (int j = 0; j <= actual.Length; j++) previous[j] = j;
        for (int i = 1; i <= truth.Length; i++) {
            current[0] = i;
            for (int j = 1; j <= actual.Length; j++) {
                int substitute = previous[j - 1] + (string.Equals(truth[i - 1], actual[j - 1], StringComparison.Ordinal) ? 0 : 1);
                current[j] = Math.Min(substitute, Math.Min(previous[j] + 1, current[j - 1] + 1));
            }
            (previous, current) = (current, previous);
        }
        return previous[actual.Length];
    }

    private static string Percent(long errors, long words) =>
        (100D * errors / Math.Max(1L, words)).ToString("0.0", CultureInfo.InvariantCulture);

    private static IReadOnlyList<string> OfficeImoPages(byte[] pdf) {
        using var stream = new MemoryStream(pdf, writable: false);
        using global::OfficeIMO.Pdf.PdfDocument document = global::OfficeIMO.Pdf.PdfDocument.Load(stream);
        return document.Pages.Select(static page => page.ExtractText()).ToArray();
    }

    private static IReadOnlyList<string> PdfPigPages(byte[] pdf) {
        using UglyToad.PdfPig.PdfDocument document = UglyToad.PdfPig.PdfDocument.Open(pdf);
        return document.GetPages().Select(static page => ContentOrderTextExtractor.GetText(page)).ToArray();
    }

    private static IReadOnlyList<string> ITextPages(byte[] pdf) {
        using var stream = new MemoryStream(pdf, writable: false);
        using var reader = new iText.Kernel.Pdf.PdfReader(stream);
        using var document = new iText.Kernel.Pdf.PdfDocument(reader);
        var pages = new string[document.GetNumberOfPages()];
        for (int page = 1; page <= pages.Length; page++) {
            pages[page - 1] = PdfTextExtractor.GetTextFromPage(document.GetPage(page), new LocationTextExtractionStrategy());
        }
        return pages;
    }
}
