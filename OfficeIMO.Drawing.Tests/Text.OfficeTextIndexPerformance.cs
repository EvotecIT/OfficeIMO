using System.Diagnostics;
using OfficeIMO.TextAnalysis;
using Xunit;
using Xunit.Abstractions;

namespace OfficeIMO.Tests;

[Trait("Category", "Performance")]
public class TextOfficeTextIndexPerformanceTests {
    private readonly ITestOutputHelper _output;

    public TextOfficeTextIndexPerformanceTests(ITestOutputHelper output) => _output = output;

    [Fact]
    public void MeasuresIndexingAndQueryingSyntheticChunks() {
        string[] vocabulary = Enumerable.Range(0, 5_000).Select(static index => "term" + index).ToArray();
        var random = new Random(29);
        string[] documents = Enumerable.Range(0, 20_000)
            .Select(_ => string.Join(" ", Enumerable.Range(0, 120).Select(_ => vocabulary[random.Next(vocabulary.Length)])))
            .ToArray();

        long before = GC.GetAllocatedBytesForCurrentThread();
        var watch = Stopwatch.StartNew();
        var index = new OfficeTextIndex();
        for (int i = 0; i < documents.Length; i++) index.Add("d" + i, documents[i]);
        TimeSpan indexing = watch.Elapsed;
        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;

        watch.Restart();
        int matches = 0;
        foreach (string query in new[] { "term1 term2", "term3 OR term4", "\"term5 term6\"", "term1*" }) {
            matches += index.Search(query).TotalMatches;
        }
        TimeSpan querying = watch.Elapsed;

        _output.WriteLine($"Indexed {index.DocumentCount} documents, {index.TokenCount} tokens, {index.TermCount} terms in {indexing.TotalMilliseconds:F0} ms, {allocated / (1024 * 1024)} MB allocated.");
        _output.WriteLine($"Four queries, {matches} matches, in {querying.TotalMilliseconds:F1} ms.");
        Assert.Equal(20_000, index.DocumentCount);
    }
}
