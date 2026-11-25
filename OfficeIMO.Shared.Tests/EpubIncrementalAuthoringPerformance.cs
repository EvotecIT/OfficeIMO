using OfficeIMO.Epub;
using System.Diagnostics;
using Xunit;
using Xunit.Abstractions;

namespace OfficeIMO.Shared.Tests;

// Builds fixed-layout books page by page (chapter, image, page canvas), the way a comic or scan
// converter does, and checks that the cost of adding a page doesn't grow with the book.
public sealed class EpubIncrementalAuthoringPerformance {
    private readonly ITestOutputHelper _output;

    public EpubIncrementalAuthoringPerformance(ITestOutputHelper output) => _output = output;

    [Fact]
    [Trait("Category", "Performance")]
    public void AddingAPageCostsTheSameAtTwoThousandPagesAsAtFifty() {
        EpubIncrementalAuthoringContracts.PageBook(20);
        double small = Measure(50);
        Measure(500);
        double large = Measure(2000);
        // Whole-book validation on every page made the last pages of a 2,000-page book hundreds of times
        // slower than the first ones. The ratio is loose because shared runners are noisy.
        Assert.True(large < small * 5, $"Per-page cost grew from {small:F3} ms at 50 pages to {large:F3} ms at 2,000 pages.");
    }

    // Returns the mean cost of the last pages added, in milliseconds per page.
    private double Measure(int pages) {
        int measured = Math.Max(25, pages / 10);
        var book = EpubIncrementalAuthoringContracts.PageBook(pages - measured);
        var stopwatch = Stopwatch.StartNew();
        for (int index = pages - measured + 1; index <= pages; index++) EpubIncrementalAuthoringContracts.AddPage(book, index);
        stopwatch.Stop();
        double perPage = stopwatch.Elapsed.TotalMilliseconds / measured;
        var save = Stopwatch.StartNew();
        byte[] bytes = book.Write().Bytes;
        save.Stop();
        _output.WriteLine($"{pages,5} pages: {perPage:F3} ms per page over the last {measured}; Write {save.Elapsed.TotalMilliseconds:F0} ms, {bytes.Length:N0} bytes");
        return perPage;
    }
}
