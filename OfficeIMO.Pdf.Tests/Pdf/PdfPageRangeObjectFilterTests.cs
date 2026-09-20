using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfPageRangeObjectFilterTests {
    [Fact]
    public void FilterPageLabels_UsesLatestRuleAcrossManySelectedPages() {
        PdfPageLabel[] rules = Enumerable.Range(0, 500)
            .Select(index => new PdfPageLabel(index, "D", "page-", index + 1))
            .ToArray();
        int[] selectedPages = Enumerable.Range(0, 125)
            .Select(index => 500 - index * 4)
            .ToArray();

        IReadOnlyList<PdfPageLabel> filtered = PdfPageRangeObjectFilter.FilterPageLabelsByPageNumbers(rules, selectedPages);

        Assert.Equal(selectedPages.Length, filtered.Count);
        for (int index = 0; index < filtered.Count; index++) {
            int sourcePage = selectedPages[selectedPages.Length - index - 1];
            Assert.Equal(sourcePage - 1, filtered[index].StartPageIndex);
            Assert.Equal(sourcePage, filtered[index].StartNumber);
            Assert.Equal("page-", filtered[index].Prefix);
        }
    }

    [Fact]
    public void FilterPageLabels_PicksLastRuleAtDuplicateStartIndex() {
        PdfPageLabel[] rules = {
            new PdfPageLabel(0, "D", "first-", 1),
            new PdfPageLabel(4, "D", "old-", 10),
            new PdfPageLabel(4, "D", "new-", 20)
        };

        PdfPageLabel filtered = Assert.Single(PdfPageRangeObjectFilter.FilterPageLabelsByPageNumbers(rules, new[] { 5 }));

        Assert.Equal(4, filtered.StartPageIndex);
        Assert.Equal(20, filtered.StartNumber);
        Assert.Equal("new-", filtered.Prefix);
    }
}
