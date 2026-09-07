using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfPageReorderPlanTests {
    [Theory]
    [InlineData(1, new[] { 4, 2 }, new[] { 2, 4, 1, 3, 5 })]
    [InlineData(6, new[] { 4, 2 }, new[] { 1, 3, 5, 2, 4 })]
    [InlineData(3, new[] { 1, 2 }, new[] { 1, 2, 3, 4, 5 })]
    public void PlannedMoveMatchesWrittenPageOrder(int before, int[] selected, int[] expected) {
        var document = OfficeIMO.Pdf.PdfDocument.Create(compose => {
            for (int page = 1; page <= 5; page++) {
                int number = page;
                compose.Page(builder => builder.Content(content => content.Item(item => item.Text("Page " + number))));
            }
        });
        var plan = PdfPageReorderPlan.Move(5, before, selected);
        Assert.Equal(expected, plan.SourcePageNumbers);
        Assert.Equal(expected.Select(page => "Page " + page), PdfReadDocument.Open(document.Pages.Move(before, selected).ToBytes()).Pages.Select(page => page.ExtractText().Trim()));
        foreach (int page in selected) Assert.Equal(Array.IndexOf(expected, page) + 1, plan.GetOutputPageNumber(page));
        Assert.Equal(!expected.SequenceEqual(new[] { 1, 2, 3, 4, 5 }), plan.HasChanges);
    }

    [Theory]
    [InlineData(true, new[] { 1, 3, 4, 6 }, new[] { 1, 3, 4, 2, 6, 5 })]
    [InlineData(false, new[] { 1, 3, 4, 6 }, new[] { 2, 1, 5, 3, 4, 6 })]
    [InlineData(true, new[] { 1, 2, 3, 4, 5, 6 }, new[] { 1, 2, 3, 4, 5, 6 })]
    public void ShiftKeepsRunsTogetherAndLeavesEdgeRunsInPlace(bool towardStart, int[] selected, int[] expected) {
        var plan = PdfPageReorderPlan.Shift(6, towardStart, selected);
        Assert.Equal(expected, plan.SourcePageNumbers);
        Assert.Equal(!expected.SequenceEqual(new[] { 1, 2, 3, 4, 5, 6 }), plan.HasChanges);
    }

    [Fact]
    public void InvalidSelectionsAndDestinationsAreRejected() {
        Assert.Throws<ArgumentException>(() => PdfPageReorderPlan.Move(4, 2, 2));
        Assert.Throws<ArgumentException>(() => PdfPageReorderPlan.Move(4, 3, 1, 1));
        Assert.Throws<ArgumentException>(() => PdfPageReorderPlan.Shift(4, true));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageReorderPlan.Move(4, 6, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageReorderPlan.Shift(4, false, 5));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageReorderPlan.Shift(0, true, 1));
    }

    [Fact]
    public void RangeResolutionPreservesOrderAndBoundsWithoutOverflowingAtTheLastInteger() {
        Assert.Equal(new[] { 3, 1, 2, 3 }, PdfPageSelection.Parse("3,1-3").Resolve(3));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageSelection.Parse("1-4").Resolve(3));
        Assert.Equal(new[] { int.MaxValue }, PdfPageSelection.From(int.MaxValue).Resolve(int.MaxValue));
    }
}
