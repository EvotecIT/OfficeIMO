using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfPageSelectorTests {
    [Theory]
    [InlineData("last", new[] { 6 })]
    [InlineData("end-2", new[] { 4 })]
    [InlineData("z-1", new[] { 5 })]
    [InlineData("5-2", new[] { 5, 4, 3, 2 })]
    [InlineData("last..last-3", new[] { 6, 5, 4, 3 })]
    [InlineData("odd", new[] { 1, 3, 5 })]
    [InlineData("even", new[] { 2, 4, 6 })]
    [InlineData("all,!1,!last", new[] { 2, 3, 4, 5 })]
    [InlineData("!odd", new[] { 2, 4, 6 })]
    [InlineData("last,1..3,!2", new[] { 6, 1, 3 })]
    public void Resolve_SupportsRelativeReverseParityAndExclusionTerms(string expression, int[] expected) {
        PdfPageSelector selector = PdfPageSelector.Parse(expression);

        Assert.Equal(expected, selector.Resolve(6));
        Assert.Equal(expected, selector.ResolveSelection(6).ToString().Split(',').Select(int.Parse));
    }

    [Fact]
    public void Parse_PreservesOrderedDuplicatesUntilAnOperationChoosesToDeduplicate() {
        PdfPageSelector selector = PdfPageSelector.Parse("3,1..2,2");

        Assert.Equal(new[] { 3, 1, 2, 2 }, selector.Resolve(4));
    }

    [Theory]
    [InlineData("")]
    [InlineData("1,,2")]
    [InlineData("!")]
    [InlineData("last+")]
    [InlineData("0")]
    public void TryParse_RejectsMalformedSelectors(string expression) {
        Assert.False(PdfPageSelector.TryParse(expression, out _));
    }

    [Fact]
    public void Resolve_RejectsOutOfRangeAndEmptySelections() {
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageSelector.Parse("last-6").Resolve(6));
        Assert.Throws<InvalidOperationException>(() => PdfPageSelector.Parse("all,!all").Resolve(6));
    }

    [Fact]
    public void DocumentOperations_ResolveSelectorsAgainstTheActualDocument() {
        byte[] source = CreateSixPageDocument();
        PdfPageSelector reverseWithoutFourth = PdfPageSelector.Parse("last..2,!4");

        IReadOnlyList<string> selectedText = PdfDocument.Open(source).Read.TextByPage(reverseWithoutFourth);
        PdfDocument extracted = PdfDocument.Open(source).Pages.Extract(reverseWithoutFourth);
        PdfDocument deleted = PdfDocument.Open(source).Pages.Delete(PdfPageSelector.Parse("odd"));
        PdfOperationResult<PdfDocument> reordered = PdfDocument.Open(source).Pages.TryReorder(PdfPageSelector.Parse("last..1"));

        Assert.Equal(new[] { "Page 6", "Page 5", "Page 3", "Page 2" }, selectedText.Select(static text => text.Trim()));
        Assert.Equal(4, extracted.Inspect().PageCount);
        Assert.Equal(new[] { "Page 6", "Page 5", "Page 3", "Page 2" }, extracted.Read.TextByPage().Select(static text => text.Trim()));
        Assert.Equal(3, deleted.Inspect().PageCount);
        Assert.True(reordered.Succeeded, string.Join(Environment.NewLine, reordered.Diagnostics));
        Assert.Equal(new[] { "Page 6", "Page 5", "Page 4", "Page 3", "Page 2", "Page 1" }, reordered.RequireValue().Read.TextByPage().Select(static text => text.Trim()));
    }

    private static byte[] CreateSixPageDocument() {
        PdfDocument document = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Page 1"));
        for (int page = 2; page <= 6; page++) {
            document.PageBreak().Paragraph(paragraph => paragraph.Text("Page " + page));
        }

        return document.ToBytes();
    }
}
