using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfLogicalListNumberingTests {
    [Fact]
    public void Read_NestsIndentedItemsAndClassifiesTheirMarkers() {
        var flat = new PdfListStyle { SpacingBefore = global::OfficeIMO.OfficeLength.Zero };
        var nested = new PdfListStyle { LeftIndent = global::OfficeIMO.OfficeLength.Points(18), SpacingBefore = global::OfficeIMO.OfficeLength.Zero };
        var document = PdfDocument.Create();
        document.AddNumberedList(new[] { "Plan the release" }, style: flat);
        document.AddList(new[] { new PdfListItem("Draft notes", marker: "(a)"), new PdfListItem("Review notes", marker: "(b)") }, style: nested);
        document.AddNumberedList(new[] { "Ship it" }, startNumber: 2, style: flat);

        PdfLogicalPage page = Assert.Single(PdfDocumentReadResult.Load(document.ToBytes(), new PdfTextLayoutOptions { ForceSingleColumn = true }).Pages);

        PdfLogicalListItem Item(string text) => Assert.Single(page.ListItems, item => item.Text == text);
        Assert.Equal((1, PdfLogicalListNumberFormat.Number, (int?)1), Describe(Item("Plan the release")));
        Assert.Equal((2, PdfLogicalListNumberFormat.LowerLetter, (int?)1), Describe(Item("Draft notes")));
        Assert.Equal((2, PdfLogicalListNumberFormat.LowerLetter, (int?)2), Describe(Item("Review notes")));
        Assert.Equal((1, PdfLogicalListNumberFormat.Number, (int?)2), Describe(Item("Ship it")));
        Assert.True(Item("Draft notes").IsOrdered);
    }

    private static (int Level, PdfLogicalListNumberFormat Format, int? Ordinal) Describe(PdfLogicalListItem item) =>
        (item.Level, item.NumberFormat, item.Ordinal);

    [Theory]
    [InlineData("(iv)", "(v)", PdfLogicalListNumberFormat.LowerRoman, 5)]
    [InlineData("(h)", "(i)", PdfLogicalListNumberFormat.LowerLetter, 9)]
    [InlineData("C.", "D.", PdfLogicalListNumberFormat.UpperLetter, 4)]
    public void ListMarkers_ReadSingleCharactersInTheirListContext(string previous, string marker, PdfLogicalListNumberFormat format, int ordinal) {
        PdfLogicalListMarker resolved = PdfLogicalListMarker.Parse(marker).ResolveAfter(PdfLogicalListMarker.Parse(previous), marker);

        Assert.Equal(format, resolved.Format);
        Assert.Equal(ordinal, resolved.Ordinal);
    }
}
