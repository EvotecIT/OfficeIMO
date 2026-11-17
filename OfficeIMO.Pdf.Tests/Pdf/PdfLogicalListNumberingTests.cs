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

    [Fact]
    public void Read_KeepsFlatListsFlatAcrossRightAlignedMarkersAndInterveningParagraphs() {
        var flat = new PdfListStyle { SpacingBefore = global::OfficeIMO.OfficeLength.Zero, SpacingAfter = global::OfficeIMO.OfficeLength.Zero };
        var indented = new PdfListStyle { LeftIndent = global::OfficeIMO.OfficeLength.Points(18), SpacingBefore = global::OfficeIMO.OfficeLength.Zero, SpacingAfter = global::OfficeIMO.OfficeLength.Zero };
        var numbers = new PdfListStyle {
            SpacingBefore = global::OfficeIMO.OfficeLength.Zero,
            MarkerAlign = OfficeAlignment.Right,
            MarkerWidth = global::OfficeIMO.OfficeLength.Points(36)
        };
        var document = PdfDocument.Create();
        // Right-aligned labels: "100." starts two digits left of "1." while both item texts line up.
        document.AddList(new[] { new PdfListItem("Scope", marker: "1."), new PdfListItem("Budget", marker: "2."), new PdfListItem("Review", marker: "100.") }, style: numbers);
        // A tight paragraph keeps the lists within one line spacing of each other; only the text between separates them.
        var tight = new PdfParagraphStyle { SpacingBefore = global::OfficeIMO.OfficeLength.Zero, SpacingAfter = global::OfficeIMO.OfficeLength.Zero };
        document.AddParagraph("Separate notes follow.", style: tight);
        document.AddList(new[] { "Reduce backlog" }, style: flat);
        document.AddParagraph("Then, on its own:", style: tight);
        document.AddList(new[] { "Archive exports" }, style: indented);

        PdfLogicalPage page = Assert.Single(PdfDocumentReadResult.Load(document.ToBytes(), new PdfTextLayoutOptions { ForceSingleColumn = true }).Pages);

        Assert.All(new[] { "Scope", "Review", "Reduce backlog", "Archive exports" },
            text => Assert.Equal(1, Assert.Single(page.ListItems, item => item.Text == text).Level));
    }

    [Fact]
    public void Read_ResolvesAnAmbiguousFirstMarkerFromTheItemAfterIt() {
        var document = PdfDocument.Create();
        document.AddList(new[] { new PdfListItem("Ninth", marker: "(i)"), new PdfListItem("Tenth", marker: "(j)") });

        PdfLogicalPage page = Assert.Single(PdfDocumentReadResult.Load(document.ToBytes(), new PdfTextLayoutOptions { ForceSingleColumn = true }).Pages);

        Assert.Equal((1, PdfLogicalListNumberFormat.LowerLetter, (int?)9), Describe(Assert.Single(page.ListItems, item => item.Text == "Ninth")));
    }

    [Fact]
    public void ListMarkers_ReadDecimalDigitsOutsideTheBasicPlane() {
        PdfLogicalListMarker marker = PdfLogicalListMarker.Parse("\U0001D7DA.");

        Assert.Equal(PdfLogicalListNumberFormat.Number, marker.Format);
        Assert.Equal(2, marker.Ordinal);
    }

    private static (int Level, PdfLogicalListNumberFormat Format, int? Ordinal) Describe(PdfLogicalListItem item) =>
        (item.Level, item.NumberFormat, item.Ordinal);

    [Theory]
    [InlineData("(iv)", "(v)", PdfLogicalListNumberFormat.LowerRoman, 5)]
    [InlineData("(h)", "(i)", PdfLogicalListNumberFormat.LowerLetter, 9)]
    [InlineData("C.", "D.", PdfLogicalListNumberFormat.UpperLetter, 4)]
    public void ListMarkers_ReadSingleCharactersInTheirListContext(string previous, string marker, PdfLogicalListNumberFormat format, int ordinal) {
        PdfLogicalListMarker resolved = PdfLogicalListMarker.Parse(marker).ResolveBeside(PdfLogicalListMarker.Parse(previous), marker);

        Assert.Equal(format, resolved.Format);
        Assert.Equal(ordinal, resolved.Ordinal);
    }
}
