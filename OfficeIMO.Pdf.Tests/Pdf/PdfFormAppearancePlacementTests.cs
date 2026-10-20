using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfFormAppearancePlacementTests {
    [Theory]
    [InlineData("text", "authored")]
    [InlineData("comb", "authored")]
    [InlineData("choice", "authored")]
    [InlineData("text", "filled")]
    [InlineData("comb", "filled")]
    [InlineData("choice", "filled")]
    [InlineData("text", "appended")]
    [InlineData("comb", "appended")]
    [InlineData("choice", "appended")]
    public void SavedSingleLineAppearancesKeepAscendersAndDescendersInsideWidget(string kind, string operation) {
        PdfDocument source = PdfDocument.Load(PdfFormPlacementFixture.CreateAuthored(kind));
        var values = new Dictionary<string, string> { ["Value"] = "QgjH" };
        PdfDocument result = operation switch {
            "filled" => source.Forms.Fill(values),
            "appended" => source.Forms.AppendRevision(values, keepNeedAppearances: false),
            _ => source
        };
        PdfFormField field = Assert.Single(result.Forms.GetForInteraction().Fields);
        PdfFormWidget widget = Assert.Single(field.Widgets);
        Assert.Equal(operation == "authored" ? "HgjQ" : "QgjH", field.Value);

        using PdfPigDocument flattened = PdfPigDocument.Open(result.Forms.Flatten().ToBytes());
        var letters = flattened.GetPage(1).Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();
        Assert.Equal(4, letters.Length);
        Assert.All(letters, letter => {
            Assert.InRange(letter.BoundingBox.Bottom, widget.Y1, widget.Y2);
            Assert.InRange(letter.BoundingBox.Top, widget.Y1, widget.Y2);
        });
    }

    [Theory]
    [InlineData(null, null, 2, false)]
    [InlineData(null, 2, null, false)]
    [InlineData(2, null, null, false)]
    [InlineData(1, 2, null, true)]
    [InlineData(1, 0, 2, true)]
    public void FlattenMissingAppearanceRetainsEffectiveRightAlignment(int? catalogAlignment, int? parentAlignment, int? widgetAlignment, bool selective) {
        PdfDocument source = PdfDocument.Load(PdfFormPlacementFixture.CreateWithoutAppearance(catalogAlignment, parentAlignment, widgetAlignment));
        PdfFormWidget widget = Assert.Single(Assert.Single(source.Forms.GetForInteraction().Fields).Widgets);
        PdfDocument result = selective ? source.Forms.Flatten("Value") : source.Forms.Flatten();

        using PdfPigDocument flattened = PdfPigDocument.Open(result.ToBytes());
        var letters = flattened.GetPage(1).Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();
        Assert.Equal("HgjQ", string.Concat(letters.Select(letter => letter.Value)));
        // Text must remain against the right inset, rather than silently reverting to the left.
        Assert.InRange(letters.Min(letter => letter.StartBaseLine.X), widget.X2 - 40D, widget.X2 - 20D);
        Assert.InRange(letters.Max(letter => letter.BoundingBox.Right), widget.X2 - 5D, widget.X2 - 2D);
    }

    [Fact]
    public void SelectiveFlattenRetainsAlignmentFromUnselectedAncestor() {
        PdfDocument source = PdfDocument.Load(PdfFormPlacementFixture.CreateWithoutAppearance(1, 2, null, nested: true));
        PdfFormField field = Assert.Single(source.Forms.GetForInteraction().Fields);
        Assert.Equal("Group.Value", field.Name);
        PdfFormWidget widget = Assert.Single(field.Widgets);

        using PdfPigDocument flattened = PdfPigDocument.Open(source.Forms.Flatten("Group.Value").ToBytes());
        var letters = flattened.GetPage(1).Letters.ToArray();
        Assert.Equal("HgjQ", string.Concat(letters.Select(letter => letter.Value)));
        Assert.InRange(letters.Min(letter => letter.StartBaseLine.X), widget.X2 - 40D, widget.X2 - 20D);
    }
}
