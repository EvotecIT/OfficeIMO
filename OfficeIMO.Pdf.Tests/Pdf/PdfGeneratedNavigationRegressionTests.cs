using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfGeneratedNavigationRegressionTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ListItemBookmarksTargetTheItemIncludingItsBody(bool columns) {
        var items = new[] { new PdfListItem("First item", bookmarkName: "item") };
        var document = PdfDocument.Create(TaggedOptions());
        if (columns) document.Compose(c => c.Page(p => p.Content(content => content.Row(row =>
            row.PercentColumn(100, column => column.RichNumbered(items))))));
        else document.RichNumbered(items);
        document.Paragraph(p => p.LinkToBookmark("Read item", "item"));

        var raw = PdfReadDocument.Open(document.ToBytes()).RawStructure();
        var link = Assert.Single(raw.Objects, obj => obj.Value.Entries.TryGetValue("Subtype", out var type) && type.Text == "Link");
        int target = link.Value.Entries["A"].Entries["SD"].Items[0].ReferenceObjectNumber!.Value;
        Assert.Equal("LI", raw.GetObject(target)!.Value.Entries["S"].Text);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void ForwardOnlyNamedLinksCompleteWithOrWithoutAStructureTarget(bool taggedTarget, bool strict) {
        var options = strict ? StrictOptions() : TaggedOptions();
        options.ObjectSerializationMode = PdfObjectSerializationMode.ForwardOnly;
        byte[] bytes = PdfDocument.Create(options).Meta(title: "Forward-only navigation").Canvas(canvas => {
            if (taggedTarget) canvas.Structure(PdfCanvasStructureRole.Section, section => section
                .NamedDestination("target", 20, 20).Text("Target", 20, 20, 100, 20));
            else canvas.NamedDestination("target", 20, 20).Text("Target", 20, 20, 100, 20);
            canvas.LinkToNamedDestination("target", 20, 60, 100, 20, "Read target");
        }).ToBytes();

        var read = PdfReadDocument.Open(bytes);
        Assert.Single(read.NamedDestinations);
        var link = Assert.Single(read.RawStructure().Objects, obj => obj.Value.Entries.TryGetValue("Subtype", out var type) && type.Text == "Link");
        Assert.Equal("GoTo", link.Value.Entries["A"].Entries["S"].Text);
        Assert.Equal(taggedTarget, link.Value.Entries["A"].Entries.ContainsKey("SD"));
    }

    [Theory]
    [InlineData("image")]
    [InlineData("text")]
    [InlineData("radio")]
    public void FullyClippedDeferredObjectsDoNotClaimTheNextVisibleBookmarkTarget(string kind) {
        byte[] image = PdfPngTestImages.CreateRgbPng(2, 2);
        var style = new PdfFormFieldStyle { AlternateName = "Hidden control" };
        var document = PdfDocument.Create(StrictOptions()).Meta(title: "Clipped target").Bookmark("target").Canvas(canvas => canvas
            .Clip(20, 20, 20, 20, clipped => {
                if (kind == "image") clipped.Image(image, 100, 100, 20, 20, alternativeText: "Hidden image");
                else if (kind == "radio") clipped.RadioButton("hidden", "Yes", true, 100, 100, 20, 20, style: style);
                else clipped.TextField("hidden", "", 100, 100, 20, 20, style: style);
            }))
            .Paragraph(p => p.Text("Visible target ").LinkToBookmark("Return", "target"));

        var raw = PdfReadDocument.Open(document.ToBytes()).RawStructure();
        var link = Assert.Single(raw.Objects, obj => obj.Value.Entries.TryGetValue("Subtype", out var type) && type.Text == "Link");
        int target = link.Value.Entries["A"].Entries["SD"].Items[0].ReferenceObjectNumber!.Value;
        Assert.Equal("P", raw.GetObject(target)!.Value.Entries["S"].Text);
    }

    [Theory]
    [InlineData("image", "text", "P")]
    [InlineData("field", "text", "P")]
    [InlineData("radio", "text", "P")]
    [InlineData("image", "image", "Figure")]
    [InlineData("field", "field", "Form")]
    [InlineData("mixed", "text", "P")]
    public void ClippedTargetsFallForwardToTheirFirstSurvivingSiblingInsideTheClip(string hidden, string visible, string expectedRole) {
        byte[] image = PdfPngTestImages.CreateRgbPng(2, 2);
        var style = new PdfFormFieldStyle { AlternateName = "Named control" };
        byte[] bytes = PdfDocument.Create(StrictOptions()).Meta(title: "Clipped sibling target").Bookmark("target")
            .Canvas(canvas => canvas.Clip(20, 20, 80, 80, clipped => {
                if (hidden == "image" || hidden == "mixed") clipped.Image(image, 120, 120, 20, 20, alternativeText: "Clipped figure");
                if (hidden == "field" || hidden == "mixed") clipped.TextField("hidden", "", 120, 150, 20, 20, style: style);
                if (hidden == "radio") clipped.RadioButton("hidden", "Yes", true, 120, 120, 20, 20, style: style);
                if (visible == "text") clipped.Text("Visible sibling", 25, 25, 60, 20);
                else if (visible == "image") clipped.Image(image, 25, 25, 20, 20, alternativeText: "Visible sibling");
                else clipped.CheckBox("visible", false, 25, 25, 20, 20, style: style);
            }))
            .Paragraph(p => p.Text("Later outside paragraph ").LinkToBookmark("Return", "target")).ToBytes();

        var raw = PdfReadDocument.Open(bytes).RawStructure();
        var link = Assert.Single(raw.Objects, obj => obj.Value.Entries.TryGetValue("Subtype", out var type) && type.Text == "Link");
        int target = link.Value.Entries["A"].Entries["SD"].Items[0].ReferenceObjectNumber!.Value;
        var firstSurvivingSibling = raw.Objects.First(obj => obj.Value.Entries.TryGetValue("Type", out var type) && type.Text == "StructElem"
            && obj.Value.Entries["S"].Text == expectedRole);
        Assert.Equal(firstSurvivingSibling.ObjectNumber, target);
    }

    private static PdfOptions StrictOptions() {
        string? font = PdfComplianceTestFonts.FindBundledTrueTypeFont();
        Assert.NotNull(font);
        return new PdfOptions().UseFontFamily("Carlito", font!).UsePdfUa(PdfComplianceProfile.PdfUa2, "en");
    }

    private static PdfOptions TaggedOptions() => new PdfOptions().ConfigurePdfUaGroundwork(PdfComplianceProfile.PdfUa2, "en");
}
