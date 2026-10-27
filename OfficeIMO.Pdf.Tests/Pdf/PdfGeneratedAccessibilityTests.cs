using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfGeneratedAccessibilityTests {
    private static PdfCreateOptions TaggedOptions() => new PdfCreateOptions().ConfigurePdfUaGroundwork(PdfComplianceProfile.PdfUa2, "en");

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NativeListsDeclareTheirLabelStyleInFlowAndColumns(bool columns) {
        PdfDocument document = PdfDocument.Create(TaggedOptions());
        if (columns) document.Compose(compose => compose.Page(page => page.Content(content => content.Row(row =>
            row.PercentColumn(100, column => column.Numbered(new[] { "First", "Second" }))))));
        else document.Numbered(new[] { "First", "Second" });

        PdfRawDocumentView raw = PdfReadDocument.Open(document.ToBytes()).RawStructure();
        PdfRawObjectView list = Assert.Single(raw.Objects, item => IsStructure(item, "L"));

        Assert.Equal("Decimal", list.Value.Entries["A"].Entries["ListNumbering"].Text);
    }

    [Fact]
    public void CanvasListLabelsContainInlineTextAndRetainNumberingOptions() {
        var options = new PdfCanvasStructureOptions { ListNumbering = PdfListNumbering.UpperRoman };
        PdfDocument document = PdfDocument.Create(TaggedOptions()).Canvas(canvas =>
            canvas.Structure(PdfCanvasStructureRole.List, list => list.Structure(PdfCanvasStructureRole.ListItem, item => item
                .Structure(PdfCanvasStructureRole.ListLabel, label => label.Text("I.", 20, 20, 20, 20))
                .Structure(PdfCanvasStructureRole.ListBody, body => body.Text("First", 45, 20, 100, 20))), options));
        options.ListNumbering = PdfListNumbering.None;

        PdfReadDocument read = PdfReadDocument.Open(document.ToBytes());
        PdfRawDocumentView raw = read.RawStructure();
        PdfRawObjectView list = Assert.Single(raw.Objects, item => IsStructure(item, "L"));
        Assert.Equal("UpperRoman", list.Value.Entries["A"].Entries["ListNumbering"].Text);
        var label = Assert.Single(read.TaggedContent!.StructureElements, element => element.StructureType == "Lbl");
        Assert.All(label.ChildElementObjectNumbers, child =>
            Assert.Equal("Span", read.TaggedContent.StructureElements.Single(element => element.ObjectNumber == child).StructureType));
    }

    [Fact]
    public void CanvasNavigationUsesItsActualContainingStructureAndPreservesPageDestinations() {
        byte[] bytes = PdfDocument.Create(TaggedOptions()).Canvas(canvas => canvas
            .Structure(PdfCanvasStructureRole.Section, section => section
                .NamedDestination("details", 20, 20).Outline("Details", 1, 20)
                .Text("Details", 20, 20, 100, 20))
            .LinkToNamedDestination("details", 20, 60, 100, 20, "Details")).ToBytes();

        PdfReadDocument read = PdfReadDocument.Open(bytes);
        PdfRawDocumentView raw = read.RawStructure();
        int sectionId = Assert.Single(raw.Objects, item => IsStructure(item, "Sect")).ObjectNumber;
        var link = Assert.Single(raw.Objects, item => item.Value.Entries.TryGetValue("Subtype", out var subtype) && subtype.Text == "Link");
        Assert.Equal(sectionId, link.Value.Entries["A"].Entries["SD"].Items[0].ReferenceObjectNumber);
        var outline = Assert.Single(raw.Objects, item => item.Value.Entries.ContainsKey("Title"));
        Assert.Equal(sectionId, outline.Value.Entries["A"].Entries["SD"].Items[0].ReferenceObjectNumber);
        Assert.Equal(1, Assert.Single(read.Outlines).PageNumber);
    }

    [Fact]
    public void NativeHeadingOutlinesRetainTheirHeadingStructure() {
        PdfCreateOptions options = TaggedOptions();
        options.CreateOutlineFromHeadings = true;
        byte[] bytes = PdfDocument.Create(options).H1("First heading").H2("Second heading").ToBytes();
        PdfRawDocumentView raw = PdfReadDocument.Open(bytes).RawStructure();
        var outlines = raw.Objects.Where(item => item.Value.Entries.ContainsKey("Title")).ToArray();
        Assert.Equal(2, outlines.Length);
        Assert.Equal(new[] { "H1", "H2" }, outlines.Select(item => {
            int id = item.Value.Entries["A"].Entries["SD"].Items[0].ReferenceObjectNumber!.Value;
            return raw.GetObject(id)!.Value.Entries["S"].Text;
        }));
    }

    [Fact]
    public void WidgetDescriptionsCoverCombinedFieldsAndSeparateRadioWidgets() {
        var style = new PdfFormFieldStyle { AlternateName = "Accessible control" };
        byte[] bytes = PdfDocument.Create(TaggedOptions()).Canvas(canvas => canvas
            .TextField("name", "", 20, 20, 100, 20, style: style)
            .CheckBox("reviewed", false, 20, 50, 20, 20, style: style)
            .ChoiceField("region", new[] { "North", "South" }, new[] { "North" }, 20, 80, 100, 20, style: style)
            .RadioButton("delivery", "Email", true, 20, 110, 20, 20, style: style)).ToBytes();
        var widgets = PdfReadDocument.Open(bytes).RawStructure().Objects.Where(item =>
            item.Value.Entries.TryGetValue("Subtype", out var subtype) && subtype.Text == "Widget").ToArray();
        Assert.Equal(4, widgets.Length);
        Assert.All(widgets, widget => Assert.Equal("Accessible control", widget.Value.Entries["Contents"].Text));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void StrictUa2RejectsCanvasNavigationWithoutATaggedTarget(bool outline) {
        PdfDocument document = PdfDocument.Create(StrictOptions()).Meta(title: "Navigation boundary")
            .Canvas(canvas => {
                canvas.Text("Visible content", 20, 20, 120, 20);
                if (outline) canvas.Outline("Details", 1, 20);
                else canvas.NamedDestination("details", 20, 20)
                    .LinkToNamedDestination("details", 20, 60, 100, 20, "Details");
            });

        var error = Assert.Throws<InvalidOperationException>(() => document.ToBytes());
        Assert.Contains(outline ? "outlines require a tagged target" : "internal links require a named destination inside tagged content", error.Message);
    }

    [Fact]
    public void Ua2OpeningActionPreservesItsPageAndAddsALogicalTarget() {
        PdfCreateOptions options = StrictOptions();
        options.SetOpenAction(pageNumber: 2, destinationTop: 100);
        byte[] bytes = PdfDocument.Create(options).Meta(title: "Opening action")
            .Paragraph(p => p.Text("First page")).PageBreak()
            .Paragraph(p => p.Text("Second page")).ToBytes();

        PdfRawDocumentView raw = PdfReadDocument.Open(bytes).RawStructure();
        var action = raw.GetObject(raw.CatalogObjectNumber!.Value)!.Value.Entries["OpenAction"];
        var pageTarget = raw.GetObject(action.Entries["D"].Items[0].ReferenceObjectNumber!.Value)!;
        var structureTarget = raw.GetObject(action.Entries["SD"].Items[0].ReferenceObjectNumber!.Value)!;
        Assert.Equal(pageTarget.ObjectNumber, structureTarget.Value.Entries["Pg"].ReferenceObjectNumber);
        Assert.Equal("GoTo", action.Entries["S"].Text);
        Assert.Equal(100D, action.Entries["D"].Items[3].Number);
    }

    private static PdfCreateOptions StrictOptions() {
        string? font = PdfComplianceTestFonts.FindBundledTrueTypeFont();
        Assert.NotNull(font);
        return new PdfCreateOptions().UseFontFamily("Carlito", font!).UsePdfUa(PdfComplianceProfile.PdfUa2, "en");
    }

    [Fact]
    public void FloatingFlowBookmarkKeepsTheFollowingParagraphAsItsLogicalTarget() {
        PdfCreateOptions options = StrictOptions();
        options.PageWidth = 400;
        options.PageHeight = 240;
        options.MarginLeft = options.MarginRight = options.MarginTop = options.MarginBottom = 40;
        var style = new PdfTableStyle {
            HeaderRowCount = 0, ColumnWidthPoints = new List<double?> { 320 }, MinRowHeight = 150,
            Position = new PdfTablePosition()
        };
        byte[] bytes = PdfDocument.Create(options).Meta(title: "Floating target")
            .Table(new[] { new[] { "Floating table" } }, style: style)
            .Bookmark("following")
            .Paragraph(p => p.Text("Following paragraph").LinkToBookmark("Read again", "following"))
            .ToBytes();

        PdfReadDocument read = PdfReadDocument.Open(bytes);
        Assert.Equal(2, Assert.Single(read.NamedDestinations).PageNumber);
        PdfRawDocumentView raw = read.RawStructure();
        var link = Assert.Single(raw.Objects, item => item.Value.Entries.TryGetValue("Subtype", out var subtype) && subtype.Text == "Link");
        int targetId = link.Value.Entries["A"].Entries["SD"].Items[0].ReferenceObjectNumber!.Value;
        Assert.Equal("P", raw.GetObject(targetId)!.Value.Entries["S"].Text);
    }

    private static bool IsStructure(PdfRawObjectView item, string type) =>
        item.Value.Entries.TryGetValue("Type", out var kind) && kind.Text == "StructElem"
        && item.Value.Entries.TryGetValue("S", out var role) && role.Text == type;

    [Theory]
    [InlineData("flow-image", "Figure")]
    [InlineData("column-image", "Figure")]
    [InlineData("canvas-image", "Figure")]
    [InlineData("text-field", "Form")]
    [InlineData("checkbox", "Form")]
    [InlineData("choice", "Form")]
    [InlineData("radio", "Form")]
    [InlineData("column-field", "Form")]
    public void NativeBookmarkTargetsTheFollowingNonTextContentInsteadOfLaterText(string surface, string expectedRole) {
        byte[] image = PdfPngTestImages.CreateRgbPng(2, 2);
        var fieldStyle = new PdfFormFieldStyle { AlternateName = "Contact control" };
        PdfDocument document = PdfDocument.Create(StrictOptions()).Meta(title: "Non-text destination");
        if (surface == "canvas-image") {
            document.Bookmark("target").Canvas(canvas => canvas.Image(image, 20, 20, 20, 20, alternativeText: "Status badge"));
            document.Paragraph(p => p.Text("Later text ").LinkToBookmark("Return to target", "target"));
        } else document.Compose(compose => compose.Page(page => page.Content(content => {
            content.Bookmark("target");
            if (surface == "column-image") content.Row(row => row.PercentColumn(100, column => column.Image(image, 20, 20, alternativeText: "Status badge")));
            else if (surface == "column-field") content.Row(row => row.PercentColumn(100, column => column.TextField("contact", style: fieldStyle)));
            else if (surface == "flow-image") content.Image(image, 20, 20, "Status badge");
            else if (surface == "text-field") content.TextField("contact", style: fieldStyle);
            else if (surface == "checkbox") content.CheckBox("contact", style: fieldStyle);
            else if (surface == "choice") content.ChoiceField("contact", new[] { "Email", "Phone" }, style: fieldStyle);
            else content.RadioButtonGroup("contact", new[] { "Email", "Phone" }, style: fieldStyle);
            content.Paragraph(p => p.Text("Later text ").LinkToBookmark("Return to target", "target"));
        })));

        PdfRawDocumentView raw = PdfReadDocument.Open(document.ToBytes()).RawStructure();
        var link = Assert.Single(raw.Objects, item => item.Value.Entries.TryGetValue("Subtype", out var subtype) && subtype.Text == "Link");
        int target = link.Value.Entries["A"].Entries["SD"].Items[0].ReferenceObjectNumber!.Value;
        Assert.Equal(expectedRole, raw.GetObject(target)!.Value.Entries["S"].Text);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(90)]
    public void CanvasCellDestinationsTargetTheirOwnCellIncludingTheFinalCell(int rotation) {
        var cells = new[] { new[] {
            new PdfTableCell(new[] { PdfTextRun.Normal("First cell") }, paragraphs: null, textRotation: rotation).WithNamedDestination("first"),
            new PdfTableCell(new[] { PdfTextRun.Normal("Last cell") }, paragraphs: null, textRotation: rotation).WithNamedDestination("last")
        } };
        PdfCreateOptions options = StrictOptions();
        byte[] bytes = PdfDocument.Create(options).Meta(title: "Cell destinations").Canvas(canvas => canvas
            .Table(cells, 20, 20, 240, 80, new PdfTableStyle { HeaderRowCount = 0 })
            .LinkToNamedDestination("first", 20, 120, 100, 20, "First cell")
            .LinkToNamedDestination("last", 140, 120, 100, 20, "Last cell")).ToBytes();

        PdfRawDocumentView raw = PdfReadDocument.Open(bytes).RawStructure();
        var cellIds = raw.Objects.Where(item => IsStructure(item, "TD")).Select(item => item.ObjectNumber).ToArray();
        var targetIds = raw.Objects.Where(item => item.Value.Entries.TryGetValue("Subtype", out var subtype) && subtype.Text == "Link")
            .Select(item => item.Value.Entries["A"].Entries["SD"].Items[0].ReferenceObjectNumber!.Value).ToArray();
        Assert.Equal(2, cellIds.Length);
        Assert.Equal(cellIds, targetIds);
    }
}
