using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentOdgListLabelTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void EmptySpansInListTextCannotBypassTheProjectionNodeBudget(bool afterVisibleParagraph) {
        var document = OdgDocument.Create(); var shape = Shape(document);
        if (afterVisibleParagraph) shape.AddParagraph("Before the list");
        var p = shape.AddList().AddItem("").Paragraphs[0];
        p.Element.Add(Enumerable.Range(0, 100001).Select(_ => new XElement(OdfNamespaces.Text + "span")));
        Assert.Throws<NotSupportedException>(() => p.Text);
        var result = document.Pages[0].ToDrawing();
        Assert.DoesNotContain(result.Value.Elements, e => e is OfficeDrawingRichText);
        Assert.Contains(result.Report.Mappings, m => m.Feature == "shape:Lists:text" && m.Status == OdfConversionMappingStatus.Skipped);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ProjectionNodeBudgetIsSharedByAllParagraphsAndFallbackContainersInTheShape(bool metadataContainer) {
        var document = OdgDocument.Create(); var shape = Shape(document); var list = shape.AddList();
        shape.TextRoot.AddFirst(new XElement(OdfNamespaces.Text + "p", "Before the list"));
        for (int i = 0; i < 2; i++) {
            var spans = Enumerable.Range(0, 50001).Select(_ => new XElement(OdfNamespaces.Text + "span"));
            list.AddItem("Body").Paragraphs[0].Element.Add(metadataContainer ?
                new XElement(OdfNamespaces.Text + "meta", spans) : spans);
        }
        var result = document.Pages[0].ToDrawing();
        Assert.DoesNotContain(result.Value.Elements, e => e is OfficeDrawingRichText);
        Assert.Contains(result.Report.Mappings, m => m.Feature == "shape:Lists:text" && m.Status == OdfConversionMappingStatus.Skipped);
    }
    private static OdgShape Shape(OdgDocument document) {
        var shape = document.AddPage().Shapes.AddTextBox(OdfRect.FromCentimeters(1, 1, 14, 15), "", "Lists");
        shape.TextRoot.RemoveNodes(); return shape;
    }
    private static OfficeDrawingRichText Project(OdgDocument document) => Assert.Single(document.Pages[0].ToDrawing().Value.Elements.OfType<OfficeDrawingRichText>());
    private static XElement Style(OdgDocument document, OdfTextList list) => document.GetXml("content.xml").Root!.Element(OdfNamespaces.Office + "automatic-styles")!
        .Elements(OdfNamespaces.Text + "list-style").Single(e => (string?)e.Attribute(OdfNamespaces.Style + "name") == list.StyleName);
    private static string?[] Labels(OdgDocument document) => Project(document).Paragraphs.Select(p => p.Label?.Run.Text).ToArray();

    [Fact]
    public void NumberingRestartsAndNestedListsLabelEachItemOnceWithoutMutatingXml() {
        var document = OdgDocument.Create(); var shape = Shape(document); var list = shape.AddList(true);
        var first = list.AddItem("One"); first.AddParagraph("Continued"); var nested = first.AddList(false); nested.AddItem("Nested");
        var second = list.AddItem("Restart"); second.StartValue = 7; list.AddItem("Next"); list.AddItem("").StartValue = 0;
        string before = document.GetXml("content.xml").ToString();
        Assert.Equal(new[] { "1.", null, "•", "7.", "8.", "0." }, Labels(document));
        Assert.Equal(before, document.GetXml("content.xml").ToString());
        Assert.Equal(36, Project(document).Paragraphs[2].Margins.Left);
        Assert.Throws<ArgumentOutOfRangeException>(() => second.StartValue = -1);
        using var zip = new MemoryStream(); document.Save(zip); zip.Position = 0;
        var reloaded = OdgDocument.Load(zip); Assert.Equal(Labels(document), Labels(reloaded));
        Assert.True(reloaded.Pages[0].Shapes[0].Lists[0].Items[0].Lists[0].Items[0].Paragraphs[0].Text == "Nested");
    }

    [Theory]
    [InlineData("a", 27, false, "aa.")]
    [InlineData("A", 28, false, "AB.")]
    [InlineData("a", 28, true, "bb.")]
    [InlineData("i", 14, false, "xiv.")]
    [InlineData("I", 49, false, "XLIX.")]
    public void NumberFormatsUseNativeCountersAndLetterSynchronization(string format, long start, bool sync, string expected) {
        var document = OdgDocument.Create(); var list = Shape(document).AddList(true); list.AddItem("Item").StartValue = start;
        XElement level = Style(document, list).Elements().Single(); level.SetAttributeValue(OdfNamespaces.Style + "num-format", format);
        level.SetAttributeValue(OdfNamespaces.Style + "num-letter-sync", sync ? "true" : "false");
        Assert.Equal(expected, Labels(document).Single());
    }

    [Fact]
    public void ContinuationByIdTakesPrecedenceOverContinueNumberingAndHeaderDoesNotIncrement() {
        var document = OdgDocument.Create(); var shape = Shape(document); var first = shape.AddList(true); first.AddItem("One"); first.AddItem("Two");
        XElement original = shape.TextRoot.Elements(OdfNamespaces.Text + "list").Single(); original.SetAttributeValue(XNamespace.Xml + "id", "original");
        XElement next = new XElement(OdfNamespaces.Text + "list", new XAttribute(OdfNamespaces.Text + "style-name", first.StyleName!),
            new XAttribute(OdfNamespaces.Text + "continue-list", "original"), new XAttribute(OdfNamespaces.Text + "continue-numbering", "false"),
            new XElement(OdfNamespaces.Text + "list-header", new XElement(OdfNamespaces.Text + "p", "Header")),
            new XElement(OdfNamespaces.Text + "list-item", new XElement(OdfNamespaces.Text + "p", "Three")));
        shape.TextRoot.Add(next);
        var implicitNext = new XElement(next); implicitNext.Attribute(OdfNamespaces.Text + "continue-list")!.Remove();
        implicitNext.SetAttributeValue(OdfNamespaces.Text + "continue-numbering", "true"); implicitNext.Element(OdfNamespaces.Text + "list-header")!.Remove(); shape.TextRoot.Add(implicitNext);
        Assert.Equal(new[] { "1.", "2.", null, "3.", "4." }, Labels(document));
    }

    [Fact]
    public void MultiLevelLabelsUseActualAncestorFormatsAndExplicitItemStyleOverrides() {
        var document = OdgDocument.Create(); var shape = Shape(document); var first = shape.AddList(true); var item = first.AddItem("One");
        var nested = item.AddList(true); nested.AddItem("Child"); XElement level = Style(document, nested).Elements().Single();
        level.SetAttributeValue(OdfNamespaces.Style + "num-format", "a"); level.SetAttributeValue(OdfNamespaces.Text + "display-levels", 2);
        Assert.Equal(new[] { "1.", "1.a." }, Labels(document));
        var alternate = shape.AddList(false); alternate.AddItem("Bullet");
        shape.TextRoot.Elements(OdfNamespaces.Text + "list").First().Elements(OdfNamespaces.Text + "list-item").First().SetAttributeValue(OdfNamespaces.Text + "style-override", alternate.StyleName);
        Assert.Equal(new[] { "•", null, "•" }, Labels(document)); // An unnumbered displayed ancestor is explicitly outside the numbered hierarchy profile.
        Assert.Contains(document.Pages[0].ToDrawing().Report.Mappings, m => m.Feature.EndsWith(":list-numbering", StringComparison.Ordinal));
    }

    [Fact]
    public void LabelFontPercentageAndAlignmentAreIndependentOfBodyRuns() {
        var document = OdgDocument.Create(); var list = Shape(document).AddList(false); var p = list.AddItem("Body").Paragraphs[0]; p.FontSize = OdfLength.Points(20);
        XElement level = Style(document, list).Elements().Single(); level.Add(new XElement(OdfNamespaces.Style + "text-properties", new XAttribute(OdfNamespaces.Fo + "font-size", "50%"),
            new XAttribute(OdfNamespaces.Fo + "font-family", "Liberation Serif"), new XAttribute(OdfNamespaces.Fo + "color", "#FF0000")));
        var text = Project(document); Assert.Equal(10, text.Paragraphs[0].Label!.Run.FontSize); Assert.Equal(20, text.Paragraphs[0].Runs[0].FontSize);
        Assert.Equal(OfficeColor.Red, text.Paragraphs[0].Label!.Run.Color); Assert.Equal("Liberation Serif", text.Paragraphs[0].Label!.Run.FontFamily);
        level.SetAttributeValue(OdfNamespaces.Text + "bullet-relative-size", "40%");
        Assert.Equal(8, Project(document).Paragraphs[0].Label!.Run.FontSize);
        string svg = OfficeDrawingSvgExporter.ToSvg(document.Pages[0].ToDrawing().Value); Assert.Contains("•", svg); Assert.Contains("#FF0000", svg);
    }

    [Fact]
    public void ModernLabelAlignmentUsesParagraphOverridesAndReportsTabStopApproximation() {
        var document = OdgDocument.Create(); var list = Shape(document).AddList(true); var p = list.AddItem("Body").Paragraphs[0]; p.MarginLeft = OdfLength.Points(40);
        XElement properties = Style(document, list).Elements().Single().Element(OdfNamespaces.Style + "list-level-properties")!;
        properties.ReplaceAttributes(new XAttribute(OdfNamespaces.Text + "list-level-position-and-space-mode", "label-alignment"));
        properties.Add(new XElement(OdfNamespaces.Style + "list-level-label-alignment", new XAttribute(OdfNamespaces.Fo + "margin-left", "30pt"),
            new XAttribute(OdfNamespaces.Fo + "text-indent", "-15pt"), new XAttribute(OdfNamespaces.Text + "label-followed-by", "space")));
        var text = Project(document); Assert.Equal(40, text.Paragraphs[0].Margins.Left); Assert.Equal(25, text.Paragraphs[0].Label!.Position);
        Assert.Equal(OfficeTextParagraphLabelFollowedBy.Space, text.Paragraphs[0].Label!.FollowedBy);
        p.MarginLeft = null; p.EnsureStyle().TextIndent = OdfLength.Points(-15);
        Assert.Equal(15, Project(document).Paragraphs[0].Label!.Position); // A list-level margin supplies the base for a paragraph-only hanging override.
        properties.Elements().Single().SetAttributeValue(OdfNamespaces.Text + "label-followed-by", "listtab");
        properties.Elements().Single().SetAttributeValue(OdfNamespaces.Text + "list-tab-stop-position", "40pt");
        Assert.Contains(document.Pages[0].ToDrawing().Report.Mappings, m => m.Feature.EndsWith(":list-tab-stops", StringComparison.Ordinal) && m.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Theory]
    [InlineData("format")]
    [InlineData("continuation")]
    [InlineData("missing-style")]
    [InlineData("consecutive")]
    [InlineData("image")]
    [InlineData("named-format")]
    public void UnsupportedLabelsRetainBodyAndGeometryAndRejectStrictPolicy(string kind) {
        var document = OdgDocument.Create(); var shape = Shape(document); var list = shape.AddList(true); list.AddItem("Retained body");
        XElement native = shape.TextRoot.Elements(OdfNamespaces.Text + "list").Single(), style = Style(document, list), level = style.Elements().Single();
        switch (kind) {
            case "format": level.SetAttributeValue(OdfNamespaces.Style + "num-format", "unsupported-format"); break;
            case "continuation": native.SetAttributeValue(OdfNamespaces.Text + "continue-list", "external-story"); break;
            case "missing-style": native.SetAttributeValue(OdfNamespaces.Text + "style-name", "Missing"); break;
            case "consecutive": style.SetAttributeValue(OdfNamespaces.Text + "consecutive-numbering", "true"); break;
            case "image": level.Name = OdfNamespaces.Text + "list-level-style-image"; break;
            case "named-format": level.SetAttributeValue(OdfNamespaces.Style + "num-list-format-name", "Named"); break;
        }
        var result = document.Pages[0].ToDrawing(); var text = Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>());
        Assert.Equal("Retained body", text.PlainText); Assert.Null(text.Paragraphs[0].Label); Assert.True(result.Report.HasSkippedOrUnsupported);
        Assert.Throws<OdfConversionLossException>(() => document.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
        Assert.Equal("Retained body", shape.Paragraphs[0].Text);
    }

    [Fact]
    public void EmptyListItemStillProducesVisibleBullet() {
        var document = OdgDocument.Create(); Shape(document).AddList().AddItem("");
        Assert.Equal("•", Labels(document).Single()); Assert.Contains("•", OfficeDrawingSvgExporter.ToSvg(document.Pages[0].ToDrawing().Value));
    }

    [Fact]
    public void NothingSeparatorDoesNotSpendACharacterOutsideTheSharedTextBudget() {
        const int maximumCharacters = 100_000;
        var document = OdgDocument.Create(); var list = Shape(document).AddList(true);
        list.AddItem(new string('a', maximumCharacters - 2));
        XElement properties = Style(document, list).Elements().Single().Element(OdfNamespaces.Style + "list-level-properties")!;
        properties.ReplaceAttributes(new XAttribute(OdfNamespaces.Text + "list-level-position-and-space-mode", "label-alignment"));
        properties.Add(new XElement(OdfNamespaces.Style + "list-level-label-alignment",
            new XAttribute(OdfNamespaces.Text + "label-followed-by", "nothing")));
        Assert.Equal(maximumCharacters, Project(document).PlainText.Length);
    }

    [Fact]
    public void IndependentImpressListContentProjectsAfterExplicitDrawingEnvelopeAdaptation() {
        string path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "Drawing", "libreoffice-list-alignment.fodp");
        var xml = XDocument.Load(path);
        // Preserve upstream bytes on disk; only adapt the shared text content's presentation envelope in memory.
        xml.Root!.SetAttributeValue(OdfNamespaces.Office + "mimetype", "application/vnd.oasis.opendocument.graphics");
        xml.Root.Element(OdfNamespaces.Office + "body")!.Element(OdfNamespaces.Office + "presentation")!.Name = OdfNamespaces.Office + "drawing";
        string nativeList = xml.Descendants(OdfNamespaces.Text + "list").Single().ToString();
        using var stream = new MemoryStream(); xml.Save(stream); stream.Position = 0;
        var document = OdgDocument.LoadFlatXml(stream); var result = document.Pages[0].ToDrawing();
        var text = Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>(), candidate => candidate.Paragraphs.Any(paragraph => paragraph.Label != null));
        Assert.Equal(4, text.Paragraphs.Count); Assert.All(text.Paragraphs, p => Assert.Equal("●", p.Label!.Run.Text));
        Assert.Equal(8.503937007874015, text.Paragraphs[0].Label!.Position, 8);
        Assert.Equal(25.51181102362205, text.Paragraphs[0].Label!.MinimumWidth!.Value, 8);
        Assert.Equal("OpenSymbol", text.Paragraphs[0].Label!.Run.FontFamily);
        Assert.Equal(10.8, text.Paragraphs[0].Label!.Run.FontSize, 8); // 45% of the fixture's 24pt graphic default; presentation-style defaults remain outside this adapter.
        Assert.Contains(result.Report.Mappings, m => m.Feature.EndsWith(":writing-mode", StringComparison.Ordinal));
        Assert.True(XNode.DeepEquals(XElement.Parse(nativeList), XElement.Parse(document.GetXml("content.xml").Descendants(OdfNamespaces.Text + "list").Single().ToString())));
    }
}
