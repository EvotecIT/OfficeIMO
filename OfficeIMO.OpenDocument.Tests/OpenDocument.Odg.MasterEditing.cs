using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentOdgMasterEditingTests {
    [Fact]
    public void SharedArtworkEditsPersistAndClonedMasterStylesDetachWithoutChangingPageContent() {
        var document = OdgDocument.Create(); var page = document.AddPage("Shared", P(240), P(180));
        var shape = page.MasterShapes.AddRectangle(R(10, 10, 30, 20), "Artwork");
        shape.FillColor = OdfColor.Parse("#2040e0"); shape.Layer = "layout";
        page.Shapes.AddRectangle(R(80, 10, 30, 20), "Local").FillColor = OdfColor.Parse("#e07020");
        var other = document.ClonePage(0, "Other");
        document = Reopen(document); page = document.Pages[0]; other = document.Pages[1];
        byte[] content = document.GetPackageEntryBytes("content.xml")!;
        other.MasterShapes[0].Bounds = R(20, 20, 40, 30);
        other.MasterShapes[0].FillColor = OdfColor.Parse("#00ff00");
        Assert.Equal(content, document.GetPackageEntryBytes("content.xml"));
        Assert.Equal(other.MasterShapes[0].Bounds, page.MasterShapes[0].Bounds);
        other.MasterPageName = document.CloneMasterPage(page.MasterPageName, "Independent");
        other.MasterShapes[0].FillColor = OdfColor.Parse("#ff0000"); other.MasterShapes[0].FillOpacity = .5;
        foreach (var read in RoundTrips(document)) {
            Assert.Equal(OdfColor.Parse("#00ff00"), read.Pages[0].MasterShapes[0].FillColor);
            Assert.Equal(OdfColor.Parse("#ff0000"), read.Pages[1].MasterShapes[0].FillColor);
            Assert.Equal(.5, read.Pages[1].MasterShapes[0].FillOpacity);
            Assert.Equal("layout", read.Pages[1].MasterShapes[0].Layer);
            Assert.Equal(20, read.Pages[0].MasterShapes[0].Bounds.X.ToPoints());
            Assert.Equal(OdfColor.Parse("#e07020"), read.Pages[0].Shapes[0].FillColor);
            Assert.Equal(OdfColor.Parse("#e07020"), read.Pages[1].Shapes[0].FillColor);
        }
    }

    [Fact]
    public void MasterCollectionsSupportGroupsGeometryImagesOrderAndConnectorDetachment() {
        var document = OdgDocument.Create(); var page = document.AddPage("Master", P(240), P(180));
        var left = page.MasterShapes.AddRectangle(R(10, 10, 20, 20), "Left");
        var right = page.MasterShapes.AddEllipse(R(100, 10, 20, 20), "Right");
        var connector = page.MasterShapes.AddConnector(left.AddGluePoint(OdgGluePointAlignment.Right),
            right.AddGluePoint(OdgGluePointAlignment.Left), "Attached");
        var local = page.Shapes.AddRectangle(R(150, 10, 20, 20));
        Assert.Throws<ArgumentException>(() => connector.AttachEndToShape(local));
        Assert.Throws<ArgumentException>(() => page.Shapes.AddConnector(left.GluePoints[0], right.GluePoints[0]));
        var group = page.MasterShapes.AddGroup("Group");
        group.Children.AddPath(R(10, 60, 30, 20), new OdfViewBox(0, 0, 30, 20), "M0 0 L30 0 L15 20 Z", "Path");
        group.Children.AddImage(Image(), "master.png", R(50, 60, 20, 20), "Image");
        group.TransformChildren("translate(5pt 10pt)");
        document = Reopen(document); page = document.Pages[0];
        page.MasterShapes.Move(3, 0);
        page.MasterShapes.RemoveAt(1); // Remove Left, retaining the connector's current start coordinate.
        foreach (var read in RoundTrips(document)) {
            Assert.Equal(new[] { "Group", "Right", "Attached" }, read.Pages[0].MasterShapes.Select(s => s.Name));
            var saved = read.Pages[0].MasterShapes[2]; Assert.Null(saved.StartShapeId); Assert.NotNull(saved.EndShapeId);
            Assert.Equal(30, saved.X1.ToPoints(), 6); Assert.Equal(20, saved.Y1.ToPoints(), 6);
            Assert.Equal(100, saved.X2.ToPoints(), 6); Assert.Equal(20, saved.Y2.ToPoints(), 6);
            var children = read.Pages[0].MasterShapes[0].Children;
            Assert.Equal("M0 0 L30 0 L15 20 Z", children[0].PathData);
            Assert.Equal(Image(), children[1].GetImageBytes());
            Assert.False(read.Pages[0].ToDrawing().Report.HasSkippedOrUnsupported);
        }
    }

    [Fact]
    public void MasterRichTextAndNestedListsUseOwningStylesAndIsolateClonedFormatting() {
        var document = OdgDocument.Create(); var page = document.AddPage("Text");
        var shape = page.MasterShapes.AddTextBox(R(10, 10, 160, 120), "Base", "Text");
        page.Shapes.AddTextBox(R(10, 150, 160, 40), "Local").AddList().AddItem("Local bullet");
        var paragraph = shape.Paragraphs[0]; paragraph.FontSize = P(12); paragraph.TextAlign = "center";
        paragraph.SetTabStops(new[] { new OdfTabStop(P(30)) });
        var run = paragraph.AddRun("Styled"); run.Bold = true; run.Color = OdfColor.Parse("#2040e0");
        var link = paragraph.AddHyperlink("Link", "https://example.invalid/master"); link.Italic = true;
        var list = shape.AddList(true); var item = list.AddItem("First"); item.StartValue = 3;
        item.AddParagraph("Continuation"); item.AddList().AddItem("Nested");
        var copy = document.ClonePage(0); copy.MasterPageName = document.CloneMasterPage(page.MasterPageName);
        document = Reopen(document); page = document.Pages[0]; copy = document.Pages[1];
        var edited = copy.MasterShapes[0].Paragraphs[0]; edited.Runs[0].Bold = false;
        edited.Runs[0].Color = OdfColor.Parse("#ff0000"); edited.Hyperlinks[0].Href = "https://example.invalid/edited";
        edited.SetTabStops(new[] { new OdfTabStop(P(40)) });
        copy.MasterShapes[0].Lists[0].Items[0].StartValue = 7;
        copy.MasterShapes[0].Lists[0].Items[0].AddParagraph("Added");
        copy.MasterShapes[0].Lists[0].Items[0].Lists[0].AddItem("Second nested");
        foreach (var read in RoundTrips(document)) {
            var original = read.Pages[0].MasterShapes[0]; var changed = read.Pages[1].MasterShapes[0];
            Assert.True(original.Paragraphs[0].Runs[0].Bold); Assert.False(changed.Paragraphs[0].Runs[0].Bold);
            Assert.Equal(OdfColor.Parse("#2040e0"), original.Paragraphs[0].Runs[0].Color);
            Assert.Equal(OdfColor.Parse("#ff0000"), changed.Paragraphs[0].Runs[0].Color);
            Assert.Equal("https://example.invalid/edited", changed.Paragraphs[0].Hyperlinks[0].Href);
            Assert.Equal("https://example.invalid/master", original.Paragraphs[0].Hyperlinks[0].Href);
            Assert.False(read.Pages[0].Shapes[0].Lists[0].IsOrdered);
            Assert.True(changed.Lists[0].IsOrdered); Assert.False(changed.Lists[0].Items[0].Lists[0].IsOrdered);
            Assert.Equal(3, original.Lists[0].Items[0].StartValue); Assert.Equal(7, changed.Lists[0].Items[0].StartValue);
            Assert.Equal(3, changed.Lists[0].Items[0].Paragraphs.Count);
            Assert.Equal(2, changed.Lists[0].Items[0].Lists[0].Items.Count);
            var tabs = changed.Paragraphs[0].Styles.First().ParagraphProperties!.Element(OdfNamespaces.Style + "tab-stops")!;
            Assert.Equal("40pt", (string?)tabs.Elements().Single().Attribute(OdfNamespaces.Style + "position"));
        }
    }

    [Fact]
    public void FieldOnlyEditsDirtyTheMasterPartAndRetainNativeFieldAttributes() {
        var document = OdgDocument.Create(); var page = document.AddPage("Field");
        page.MasterShapes.AddTextBox(R(10, 10, 80, 40), "").Paragraphs[0].AddField(OdfTextFieldKind.Date, "Old");
        document = Reopen(document); var field = document.Pages[0].MasterShapes[0].Paragraphs[0].Fields[0];
        field.DisplayText = "2026-10-07"; field.IsFixed = true;
        foreach (var read in RoundTrips(document)) {
            var saved = read.Pages[0].MasterShapes[0].Paragraphs[0].Fields[0];
            Assert.Equal("2026-10-07", saved.DisplayText); Assert.True(saved.IsFixed); Assert.Equal(OdfTextFieldKind.Date, saved.Kind);
        }
    }

    [Fact]
    public void EditingIndependentMasterTextPreservesNativeArtworkAndOwnedImages() {
        var document = OdgDocument.Load(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Drawing", "libreoffice-master-artwork.odg"));
        var page = document.Pages[0];
        var opaque = page.MasterShapes.Where(s => s.ElementName == "custom-shape").Select(s => s.ToXml().ToString()).ToArray();
        Assert.NotEmpty(opaque);
        var text = page.MasterShapes.Last(); text.Name = "EditedHeading"; text.Text = "Edited native master"; text.FontSize = P(14);
        var added = page.MasterShapes.AddImage(Image(), "added.png", R(10, 10, 20, 20), "Added");
        foreach (var read in RoundTrips(document)) {
            var master = read.Pages[0].MasterShapes;
            Assert.Equal("Edited native master", master.Single(s => s.Name == text.Name).Text);
            Assert.Equal(Image(), master.Single(s => s.Name == added.Name).GetImageBytes());
            Assert.Equal(opaque, master.Where(s => s.ElementName == "custom-shape").Select(s => s.ToXml().ToString()));
            Assert.True(read.Pages[0].ToDrawing().Report.HasSkippedOrUnsupported);
        }
    }

    [Fact]
    public void NativeListEditingCreatesAnOmittedAutomaticStyleContainerBeforeTheBody() {
        var document = OdgDocument.Create(); var page = document.AddPage("Native text");
        var rectangle = new XElement(OdfNamespaces.Draw + "rect"); OdfShape.ApplyBounds(rectangle, R(10, 10, 120, 80));
        rectangle.Add(new XElement(OdfNamespaces.Text + "p", "Native paragraph")); page.Element.Add(rectangle);
        document.GetXml("content.xml").Root!.Element(OdfNamespaces.Office + "automatic-styles")!.Remove();
        document.MarkPartDirty("content.xml"); document = Reopen(document);
        document.Pages[0].Shapes[0].AddList(true).AddItem("Added numbered item");
        foreach (var read in RoundTrips(document)) {
            Assert.True(read.Pages[0].Shapes[0].Lists[0].IsOrdered);
            Assert.Equal("Added numbered item", read.Pages[0].Shapes[0].Lists[0].Items[0].Paragraphs[0].Text);
            var content = read.GetXml("content.xml").Root!.Elements().Select(e => e.Name).ToArray();
            Assert.True(Array.IndexOf(content, OdfNamespaces.Office + "automatic-styles") < Array.IndexOf(content, OdfNamespaces.Office + "body"));
        }
    }

    [Theory]
    [InlineData("master")]
    [InlineData("nested")]
    [InlineData("page")]
    public void AddingListsPreservesExistingCommonNumberedListFormatting(string scope) {
        var document = OdgDocument.Create(); var page = document.AddPage("Common list");
        var shape = (scope == "page" ? page.Shapes : page.MasterShapes).AddTextBox(R(10, 10, 160, 100), "", "Text");
        shape.AddList(true).AddItem("Original numbered item");
        string originalName = shape.Lists[0].StyleName!;
        var part = scope == "page" ? "content.xml" : "styles.xml";
        var definition = document.GetXml(part).Descendants(OdfNamespaces.Text + "list-style").Single();
        definition.Remove(); document.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "styles")!.Add(definition);
        document.MarkPartDirty(part); document.MarkPartDirty("styles.xml"); document = Reopen(document);
        shape = (scope == "page" ? document.Pages[0].Shapes : document.Pages[0].MasterShapes)[0];
        Assert.True(shape.Lists[0].IsOrdered);
        var added = scope == "nested" ? shape.Lists[0].Items[0].AddList() : shape.AddList();
        added.AddItem("New bullet item"); Assert.NotEqual(originalName, added.StyleName);
        foreach (var read in RoundTrips(document)) {
            var saved = (scope == "page" ? read.Pages[0].Shapes : read.Pages[0].MasterShapes)[0];
            Assert.True(saved.Lists[0].IsOrdered); Assert.Equal(originalName, saved.Lists[0].StyleName);
            var text = Assert.Single(read.Pages[0].ToDrawing().Value.Elements.OfType<OfficeDrawingRichText>());
            Assert.Equal("1.", text.Paragraphs.First(p => p.Label != null).Label!.Run.Text);
            Assert.False((scope == "nested" ? saved.Lists[0].Items[0].Lists[0] : saved.Lists[1]).IsOrdered);
        }
    }

    private static OdgDocument Reopen(OdgDocument document) {
        using var stream = new MemoryStream(); document.Save(stream); stream.Position = 0; return OdgDocument.Load(stream);
    }
    private static IEnumerable<OdgDocument> RoundTrips(OdgDocument document) {
        yield return Reopen(document);
        using var stream = new MemoryStream(); document.SaveFlatXml(stream); stream.Position = 0; yield return OdgDocument.LoadFlatXml(stream);
    }
    private static OdfLength P(double value) => OdfLength.Points(value);
    private static OdfRect R(double x, double y, double width, double height) => new OdfRect(P(x), P(y), P(width), P(height));
    private static byte[] Image() {
        var bitmap = new OfficeRasterImage(2, 2); bitmap.Fill(OfficeColor.Parse("#2040e0")); return OfficePngWriter.Encode(bitmap);
    }
}
