using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public class OpenDocumentOdgCloningTests {
    [Fact]
    public void ClonedConnectedPageEditsIndependentlyAndReusesImageBytesAcrossPackageAndFlatOutput() {
        var document = OdgDocument.Create(); var page = document.AddPage("Source");
        page.Layers.Add("Content");
        var group = page.Shapes.AddGroup("Group");
        var start = group.Children.AddRectangle(Rect(20, 30, 40, 30), "Start");
        start.Text = "Original"; start.Layer = "Content"; start.FillColor = OdfColor.Parse("123456");
        var end = page.Shapes.AddRectangle(Rect(200, 30, 40, 30), "End");
        var connector = page.Shapes.AddConnector(start.AddGluePoint(OdgGluePointAlignment.Right), end.AddGluePoint(OdgGluePointAlignment.Left), "Link");
        connector.SetConnectorRoute(new[] { OfficePathCommand.MoveTo(60, 45), OfficePathCommand.LineTo(200, 45) });
        byte[] image = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
        page.Shapes.AddImage(image, "pixel.png", Rect(20, 100, 20, 20), "Image");
        string[] entries = document.PackageEntries.ToArray();
        OdgPage copy = document.ClonePage(0, "Copy");
        Assert.Equal(entries, document.PackageEntries);
        Assert.Equal(page.MasterPageName, copy.MasterPageName);
        Assert.Equal(connector.ConnectorRouteCommands, copy.Shapes[2].ConnectorRouteCommands);
        var copiedStart = copy.Shapes[0].Children[0];
        Assert.NotEqual(start.XmlId, copiedStart.XmlId);
        Assert.Equal(copiedStart.XmlId, copy.Shapes[2].StartShapeId);
        Assert.Equal(start.GluePoints[0].Id, copiedStart.GluePoints[0].Id);
        copiedStart.Text = "Changed"; copiedStart.FillColor = OdfColor.Parse("ABCDEF");
        copiedStart.Bounds = Rect(30, 40, 50, 30);
        copy.Shapes[2].SetConnectorRoute(new[] { OfficePathCommand.MoveTo(80, 55), OfficePathCommand.LineTo(200, 45) });
        copy.Layers.Find("Content")!.IsProtected = true;
        foreach (var read in RoundTrips(document)) {
            var original = read.Pages[0]; var cloned = read.Pages[1];
            Assert.Equal("Original", original.Shapes[0].Children[0].Text);
            Assert.Equal(OdfColor.Parse("123456"), original.Shapes[0].Children[0].FillColor);
            Assert.Equal("Changed", cloned.Shapes[0].Children[0].Text);
            Assert.Equal(OdfColor.Parse("ABCDEF"), cloned.Shapes[0].Children[0].FillColor);
            Assert.Equal(60, original.Shapes[2].X1.ToPoints(), 3);
            Assert.Equal(80, cloned.Shapes[2].X1.ToPoints(), 3);
            Assert.False(original.Layers.Find("Content")!.IsProtected);
            Assert.True(cloned.Layers.Find("Content")!.IsProtected);
            Assert.Equal(image, cloned.Shapes[3].GetImageBytes());
            Assert.Equal(original.Shapes[3].GetImageBytes(), cloned.Shapes[3].GetImageBytes());
        }
    }

    [Fact]
    public void RemapsDualIdentifiersNavigationTextChainsAndLocalHyperlinksWithoutChangingSource() {
        var document = OdgDocument.Create(); var page = document.AddPage("Source");
        var first = page.Shapes.AddTextBox(Rect(0, 0, 50, 20), "First", "FirstFrame");
        var second = page.Shapes.AddTextBox(Rect(0, 30, 50, 20), "Second", "SecondFrame");
        first.XmlId = "first"; second.XmlId = "second";
        first.Element.SetAttributeValue(OdfNamespaces.Draw + "id", "first");
        first.Element.Element(OdfNamespaces.Draw + "text-box")!.SetAttributeValue(OdfNamespaces.Draw + "chain-next-name", second.Name);
        first.Element.Element(OdfNamespaces.Draw + "text-box")!.Element(OdfNamespaces.Text + "p")!.Add(
            new XElement(OdfNamespaces.Text + "a", new XAttribute(OdfNamespaces.XLink + "href", "#second"), "Shape"),
            new XElement(OdfNamespaces.Text + "a", new XAttribute(OdfNamespaces.XLink + "href", "#Source"), "Page"),
            new XElement(OdfNamespaces.Text + "a", new XAttribute(OdfNamespaces.XLink + "href", "https://example.com/"), "External"));
        page.Element.SetAttributeValue(OdfNamespaces.Draw + "nav-order", "first second");
        document.MarkPartDirty("content.xml");
        string before = page.Element.ToString();
        var clone = document.ClonePage(0, "Copied");
        Assert.Equal(before, page.Element.ToString());
        foreach (var read in RoundTrips(document)) {
            var copied = read.Pages[1]; var a = copied.Shapes[0]; var b = copied.Shapes[1];
            Assert.NotEqual("first", a.XmlId); Assert.NotEqual("second", b.XmlId);
            Assert.Equal(a.XmlId, (string?)a.Element.Attribute(OdfNamespaces.Draw + "id"));
            Assert.Equal(a.XmlId + " " + b.XmlId, (string?)copied.Element.Attribute(OdfNamespaces.Draw + "nav-order"));
            Assert.Equal(b.Name, (string?)a.Element.Element(OdfNamespaces.Draw + "text-box")!.Attribute(OdfNamespaces.Draw + "chain-next-name"));
            Assert.Equal(new[] { "#" + b.XmlId, "#Copied", "https://example.com/" },
                a.Element.Descendants(OdfNamespaces.Text + "a").Select(element => (string?)element.Attribute(OdfNamespaces.XLink + "href")));
        }
    }

    [Fact]
    public void RemapsUriEscapedPageAndFrameNames() {
        var document = OdgDocument.Create(); var source = document.AddPage("Source page");
        var frame = source.Shapes.AddTextBox(Rect(0, 0, 100, 20), "Links", "Named frame");
        frame.Element.Element(OdfNamespaces.Draw + "text-box")!.Element(OdfNamespaces.Text + "p")!.Add(
            new XElement(OdfNamespaces.Text + "a", new XAttribute(OdfNamespaces.XLink + "href", "#Source%20page"), "Page"),
            new XElement(OdfNamespaces.Text + "a", new XAttribute(OdfNamespaces.XLink + "href", "#Named%20frame"), "Frame"));
        document.MarkPartDirty("content.xml"); document.ClonePage(0, "Copied page");
        foreach (var read in RoundTrips(document)) {
            var copied = read.Pages[1].Shapes[0];
            Assert.Equal(new[] { "#Copied%20page", "#" + Uri.EscapeDataString(copied.Name) },
                copied.Element.Descendants(OdfNamespaces.Text + "a").Select(element => (string?)element.Attribute(OdfNamespaces.XLink + "href")));
        }
    }

    [Fact]
    public void RemapsContinuedListsAndNoteReferencesWhileRetainingExternalNoteTargets() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var frame = page.Shapes.AddTextBox(Rect(0, 0, 100, 100), "Text");
        var text = frame.Element.Element(OdfNamespaces.Draw + "text-box")!;
        text.Add(new XElement(OdfNamespaces.Text + "list", new XAttribute(XNamespace.Xml + "id", "list"),
            new XElement(OdfNamespaces.Text + "list-item", new XElement(OdfNamespaces.Text + "p", "First"))),
            new XElement(OdfNamespaces.Text + "list", new XAttribute(OdfNamespaces.Text + "continue-list", "list"),
                new XElement(OdfNamespaces.Text + "list-item", new XElement(OdfNamespaces.Text + "p", "Second"))));
        text.Element(OdfNamespaces.Text + "p")!.Add(new XElement(OdfNamespaces.Text + "note", new XAttribute(OdfNamespaces.Text + "id", "note"),
            new XAttribute(OdfNamespaces.Text + "note-class", "footnote"), new XElement(OdfNamespaces.Text + "note-citation", "1"),
            new XElement(OdfNamespaces.Text + "note-body", new XElement(OdfNamespaces.Text + "p", "Detail"))),
            new XElement(OdfNamespaces.Text + "note-ref", new XAttribute(OdfNamespaces.Text + "ref-name", "note"), "1"),
            new XElement(OdfNamespaces.Text + "note-ref", new XAttribute(OdfNamespaces.Text + "ref-name", "outside"), "2"),
            new XElement(OdfNamespaces.Text + "bookmark-ref", new XAttribute(OdfNamespaces.Text + "ref-name", "note"), "External bookmark"));
        document.MarkPartDirty("content.xml"); document.ClonePage(0);
        foreach (var read in RoundTrips(document)) {
            XElement copy = read.Pages[1].Shapes[0].Element;
            var lists = copy.Descendants(OdfNamespaces.Text + "list").ToArray();
            string id = (string)lists[0].Attribute(XNamespace.Xml + "id")!;
            Assert.NotEqual("list", id);
            Assert.Equal(id, (string?)lists[1].Attribute(OdfNamespaces.Text + "continue-list"));
            string note = (string)copy.Descendants(OdfNamespaces.Text + "note").Single().Attribute(OdfNamespaces.Text + "id")!;
            Assert.NotEqual("note", note);
            Assert.Equal(new[] { note, "outside" }, copy.Descendants(OdfNamespaces.Text + "note-ref").Select(element => (string?)element.Attribute(OdfNamespaces.Text + "ref-name")));
            Assert.Equal("note", (string?)copy.Descendants(OdfNamespaces.Text + "bookmark-ref").Single().Attribute(OdfNamespaces.Text + "ref-name"));
        }
    }

    [Fact]
    public void NumberedParagraphCopiesUseIndependentListGroups() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var text = page.Shapes.AddTextBox(Rect(0, 0, 100, 100), "Lists").Element.Element(OdfNamespaces.Draw + "text-box")!;
        foreach (var item in new[] { ("ListA", "First"), ("ListA", "Second"), ("ListB", "Other") })
            text.Add(new XElement(OdfNamespaces.Text + "numbered-paragraph", new XAttribute(OdfNamespaces.Text + "list-id", item.Item1),
                new XElement(OdfNamespaces.Text + "p", item.Item2)));
        document.MarkPartDirty("content.xml"); document.ClonePage(0);
        foreach (var read in RoundTrips(document)) {
            var original = read.Pages[0].Element.Descendants(OdfNamespaces.Text + "numbered-paragraph").Attributes(OdfNamespaces.Text + "list-id").Select(attribute => attribute.Value).ToArray();
            var copied = read.Pages[1].Element.Descendants(OdfNamespaces.Text + "numbered-paragraph").Attributes(OdfNamespaces.Text + "list-id").Select(attribute => attribute.Value).ToArray();
            Assert.Equal(new[] { "ListA", "ListA", "ListB" }, original);
            Assert.Equal(copied[0], copied[1]); Assert.NotEqual(copied[0], copied[2]);
            Assert.DoesNotContain(copied[0], original); Assert.DoesNotContain(copied[2], original);
        }
    }

    [Fact]
    public void ExternalNoteNameDoesNotBindToAClonedShapeWithTheSameXmlId() {
        var document = OdgDocument.Create(); var source = document.AddPage("Source");
        var frame = source.Shapes.AddTextBox(Rect(0, 0, 100, 20), "Reference"); frame.XmlId = "outside";
        frame.Element.Element(OdfNamespaces.Draw + "text-box")!.Element(OdfNamespaces.Text + "p")!.Add(
            new XElement(OdfNamespaces.Text + "note-ref", new XAttribute(OdfNamespaces.Text + "ref-name", "outside"), "1"));
        var other = document.AddPage("Notes").Shapes.AddTextBox(Rect(0, 0, 100, 20), "Notes");
        other.Element.Element(OdfNamespaces.Draw + "text-box")!.Element(OdfNamespaces.Text + "p")!.Add(
            new XElement(OdfNamespaces.Text + "note", new XAttribute(OdfNamespaces.Text + "id", "outside"),
                new XAttribute(OdfNamespaces.Text + "note-class", "footnote"), new XElement(OdfNamespaces.Text + "note-citation", "1"),
                new XElement(OdfNamespaces.Text + "note-body", new XElement(OdfNamespaces.Text + "p", "External detail"))));
        document.MarkPartDirty("content.xml"); document.ClonePage(0);
        foreach (var read in RoundTrips(document)) {
            var copied = read.Pages[2].Shapes[0];
            Assert.NotEqual("outside", copied.XmlId);
            Assert.Equal("outside", (string?)copied.Element.Descendants(OdfNamespaces.Text + "note-ref").Single().Attribute(OdfNamespaces.Text + "ref-name"));
        }
    }

    [Fact]
    public void ClonedMasterRetainsArtworkAndReferencesWithIndependentLayoutAndLayers() {
        var document = OdgDocument.Create(); var source = document.AddPage("Source", OdfLength.Points(300), OdfLength.Points(200));
        source.MasterLayers.Add("Art", OdgLayerDisplay.Screen);
        var master = source.Master!;
        master.SetAttributeValue(OdfNamespaces.Style + "next-style-name", source.MasterPageName);
        master.Add(new XElement(OdfNamespaces.Draw + "rect", new XAttribute(XNamespace.Xml + "id", "masterArt"),
            new XAttribute(OdfNamespaces.Draw + "name", "MasterArt"), new XAttribute(OdfNamespaces.Draw + "layer", "Art"),
            new XAttribute(OdfNamespaces.Svg + "x", "0cm"), new XAttribute(OdfNamespaces.Svg + "y", "0cm"),
            new XAttribute(OdfNamespaces.Svg + "width", "1cm"), new XAttribute(OdfNamespaces.Svg + "height", "1cm"),
            new XElement(OdfNamespaces.Text + "p", new XElement(OdfNamespaces.Text + "a", new XAttribute(OdfNamespaces.XLink + "href", "#masterArt"), "Logo"))));
        document.MarkPartDirty("styles.xml");
        var copied = document.ClonePage(0, "Copy"); copied.MasterPageName = document.CloneMasterPage(source.MasterPageName, "Independent");
        copied.Width = OdfLength.Points(450); copied.MasterLayers.Find("Art")!.Display = OdgLayerDisplay.None;
        foreach (var read in RoundTrips(document)) {
            Assert.Equal(300, read.Pages[0].Width.ToPoints(), 3);
            Assert.Equal(450, read.Pages[1].Width.ToPoints(), 3);
            Assert.Equal(OdgLayerDisplay.Screen, read.Pages[0].MasterLayers.Find("Art")!.Display);
            Assert.Equal(OdgLayerDisplay.None, read.Pages[1].MasterLayers.Find("Art")!.Display);
            XElement clonedMaster = read.Pages[1].Master!;
            Assert.Equal("Independent", (string?)clonedMaster.Attribute(OdfNamespaces.Style + "next-style-name"));
            var art = clonedMaster.Element(OdfNamespaces.Draw + "rect")!;
            Assert.NotEqual("masterArt", (string?)art.Attribute(XNamespace.Xml + "id"));
            Assert.Equal("#" + (string?)art.Attribute(XNamespace.Xml + "id"), (string?)art.Descendants(OdfNamespaces.Text + "a").Single().Attribute(OdfNamespaces.XLink + "href"));
            Assert.DoesNotContain(read.Pages[1].ToDrawing().Value.Elements.OfType<OfficeDrawingRichText>(), text => text.PlainText.Contains("Logo"));
            Assert.Contains(read.Pages[0].ToDrawing().Value.Elements.OfType<OfficeDrawingRichText>(), text => text.PlainText.Contains("Logo"));
        }
    }

    [Theory]
    [InlineData("libreoffice-routed-glue.odg")]
    [InlineData("libreoffice-sheared-geometry.odg")]
    public void ClonesIndependentProducerGeometryAndRoutesWithoutChangingRenderedPlacement(string fixture) {
        var document = OdgDocument.Load(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Drawing", fixture));
        var source = document.Pages[0];
        string sourceXml = source.Element.ToString();
        string before = OfficeDrawingSvgExporter.ToSvg(source.ToDrawing().Value);
        var clone = document.ClonePage(0);
        Assert.Equal(sourceXml, source.Element.ToString());
        Assert.Equal(before, OfficeDrawingSvgExporter.ToSvg(clone.ToDrawing().Value));
        var reopened = OdgDocument.Load(new MemoryStream(document.ToBytes(new OdfSaveOptions { CompatibilityProfile = OdfCompatibilityProfile.PreserveSource })));
        Assert.Equal(before, OfficeDrawingSvgExporter.ToSvg(reopened.Pages.Last().ToDrawing().Value));
        var connectors = reopened.Pages.Last().Shapes.Where(shape => shape.IsConnector).ToArray();
        Assert.All(connectors, connector => Assert.Contains(reopened.Pages.Last().Shapes, shape => shape.XmlId == connector.StartShapeId));
    }

    [Theory]
    [InlineData("object")]
    [InlineData("forms")]
    [InlineData("bookmark")]
    [InlineData("animation")]
    [InlineData("attachment")]
    [InlineData("duplicateId")]
    [InlineData("chain")]
    [InlineData("indexRange")]
    public void UnsupportedOrAmbiguousCloningLeavesAllPartsUnchanged(string invalid) {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var a = page.Shapes.AddTextBox(Rect(0, 0, 20, 20), "Text"); a.XmlId = "a";
        var b = page.Shapes.AddRectangle(Rect(50, 0, 20, 20));
        switch (invalid) {
            case "object": a.Element.Add(new XElement(OdfNamespaces.Draw + "object")); break;
            case "forms": page.Element.Add(new XElement(OdfNamespaces.Office + "forms")); break;
            case "bookmark": a.Element.Add(new XElement(OdfNamespaces.Text + "bookmark", new XAttribute(OdfNamespaces.Text + "name", "mark"))); break;
            case "animation": page.Element.Add(new XElement(XName.Get("par", "urn:oasis:names:tc:opendocument:xmlns:animation:1.0"))); break;
            case "attachment": page.Shapes.AddConnector(new OfficePoint(0, 0), new OfficePoint(50, 0)).Element.SetAttributeValue(OdfNamespaces.Draw + "start-shape", "missing"); break;
            case "duplicateId": b.Element.SetAttributeValue(XNamespace.Xml + "id", "a"); break;
            case "chain": a.Element.Element(OdfNamespaces.Draw + "text-box")!.SetAttributeValue(OdfNamespaces.Draw + "chain-next-name", "missing"); break;
            case "indexRange": a.Element.Add(new XElement(OdfNamespaces.Text + "alphabetical-index-mark-start", new XAttribute(OdfNamespaces.Text + "id", "range")),
                new XElement(OdfNamespaces.Text + "alphabetical-index-mark-end", new XAttribute(OdfNamespaces.Text + "id", "range"))); break;
        }
        document.MarkPartDirty("content.xml");
        string[] before = Parts(document);
        if (invalid is "object" or "forms" or "bookmark" or "animation" or "indexRange") Assert.Throws<NotSupportedException>(() => document.ClonePage(0));
        else Assert.Throws<InvalidDataException>(() => document.ClonePage(0));
        Assert.Equal(before, Parts(document));
    }

    [Fact]
    public void MissingMasterLayoutAndDuplicatePageNamesAreRejectedBeforeMutation() {
        var document = OdgDocument.Create(); var page = document.AddPage("Original");
        string[] before = Parts(document);
        Assert.Throws<ArgumentException>(() => document.ClonePage(0, "Original"));
        Assert.Equal(before, Parts(document));
        Assert.Throws<ArgumentException>(() => document.CloneMasterPage(page.MasterPageName, "Review Master"));
        Assert.Throws<ArgumentException>(() => document.CloneMasterPage(page.MasterPageName, "1Master"));
        Assert.Equal(before, Parts(document));
        page.Master!.SetAttributeValue(OdfNamespaces.Style + "page-layout-name", "missing");
        document.MarkPartDirty("styles.xml"); before = Parts(document);
        Assert.Throws<InvalidDataException>(() => document.CloneMasterPage(page.MasterPageName));
        Assert.Equal(before, Parts(document));
    }

    private static string[] Parts(OdgDocument document) => new[] { document.GetXml("content.xml").ToString(), document.GetXml("styles.xml").ToString() };
    private static OdfRect Rect(double x, double y, double width, double height) => new OdfRect(OdfLength.Points(x), OdfLength.Points(y), OdfLength.Points(width), OdfLength.Points(height));
    private static OdgDocument[] RoundTrips(OdgDocument document) {
        using var flat = new MemoryStream(); document.SaveFlatXml(flat); flat.Position = 0;
        return new[] { OdgDocument.Load(new MemoryStream(document.ToBytes())), OdgDocument.LoadFlatXml(flat) };
    }
}
