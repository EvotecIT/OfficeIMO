using System;
using System.IO;
using System.Text;
using System.Xml.Linq;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsPdfNavigationTests {
    [Theory]
    [InlineData("https://example.com/report", "https://example.com/report")]
    [InlineData("https://example.com/ż?q=é", "https://example.com/%C5%BC?q=%C3%A9")]
    public void NativeWebLinkExportsThroughSharedDrawingWriter(string uri, string expected) {
        var doc = XpsDocument.Create(); var page = doc.AddPage(200, 100).AddPath("M20,30H60V50H20Z");
        var xml = page.GetMarkup(); xml.Elements().Single().SetAttributeValue("FixedPage.NavigateUri", uri); page.ReplaceMarkup(xml);
        var read = PdfReadDocument.Open(doc.ToPdf()); var link = Assert.Single(read.Pages[0].GetLinkAnnotations());
        Assert.Equal(expected, link.Uri);
    }

    [Fact]
    public void DrawingFragmentLinksResolveAndSurviveClippedTransformedGroups() {
        var child = new OfficeDrawing(100, 100).AddLink("#chapter", 10, 20, 50, 40, "Open\u00a0report");
        var clip = new OfficeDrawing(100, 100).AddClippedDrawing(child, 0, 0, OfficeClipPath.Rectangle(40, 45));
        var drawing = new OfficeDrawing(200, 200).AddEffectDrawing(clip, OfficeTransform.Translate(20, 30));
        var pdf = PdfDocument.Create();
        pdf.Compose(b => b.Page(p => p.Size(200, 200).Margin(new PageMargins(0, 0, 0, 0)).Canvas(c => {
            c.NamedDestination("chapter", 25, 40); c.Drawing(drawing, 0, 0, 200, 200);
        })));
        using var output = new MemoryStream(); pdf.Save(output);
        var read = PdfReadDocument.Open(output.ToArray());
        var link = Assert.Single(read.Pages[0].GetLinkAnnotations());
        Assert.Equal("chapter", link.DestinationName);
        Assert.Equal("Open\u00a0report", link.Contents);
        Assert.Equal(1, Assert.Single(read.NamedDestinations).PageNumber);
        Assert.Equal(25, Assert.Single(read.NamedDestinations).DestinationLeft);
        Assert.Equal(30, link.X1); Assert.Equal(60, link.X2);
        Assert.Equal(125, link.Y1); Assert.Equal(150, link.Y2);
    }
    [Theory]
    [InlineData(XpsFormat.Xps)]
    [InlineData(XpsFormat.OpenXps)]
    public void NativeNamedTargetsAndOutlinesMapToPdfAfterPageMoves(XpsFormat format) {
        var doc = XpsDocumentStructureTests.Fixture(format); var source = doc.Pages[0]; var target = doc.Pages[1];
        var xml = target.GetMarkup(); var ns = xml.Name.Namespace;
        xml.Add(new XElement(ns + "Canvas", new XAttribute("RenderTransform", "1,0,0,1,10,5"),
            new XElement(ns + "Path", new XAttribute("Name", "section"), new XAttribute("Data", "M80,40H120V60H80Z"), new XAttribute("Fill", "#FF000000"))));
        target.ReplaceMarkup(xml);
        source.AddPath("M10,10H30V30H10Z"); source.AddPath("M40,10H60V30H40Z");
        xml = source.GetMarkup(); xml.Elements().First().SetAttributeValue("FixedPage.NavigateUri", "/" + target.PartName + "#%73ection");
        xml.Elements().Last().SetAttributeValue("FixedPage.NavigateUri", "/" + source.PartName); source.ReplaceMarkup(xml);
        const string part = "Documents/1/Structure/DocumentStructure.struct";
        var structure = XElement.Parse(Encoding.UTF8.GetString(doc.GetPartBytes(part)));
        var outline = structure.Descendants(doc.StructureNamespace + "OutlineEntry").Single();
        outline.SetAttributeValue("Description", "Chapter € “quoted” —"); outline.SetAttributeValue("OutlineTarget", "/" + target.PartName + "#%73ection");
        outline.AddAfterSelf(new XElement(doc.StructureNamespace + "OutlineEntry", new XAttribute("OutlineLevel", "2"), new XAttribute("Description", "Web"), new XAttribute("OutlineTarget", "https://example.com/ż?q=é")));
        structure.Elements(doc.StructureNamespace + "Story").Remove();
        doc.Documents[0].ReplaceDocumentStructureMarkup(structure);
        Check(doc.ToPdf(), 1, 2);
        doc.Documents[0].MovePage(1, 0);
        Check(doc.ToPdf(), 2, 1);
        void Check(byte[] bytes, int sourcePage, int targetPage) {
            var read = PdfReadDocument.Open(bytes); var links = read.Pages[sourcePage - 1].GetLinkAnnotations(); Assert.Equal(2, links.Count);
            var destination = read.NamedDestinations.Single(d => d.Name == links[0].DestinationName);
            Assert.Equal(targetPage, destination.PageNumber); Assert.Equal(67.5, destination.DestinationLeft); Assert.Equal(758.25, destination.DestinationTop);
            Assert.Equal(sourcePage, read.NamedDestinations.Single(d => d.Name == links[1].DestinationName).PageNumber);
            var bookmark = Assert.Single(read.Outlines); Assert.Equal("Chapter € “quoted” —", bookmark.Title); Assert.Equal(targetPage, bookmark.PageNumber);
            Assert.Equal(67.5, bookmark.DestinationLeft); Assert.Equal("Web", Assert.Single(bookmark.Children).Title);
            Assert.Contains("/URI (https://example.com/%C5%BC?q=%C3%A9)", Encoding.ASCII.GetString(bytes));
        }
    }

    [Fact]
    public void EncodedSequenceTargetsKeepTheirPageAfterInsertion() {
        var doc = XpsDocumentStructureTests.Fixture(XpsFormat.Xps);
        var source = doc.Pages[0]; var target = doc.Pages[1];
        source.AddPath("M0,0H20V20H0Z"); var xml = source.GetMarkup();
        xml.Elements().Single().SetAttributeValue("FixedPage.NavigateUri", "/FixedDocumentSequence.fdseq#%32"); source.ReplaceMarkup(xml);
        const string part = "Documents/1/Structure/DocumentStructure.struct";
        var structure = XElement.Parse(Encoding.UTF8.GetString(doc.GetPartBytes(part)));
        structure.Descendants(doc.StructureNamespace + "OutlineEntry").Single().SetAttributeValue("OutlineTarget", "/FixedDocumentSequence.fdseq#%32");
        structure.Elements(doc.StructureNamespace + "Story").Remove();
        doc.Documents[0].ReplaceDocumentStructureMarkup(structure);
        doc.Documents[0].MovePage(1, 0);
        var read = PdfReadDocument.Open(doc.ToPdf()); var link = Assert.Single(read.Pages[1].GetLinkAnnotations());
        Assert.Equal(1, read.NamedDestinations.Single(d => d.Name == link.DestinationName).PageNumber);
        Assert.Equal(1, Assert.Single(read.Outlines).PageNumber);
        Assert.Contains(target.PartName, source.GetMarkup().ToString());
    }

    [Fact]
    public void LocalFragmentsStayOnRepeatedOccurrenceWhileAbsoluteLinksUseFirst() {
        var doc = XpsDocument.Create(); var page = doc.AddPage(100, 100);
        page.AddPath("M10,10H20V20H10Z"); page.AddPath("M30,10H40V20H30Z");
        var xml = page.GetMarkup(); xml.SetAttributeValue("Name", "root");
        xml.Elements().First().SetAttributeValue("FixedPage.NavigateUri", "#root");
        xml.Elements().Last().SetAttributeValue("FixedPage.NavigateUri", "/" + page.PartName + "#root"); page.ReplaceMarkup(xml);
        doc.Documents[0].InsertPage(1, page);
        var read = PdfReadDocument.Open(doc.ToPdf()); var links = read.Pages[1].GetLinkAnnotations(); Assert.Equal(2, links.Count);
        Assert.Equal(2, read.NamedDestinations.Single(d => d.Name == links[0].DestinationName).PageNumber);
        Assert.Equal(1, read.NamedDestinations.Single(d => d.Name == links[1].DestinationName).PageNumber);
    }

    [Fact]
    public void DrawingLinksInMasksAreExcludedAndPatternLinksStayInsideTheirArea() {
        var tile = new OfficeDrawing(20, 20).AddLink("https://example.com/tile", 0, 0, 20, 20);
        var pattern = new OfficeDrawing(100, 100).AddTilingPattern(tile, new OfficeImagePlacement(15, 15, 25, 25), 20, 20);
        var mask = new OfficeDrawing(100, 100).AddShape(OfficeShape.Rectangle(100, 100), 0, 0).AddLink("https://example.com/mask", 0, 0, 100, 100);
        var drawing = new OfficeDrawing(100, 100).AddEffectDrawing(pattern, OfficeTransform.Identity, OfficeBlendMode.Normal, new OfficeDrawingSoftMask(mask));
        var pdf = PdfDocument.Create(); pdf.Compose(b => b.Page(p => p.Size(100, 100).Margin(new PageMargins(0, 0, 0, 0)).Canvas(c => c.Drawing(drawing, 0, 0, 100, 100))));
        using var output = new MemoryStream(); pdf.Save(output);
        var links = PdfReadDocument.Open(output.ToArray()).Pages[0].GetLinkAnnotations(); Assert.NotEmpty(links);
        foreach (var link in links) {
            Assert.Equal("https://example.com/tile", link.Uri);
            Assert.InRange(link.X1, 15, 40); Assert.InRange(link.X2, 15, 40);
            Assert.InRange(link.Y1, 60, 85); Assert.InRange(link.Y2, 60, 85);
        }
    }

}
