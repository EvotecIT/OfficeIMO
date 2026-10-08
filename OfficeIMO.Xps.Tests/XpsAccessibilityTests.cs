using System;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Xps;
using Xunit;
using static OfficeIMO.Xps.Tests.XpsLogicalStructureTests;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsAccessibilityTests {
    private const string Short = "A dog & its bowl";
    private const string Long = "A brown dog beside a <blue> bowl. 犬";
    private const string Alternative = Short + "\n" + Long;

    [Theory]
    [InlineData(XpsFormat.Xps, false)]
    [InlineData(XpsFormat.Xps, true)]
    [InlineData(XpsFormat.OpenXps, false)]
    [InlineData(XpsFormat.OpenXps, true)]
    public void AuthoredGraphicDescriptionsReachSvgFigureTagsAndReaderAssets(XpsFormat format, bool canvas) {
        var doc = Described(format, canvas); var page = doc.Pages[0];
        page.ReplaceStoryFragmentsMarkup(Fragments(doc, null, null, new XElement(doc.StructureNamespace + "FigureStructure",
            new XElement(doc.StructureNamespace + "NamedElement", new XAttribute("NameReference", "Graphic")))));
        doc = XpsDocument.Load(doc.Save()); page = doc.Pages[0];
        var native = Assert.Single(Assert.Single(doc.ReadLogicalStructure().Stories).Blocks).Children[0].Content!;
        Assert.Equal(Short, native.AccessibilityName); Assert.Equal(Long, native.AccessibilityHelpText);
        var svg = XElement.Parse(page.ToSvg().Svg); XNamespace ns = svg.Name.Namespace;
        Assert.Equal(Short, Assert.Single(svg.Descendants(ns + "title")).Value);
        Assert.Equal(Long, Assert.Single(svg.Descendants(ns + "desc")).Value);
        var read = PdfReadDocument.Open(doc.ToPdf());
        Assert.Equal(Alternative, Assert.Single(read.TaggedContent!.StructureElements, e => e.StructureType == "Figure").AlternateText);
        var model = doc.ToOfficeDocumentModel();
        var asset = Assert.Single(model.Assets);
        Assert.Equal(Short, asset.Title); Assert.Equal(Alternative, asset.AltText);
        Assert.Equal(1, asset.Location.Page); Assert.Null(asset.PayloadBytes);
        var result = doc.ToOfficeDocumentReadResult();
        var restored = OfficeDocumentReadResultJson.Deserialize(OfficeDocumentReadResultJson.Serialize(result));
        Assert.Equal(Short, Assert.Single(restored.Assets).Title);
        Assert.Equal(Alternative, Assert.Single(restored.Assets).AltText);
        Assert.Empty(result.Blocks); Assert.Empty(result.Chunks);
        Assert.Equal(OfficeColor.FromRgb(51, 102, 153), OfficeDrawingRasterRenderer.Render(page.ToDrawing(), scale: 1).GetPixel(20, 20));
    }

    [Theory]
    [InlineData(XpsFormat.Xps)]
    [InlineData(XpsFormat.OpenXps)]
    public void DescriptionsDoNotInventNativeStoryStructureOrSearchText(XpsFormat format) {
        var doc = Described(format, false);
        Assert.False(doc.ReadLogicalStructure().HasNativeStructure);
        Assert.Single(doc.ToOfficeDocumentReadResult().Assets);
        var read = PdfReadDocument.Open(doc.ToPdf());
        Assert.Null(read.TaggedContent); Assert.Empty(read.Pages[0].GetTextSpans());
    }

    [Fact]
    public void FigureDescriptionUsesEveryAuthoredTargetInReferenceOrder() {
        var doc = Described(XpsFormat.OpenXps, false); var page = doc.Pages[0]; var xml = page.GetMarkup();
        var extra = new XElement(xml.Elements().Single()); extra.SetAttributeValue("Name", "Second");
        extra.SetAttributeValue("AutomationProperties.Name", "Second graphic"); extra.Attribute("AutomationProperties.HelpText")!.Remove();
        xml.Add(extra); page.ReplaceMarkup(xml);
        page.ReplaceStoryFragmentsMarkup(Fragments(doc, null, null, new XElement(doc.StructureNamespace + "FigureStructure",
            new XElement(doc.StructureNamespace + "NamedElement", new XAttribute("NameReference", "Second")),
            new XElement(doc.StructureNamespace + "NamedElement", new XAttribute("NameReference", "Graphic")))));
        var read = PdfReadDocument.Open(doc.ToPdf());
        Assert.Equal("Second graphic\n" + Alternative, Assert.Single(read.TaggedContent!.StructureElements, e => e.StructureType == "Figure").AlternateText);
    }

    [Fact]
    public void PreviewPageLimitDoesNotCountDescriptionOnlyAssets() {
        var doc = Described(XpsFormat.OpenXps, false); var page = doc.Pages[0]; var xml = page.GetMarkup();
        var first = xml.Elements().Single();
        for (int i = 0; i < 512; i++) {
            var extra = new XElement(first); extra.Attribute("Name")!.Remove(); xml.Add(extra);
        }
        page.ReplaceMarkup(xml); doc.AddPage(100, 80);
        var model = doc.ToOfficeDocumentModel(includeSvgPreviewAssets: true);
        Assert.Equal(2, model.Assets.Count(asset => asset.Kind == "page-preview"));
        Assert.Equal(513, model.Assets.Count(asset => asset.Kind == "graphic-description"));
    }

    internal static XpsDocument Described(XpsFormat format, bool canvas) {
        var doc = XpsDocument.Create(format); var page = doc.AddPage(100, 80).AddPath("M10,10H70V60H10Z", "#FF336699");
        var xml = page.GetMarkup(); XElement element = xml.Elements().Single();
        if (canvas) { element.Remove(); element = new XElement(xml.Name.Namespace + "Canvas", element); xml.Add(element); }
        element.SetAttributeValue("Name", "Graphic"); element.SetAttributeValue("AutomationProperties.Name", Short);
        element.SetAttributeValue("AutomationProperties.HelpText", Long); page.ReplaceMarkup(xml);
        return doc;
    }
}
