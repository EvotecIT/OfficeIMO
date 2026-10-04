using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsBoundsAndNavigationTests {
    private static readonly XNamespace Svg = "http://www.w3.org/2000/svg";
    private static readonly XNamespace Key = "http://schemas.openxps.org/oxps/v1.0/resourcedictionary-key";
    private const string DocumentPart = "Documents/1/FixedDocument.fdoc";
    private const string SequencePart = "FixedDocumentSequence.fdseq";

    [Fact]
    public void RepeatedPageAndDocumentReferencesShareEditsAndRetainOrder() {
        var original = XpsDocument.Create(); original.AddPage(100, 100).AddPath("M0,0L10,10"); original.AddPage(200, 100);
        byte[] bytes = XpsDocumentTests.Rewrite(original.Save(), DocumentPart, data => Edit(data, xml => {
            for (int i = 0; i < 100; i++) xml.Add(new XElement(xml.Elements().First()));
        }));
        bytes = XpsDocumentTests.Rewrite(bytes, SequencePart, data => Edit(data, xml => xml.Add(new XElement(xml.Elements().Single()))));
        var loaded = XpsDocument.Load(bytes);
        Assert.Equal(204, loaded.Pages.Count);
        Assert.Equal(200, loaded.Pages[1].Width);
        Assert.Same(loaded.Pages[0], loaded.Pages[2]);
        Assert.Same(loaded.Pages[0], loaded.Pages[102]);
        loaded.Pages[2].AddPath("M0,0L30,30", "#FF0000FF");
        Assert.Equal(2, loaded.Pages[0].GetMarkup().Elements().Count());
        var reopened = XpsDocument.Load(loaded.Save());
        Assert.Equal(204, reopened.Pages.Count);
        Assert.Equal(2, reopened.Pages[102].GetMarkup().Elements().Count());
        Assert.Throws<InvalidDataException>(() => XpsDocument.Load(bytes, new XpsReadOptions { MaximumPages = 203 }));
    }

    [Fact]
    public void DuplicateTargetsUseDocumentScopeAndSequenceFirstOccurrence() {
        var doc = XpsDocument.Create(); var one = doc.AddPage(100, 100); var two = doc.AddPage(100, 100); var three = doc.AddPage(100, 100);
        XNamespace ns = one.GetMarkup().Name.Namespace;
        foreach (var page in doc.Pages) {
            var xml = page.GetMarkup(); xml.SetAttributeValue("Name", "same");
            xml.Add(new XElement(ns + "Path", new XAttribute("Data", "M0,0L10,10"), new XAttribute("FixedPage.NavigateUri", "/Documents/2/FixedDocument.fdoc#same")));
            page.ReplaceMarkup(xml);
        }
        XElement PageRef(string source) => new(ns + "PageContent", new XAttribute("Source", "/" + source), new XElement(ns + "PageContent.LinkTargets", new XElement(ns + "LinkTarget", new XAttribute("Name", "same")), new XElement(ns + "LinkTarget", new XAttribute("Name", "absent"))));
        doc.AddResource("Documents/2/FixedDocument.fdoc", Encoding.UTF8.GetBytes(new XElement(ns + "FixedDocument", PageRef(two.PartName), PageRef(three.PartName)).ToString()), "application/vnd.ms-package.xps-fixeddocument+xml");
        byte[] bytes = XpsDocumentTests.Rewrite(doc.Save(), DocumentPart, _ => Encoding.UTF8.GetBytes(new XElement(ns + "FixedDocument", PageRef(one.PartName)).ToString()));
        bytes = XpsDocumentTests.Rewrite(bytes, SequencePart, data => Edit(data, xml => xml.Add(new XElement(ns + "DocumentReference", new XAttribute("Source", "/Documents/2/FixedDocument.fdoc")))));
        var loaded = XpsDocument.Load(bytes);
        string? Href() => (string?)XElement.Parse(loaded.Pages[0].ToSvg().Svg).Descendants(Svg + "a").Single().Attribute("href");
        void Navigate(string target) { var xml = loaded.Pages[0].GetMarkup(); xml.Elements().Single().SetAttributeValue("FixedPage.NavigateUri", target); loaded.Pages[0].ReplaceMarkup(xml); }
        Assert.Equal("page-2.svg#xps-same", Href());
        Navigate("/FixedDocumentSequence.fdseq#same"); Assert.Equal("#xps-same", Href());
        Navigate("/Documents/2/FixedDocument.fdoc#absent"); Assert.Equal("page-2.svg", Href());
        Navigate("/Documents/2/FixedDocument.fdoc#unknown"); Assert.Equal("page-2.svg", Href());
        Navigate("/FixedDocumentSequence.fdseq#3"); Assert.Equal("page-3.svg", Href());
        Assert.Equal("xps-same", (string?)XElement.Parse(loaded.Pages[1].ToSvg().Svg).Attribute("id"));
    }

    [Fact]
    public void RepeatedGradientsAreBoundedByExpandedNodeBudget() {
        var page = XpsDocument.Create().AddPage(100, 100); var xml = page.GetMarkup(); XNamespace ns = xml.Name.Namespace;
        var brush = new XElement(ns + "LinearGradientBrush", new XAttribute(Key + "Key", "paint"), new XElement(ns + "LinearGradientBrush.GradientStops",
            Enumerable.Range(0, 512).Select(i => new XElement(ns + "GradientStop", new XAttribute("Offset", i / 511D), new XAttribute("Color", "#FF000000")))));
        xml.Add(new XElement(ns + "FixedPage.Resources", new XElement(ns + "ResourceDictionary", brush)));
        for (int i = 0; i < 256; i++) xml.Add(new XElement(ns + "Path", new XAttribute("Data", "M0,0L10,10"), new XAttribute("Fill", "{StaticResource paint}")));
        page.ReplaceMarkup(xml);
        Assert.Contains("budget", Assert.Throws<InvalidDataException>(() => page.ToSvg(true)).Message);
    }

    [Fact]
    public void RepeatedLongGeometryIsBoundedEvenWithFewCommands() {
        var page = XpsDocument.Create().AddPage(100, 100); var xml = page.GetMarkup(); XNamespace ns = xml.Name.Namespace;
        var geometry = new XElement(ns + "PathGeometry", new XAttribute(Key + "Key", "geometry"), new XAttribute("Figures", "M0,0" + new string(' ', 256 * 1024) + "L10,10"));
        xml.Add(new XElement(ns + "FixedPage.Resources", new XElement(ns + "ResourceDictionary", geometry)));
        for (int i = 0; i < 160; i++) xml.Add(new XElement(ns + "Path", new XAttribute("Data", "{StaticResource geometry}"), new XAttribute("Fill", "#FF000000")));
        page.ReplaceMarkup(xml);
        Assert.Contains("output budget", Assert.Throws<InvalidDataException>(() => page.ToSvg(true)).Message);
    }

    [Fact]
    public void NativeMiterDefaultIsRetainedAndClippedMitersAreDiagnosed() {
        var page = XpsDocument.Create().AddPage(100, 100).AddPath("M0,90L50,0L70,90", null, "#FF000000", 3);
        Assert.Contains("stroke-miterlimit=\"10\"", page.ToSvg().Svg);
        var xml = page.GetMarkup(); xml.Elements().Single().SetAttributeValue("StrokeMiterLimit", "2"); page.ReplaceMarkup(xml);
        Assert.Contains("Clipped miter stroke join", page.ToSvg(true).Diagnostics);
        Assert.Throws<NotSupportedException>(() => page.ToSvg());
        xml.Elements().Single().SetAttributeValue("StrokeLineJoin", "Bevel"); page.ReplaceMarkup(xml);
        Assert.True(page.ToSvg().IsComplete);
    }

    [Theory]
    [InlineData("M10,10 L50,10 L50,10 L50,50")]
    [InlineData("M10,10 L50,10 C50,10 50,10 50,10 L50,50")]
    [InlineData("M10,10 L10,10 L50,10 L50,50 Z")]
    public void DegenerateSegmentsUseTheNativeImpliedMiterLimit(string geometry) {
        var page = XpsDocument.Create().AddPage(100, 100).AddPath(geometry, null, "#FF000000", 4);
        Assert.Contains("Clipped miter stroke join", page.ToSvg(true).Diagnostics);
        Assert.Throws<NotSupportedException>(() => page.ToSvg());
        var xml = page.GetMarkup(); xml.Elements().Single().SetAttributeValue("StrokeLineJoin", "Round"); page.ReplaceMarkup(xml);
        Assert.True(page.ToSvg().IsComplete);
    }

    [Fact]
    public void DefaultFlatDashCapsDoNotSilentlyBecomeRound() {
        var page = XpsDocument.Create().AddPage(100, 100).AddPath("M10,20L90,20", null, "#FF000000", 3);
        var xml = page.GetMarkup(); var path = xml.Elements().Single();
        path.SetAttributeValue("StrokeStartLineCap", "Round"); path.SetAttributeValue("StrokeEndLineCap", "Round"); path.SetAttributeValue("StrokeDashArray", "2,1"); page.ReplaceMarkup(xml);
        Assert.Contains("Separate stroke dash caps", page.ToSvg(true).Diagnostics);
        Assert.Throws<NotSupportedException>(() => page.ToSvg());
        path.SetAttributeValue("StrokeDashCap", "Round"); page.ReplaceMarkup(xml);
        Assert.True(page.ToSvg().IsComplete);
    }

    private static byte[] Edit(byte[] data, Action<XElement> edit) { var xml = XElement.Parse(Encoding.UTF8.GetString(data)); edit(xml); return Encoding.UTF8.GetBytes(xml.ToString()); }
}
