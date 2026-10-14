using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentOdgTabStopTests {
    [Fact]
    public void NativeProducerExportRetainsStopsStyledFieldsAndLeadingTabs() {
        string path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "Drawing", "libreoffice-tab-stops.fodg");
        var document = OdgDocument.LoadFlatXml(path); var result = document.Pages[0].ToDrawing();
        var paragraphs = result.Value.Elements.OfType<OfficeDrawingRichText>().SelectMany(t => t.Paragraphs).ToArray();
        Assert.Equal(8, paragraphs.Length);
        Assert.Equal(new[] { OfficeTextTabAlignment.Left, OfficeTextTabAlignment.Center, OfficeTextTabAlignment.Right, OfficeTextTabAlignment.Character },
            paragraphs.Take(4).Select(p => Assert.Single(p.TabStops!.Stops).Alignment));
        Assert.Equal(",", Assert.Single(paragraphs[3].TabStops!.Stops).Character);
        Assert.Equal("\t\tA\tB\n\tNext", string.Concat(paragraphs[5].Runs.Select(r => r.Text)));
        Assert.Equal(36, paragraphs[5].TabStops!.DefaultInterval, 2);
        Assert.Contains(paragraphs[6].Runs, r => r.Text == ",45" && r.Bold && Math.Abs(r.FontSize - 20) < .01);
        Assert.DoesNotContain(result.Report.Mappings, m => m.Feature.EndsWith(":tab-stops", StringComparison.Ordinal) || m.Feature.EndsWith(":tab-leaders", StringComparison.Ordinal));
        using var stream = new MemoryStream(); document.Save(stream); stream.Position = 0;
        var reopened = OdgDocument.Load(stream);
        Assert.Equal(document.Pages[0].Shapes[5].Text, reopened.Pages[0].Shapes[5].Text);
    }
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TypedTabEditingRetainsAllAlignmentsAcrossPackageAndFlatRoundTrips(bool flat) {
        var document = OdgDocument.Create(); var shape = document.AddPage().Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 12, 8));
        var p = shape.AddParagraph("A\tB\tC\t123,45");
        p.SetTabStops(new[] { new OdfTabStop(OdfLength.Points(40)), new OdfTabStop(OdfLength.Points(90), "center"),
            new OdfTabStop(OdfLength.Points(140), "right"), new OdfTabStop(OdfLength.Points(200), "char", ",") });
        using var stream = new MemoryStream(); if (flat) document.SaveFlatXml(stream); else document.Save(stream); stream.Position = 0;
        var reopened = flat ? OdgDocument.LoadFlatXml(stream) : OdgDocument.Load(stream);
        string before = reopened.GetXml("content.xml").ToString();
        var result = reopened.Pages[0].ToDrawing(); var text = Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>());
        var tabs = Assert.Single(text.Paragraphs).TabStops!;
        Assert.Equal(new[] { 40D, 90, 140, 200 }, tabs.Stops.Select(s => s.Position));
        Assert.Equal(OfficeTextTabAlignment.Character, tabs.Stops[3].Alignment); Assert.Equal(",", tabs.Stops[3].Character);
        Assert.Equal("A\tB\tC\t123,45", text.PlainText); Assert.Equal(before, reopened.GetXml("content.xml").ToString());
        Assert.DoesNotContain(result.Report.Mappings, m => m.Feature.EndsWith(":tab-stops", StringComparison.Ordinal));
        Assert.Contains("123,45", OfficeDrawingSvgExporter.ToSvg(result.Value));
    }

    [Fact]
    public void EmptyNearTabSetOverridesWholeParentSetAndDocumentDefaultControlsGridAndOrigin() {
        var document = OdgDocument.Create(); var shape = document.AddPage().Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 12, 8));
        var parent = document.Styles.CreateNamed("ParentTabs", OdfStyleFamily.Paragraph);
        parent.Element.Add(new XElement(OdfNamespaces.Style + "paragraph-properties", new XElement(OdfNamespaces.Style + "tab-stops",
            new XElement(OdfNamespaces.Style + "tab-stop", new XAttribute(OdfNamespaces.Style + "position", "70pt")))));
        var p = shape.AddParagraph("\tA"); p.StyleName = parent.Name; p.MarginLeft = OdfLength.Points(20);
        document.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "styles")!.Add(new XElement(OdfNamespaces.Style + "default-style",
            new XAttribute(OdfNamespaces.Style + "family", "paragraph"), new XElement(OdfNamespaces.Style + "paragraph-properties",
                new XAttribute(OdfNamespaces.Style + "tab-stop-distance", "25pt"), new XAttribute(OdfNamespaces.Text + "relative-tab-stop-position", "false"))));
        var inherited = Assert.Single(document.Pages[0].ToDrawing().Value.Elements.OfType<OfficeDrawingRichText>()).Paragraphs[0].TabStops!;
        Assert.Equal(70, Assert.Single(inherited.Stops).Position); Assert.Equal(25, inherited.DefaultInterval); Assert.Equal(-20, inherited.Origin);
        p.SetTabStops(Array.Empty<OdfTabStop>());
        var cleared = Assert.Single(document.Pages[0].ToDrawing().Value.Elements.OfType<OfficeDrawingRichText>()).Paragraphs[0].TabStops!;
        Assert.Empty(cleared.Stops); Assert.Equal(25, cleared.DefaultInterval);
        Assert.Single(parent.Element.Descendants(OdfNamespaces.Style + "tab-stop"));
    }

    [Fact]
    public void UnsupportedLineLeadersAndInvalidStopsAreRetainedAndReportedWhileBodyProjects() {
        var document = OdgDocument.Create(); var shape = document.AddPage().Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 12, 8), "Tabs");
        var p = shape.AddParagraph("A\tB"); p.SetTabStops(new[] { new OdfTabStop(OdfLength.Points(50)) });
        var container = p.EnsureStyle().ParagraphProperties!.Element(OdfNamespaces.Style + "tab-stops")!;
        container.Elements().Single().SetAttributeValue(OdfNamespaces.Style + "leader-style", "unknown-pattern");
        container.Add(new XElement(OdfNamespaces.Style + "tab-stop", new XAttribute(OdfNamespaces.Style + "position", "80%")));
        string before = document.GetXml("content.xml").ToString(); var result = document.Pages[0].ToDrawing();
        Assert.Contains(result.Report.Mappings, m => m.Feature == "shape:Tabs:text:tab-leaders" && m.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Contains(result.Report.Mappings, m => m.Feature == "shape:Tabs:text:tab-stops" && m.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Equal("A\tB", Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>()).PlainText);
        Assert.Equal(before, document.GetXml("content.xml").ToString());
        Assert.Throws<OdfConversionLossException>(() => document.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
        Assert.Throws<ArgumentException>(() => p.SetTabStops(new[] { new OdfTabStop(OdfLength.Points(40)), new OdfTabStop(OdfLength.Points(40)) }));
        Assert.Equal(before, document.GetXml("content.xml").ToString());
    }

    [Fact]
    public void LiteralXmlIndentationCollapsesAndExplicitTabRetainsItsStopSemantics() {
        var document = OdgDocument.Create(); var shape = document.AddPage().Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 12, 8), "XmlSpaces");
        var p = shape.AddParagraph(); p.Element.Add(new XText("\n  "), new XElement(OdfNamespaces.Text + "tab"), new XText("A"));
        var result = document.Pages[0].ToDrawing();
        var text = Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>());
        Assert.Equal("\tA", text.PlainText);
        Assert.DoesNotContain(result.Report.Mappings, m => m.Feature == "shape:XmlSpaces:text:xml-whitespace");
        Assert.False(result.Report.HasSkippedOrUnsupported);
        document.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
    }
}
