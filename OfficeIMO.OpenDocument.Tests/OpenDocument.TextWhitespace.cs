using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentTextWhitespaceTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ParagraphWhitespaceCollapsesAcrossStyledAndHyperlinkBoundariesWithoutEditingXml(bool flat) {
        var document = OdgDocument.Create();
        var page = document.AddPage();
        var shape = page.Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 12, 8), "Whitespace");
        var paragraph = shape.AddParagraph();
        var span = paragraph.AddRun(); span.Bold = true;
        span.Element.Add(new XText(" \n Beta  "), new XElement(OdfNamespaces.Text + "a",
            new XAttribute(OdfNamespaces.XLink + "href", "https://example.com/"), " \t Gamma "), new XText("  "));
        paragraph.Element.AddFirst(new XText("\r\n  Alpha \t "));
        paragraph.Element.Add(new XText("\r\n Delta \n"));
        document.MarkPartDirty("content.xml");
        string before = paragraph.Element.ToString(SaveOptions.DisableFormatting);
        Assert.Equal("Alpha Beta Gamma Delta", paragraph.Text);
        Assert.Equal(paragraph.Text, string.Concat(paragraph.InlineNodes.Select(n => n.Text)));
        Assert.Equal("Beta Gamma ", span.Text);
        Assert.Equal("Gamma ", Assert.Single(paragraph.Hyperlinks).Text);
        var projected = Assert.Single(page.ToDrawing().Value.Elements.OfType<OfficeDrawingRichText>());
        Assert.Equal(paragraph.Text, projected.PlainText);
        Assert.Contains(projected.Runs, r => r.Bold && r.Text.Contains("Beta"));
        Assert.Contains(projected.Runs, r => r.Text == "Gamma ");
        Assert.Contains("https://example.com/", OfficeDrawingSvgExporter.ToSvg(page.ToDrawing().Value));
        Assert.Equal(before, paragraph.Element.ToString(SaveOptions.DisableFormatting));
        using var bytes = new MemoryStream();
        if (flat) document.SaveFlatXml(bytes); else document.Save(bytes);
        bytes.Position = 0;
        var read = flat ? OdgDocument.LoadFlatXml(bytes) : OdgDocument.Load(bytes);
        Assert.Equal(paragraph.Text, read.Pages[0].Shapes[0].Text);
    }

    [Fact]
    public void ExplicitNativeWhitespaceAndNonXmlUnicodeSpacesRemainDistinct() {
        var p = new XElement(OdfNamespaces.Text + "p", "\n  ",
            new XElement(OdfNamespaces.Text + "s", new XAttribute(OdfNamespaces.Text + "c", 2)), " \n ",
            new XElement(OdfNamespaces.Text + "tab"), "\n ", new XElement(OdfNamespaces.Text + "span", " A  B "),
            new XElement(OdfNamespaces.Text + "line-break"),
            new XElement(OdfNamespaces.Text + "s", new XAttribute(OdfNamespaces.Text + "c", 3)), " \n ");
        Assert.Equal("   \t A B \n   ", OdfTextCodec.Read(p));
        Assert.Equal("A\u00a0\u00a0B\u2003\u2003C\u202f", OdfTextCodec.Read(new XElement(OdfNamespaces.Text + "p", " A\u00a0\u00a0B\u2003\u2003C\u202f ")));
    }

    [Fact]
    public void NativeInlineSnapshotsAgreeAcrossDocumentFamilies() {
        var odt = OdtDocument.Create(); var text = odt.AddParagraph();
        text.Element.Add("\n One  ", new XElement(OdfNamespaces.Text + "span", " Two \t "), " Three \n");
        Assert.Equal("One Two Three", text.Text);
        Assert.Equal(text.Text, string.Concat(text.InlineNodes.Select(n => n.Text)));
        var odp = OdpPresentation.Create(); var slide = odp.AddSlide();
        var paragraph = slide.AddTextBox(OdfRect.FromCentimeters(1, 1, 10, 3)).AddParagraph();
        var element = odp.Package.GetXml("content.xml").Descendants(OdfNamespaces.Text + "p").Single();
        element.Add("\n One  ", new XElement(OdfNamespaces.Text + "span", " Two \t "), " Three \n");
        Assert.Equal(text.Text, paragraph.Text);
        Assert.Equal(paragraph.Text, string.Concat(paragraph.InlineNodes.Select(n => n.Text)));
        var ods = OdsDocument.Create(); var cell = ods.AddSheet("Data").Cell(0, 0); cell.SetString("Seed");
        cell.Element.Element(OdfNamespaces.Text + "p")!.ReplaceNodes(element.Nodes().Select(n => n is XElement e ? (XNode)new XElement(e) : new XText(((XText)n).Value)));
        Assert.Equal(text.Text, cell.Text);
    }

    [Fact]
    public void CollapsedSourceCharactersStillConsumeTheSharedBudget() {
        int budget = 8;
        var p = new XElement(OdfNamespaces.Text + "p", " A   B ");
        Assert.Equal("A B", OdfTextCodec.Read(p, ref budget));
        Assert.Equal(1, budget);
        int small = 8;
        Assert.Throws<InvalidDataException>(() => OdfTextCodec.Read(new XElement(OdfNamespaces.Text + "p", new string(' ', 9)), ref small));
        var document = OdgDocument.Create(); var page = document.AddPage();
        var shape = page.Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 12, 8), "Budget");
        shape.AddParagraph().Element.Add(new string(' ', 100_000), "B");
        Assert.Equal("B", shape.Text);
        var result = page.ToDrawing();
        Assert.DoesNotContain(result.Value.Elements, e => e is OfficeDrawingRichText);
        Assert.Contains(result.Report.Mappings, m => m.Feature == "shape:Budget:text" && m.Status == OdfConversionMappingStatus.Skipped);
    }

    [Fact]
    public void CaseTransformsKeepRawWhitespaceOffsetsAndOtherStories() {
        var document = OdgDocument.Create(); var shape = document.AddPage().Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 12, 8));
        var paragraph = shape.AddParagraph(); paragraph.Element.Add(" \n mixed  ", new XElement(OdfNamespaces.Text + "span", " tail \n"));
        var note = new XElement(OdfNamespaces.Office + "annotation", new XElement(OdfNamespaces.Text + "p", "Do Not Edit"));
        paragraph.Element.Add(note); string before = note.ToString();
        paragraph.TransformTextCase(OfficeTextCase.Uppercase);
        Assert.Equal("MIXED TAIL", paragraph.Text);
        Assert.StartsWith(" \n MIXED  ", ((XText)paragraph.Element.FirstNode!).Value);
        Assert.Equal(before, note.ToString());
    }

    [Fact]
    public void RubyAnnotationsStayOutsideTheVisibleBaseTextAndCaseTransform() {
        var p = new XElement(OdfNamespaces.Text + "p", new XElement(OdfNamespaces.Text + "ruby",
            new XElement(OdfNamespaces.Text + "ruby-base", new XElement(OdfNamespaces.Text + "span", " base  text ")),
            new XElement(OdfNamespaces.Text + "ruby-text", "Annotation")));
        Assert.Equal("base text", OdfTextCodec.Read(p));
        OdfTextCodec.TransformTextCase(p, OfficeTextCase.Uppercase);
        Assert.Equal("BASE TEXT", OdfTextCodec.Read(p));
        Assert.Equal("Annotation", p.Descendants(OdfNamespaces.Text + "ruby-text").Single().Value);
    }

    [Fact]
    public void CachedFieldCharactersRemainOpaqueWhileSurroundingLiteralSpacesCollapse() {
        var field = new XElement(OdfNamespaces.Text + "date", "  cached  ");
        var p = new XElement(OdfNamespaces.Text + "p", " A ", field, " B ");
        Assert.Equal("A   cached   B", OdfTextCodec.Read(p));
        Assert.Equal("  cached  ", OdfTextCodec.Read(field));
        Assert.Equal("  cached  ", OdfTextCodec.Snapshot(p).ReadNode(field));
    }
}
