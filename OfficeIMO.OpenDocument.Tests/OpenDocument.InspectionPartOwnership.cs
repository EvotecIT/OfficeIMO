using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.OpenDocument.Testing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentInspectionPartOwnershipTests {
    [Theory]
    [InlineData("paragraph")]
    [InlineData("table")]
    [InlineData("inline")]
    public void AuthoredHeaderBlockEditsPersistInStylesPart(string route) {
        OdtDocument created = OdtDocument.Create(); created.AddParagraph("Unchanged body"); created.PageLayout.Header.AddParagraph("Old header");
        byte[] source = RewriteStyles(created.ToBytes(), xml => xml.Descendants(OdfNamespaces.Style + "header").Single().Add(Table()));
        using var input = new MemoryStream(source); OdtDocument document = OdtDocument.Load(input);
        OdtHeaderFooter header = Assert.Single(document.HeaderFooters);
        if (route == "paragraph") header.ContentBlocks.Single(block => block.Paragraph != null).Paragraph!.Text = "Edited header";
        else if (route == "table") header.ContentBlocks.Single(block => block.Table != null).Table!.Cell(1, 1).Text = "Edited table";
        else header.ContentBlocks.Single(block => block.Table != null).Table!.Cell(1, 1).Paragraphs[0].Spans[0].Text = "Edited inline";
        using var saved = new MemoryStream(document.ToBytes()); OdtDocument restored = OdtDocument.Load(saved);
        OdtHeaderFooter restoredHeader = Assert.Single(restored.HeaderFooters);
        if (route == "paragraph") Assert.Equal("Edited header", restoredHeader.ContentBlocks.Single(block => block.Paragraph != null).Paragraph!.Text);
        else Assert.Contains(route == "table" ? "Edited table" : "Edited inline", restoredHeader.ContentBlocks.Single(block => block.Table != null).Table!.Cell(1, 1).Text);
        Assert.Equal("Unchanged body", Assert.Single(restored.Paragraphs).Text);
    }

    [Theory]
    [InlineData("paragraph")]
    [InlineData("table")]
    [InlineData("inline")]
    public void AuthoredMasterShapeTextEditsPersistInStylesPart(string route) {
        OdpPresentation created = OdpPresentation.Create(); created.AddSlide("Unchanged slide");
        byte[] source = RewriteStyles(created.ToBytes(), xml => {
            XElement master = xml.Descendants(OdfNamespaces.Style + "master-page").First();
            master.Add(new XElement(OdfNamespaces.Draw + "frame", new XAttribute(OdfNamespaces.Draw + "name", "MasterText"),
                new XElement(OdfNamespaces.Draw + "text-box", new XElement(OdfNamespaces.Text + "p", "Old master",
                    new XElement(OdfNamespaces.Text + "span", "Old inline")))));
            master.Add(new XElement(OdfNamespaces.Draw + "frame", new XAttribute(OdfNamespaces.Draw + "name", "MasterTable"), Table()));
        });
        using var input = new MemoryStream(source); OdpPresentation document = OdpPresentation.Load(input);
        OdpMasterPage master = document.MasterPages.First();
        if (route == "paragraph") master.Shapes.Single(shape => shape.Name == "MasterText").TextParagraphs[0].Text = "Edited master";
        else if (route == "table") master.Shapes.OfType<OdpTable>().Single().Cell(1, 1).Text = "Edited table";
        else master.Shapes.Single(shape => shape.Name == "MasterText").TextParagraphs[0].Runs[0].Text = "Edited inline";
        using var saved = new MemoryStream(document.ToBytes()); OdpPresentation restored = OdpPresentation.Load(saved);
        OdpMasterPage restoredMaster = restored.MasterPages.First();
        if (route == "paragraph") Assert.Equal("Edited master", restoredMaster.Shapes.Single(shape => shape.Name == "MasterText").TextParagraphs[0].Text);
        else if (route == "table") Assert.Equal("Edited table", restoredMaster.Shapes.OfType<OdpTable>().Single().Cell(1, 1).Text);
        else Assert.Equal("Edited inline", restoredMaster.Shapes.Single(shape => shape.Name == "MasterText").TextParagraphs[0].Runs[0].Text);
        Assert.Equal("Unchanged slide", Assert.Single(restored.Slides).Name);
    }

    [Fact]
    public void MasterParagraphRunLinkAndListStylesStayInTheirOwningPart() {
        OdpPresentation created = OdpPresentation.Create(); created.AddSlide("Unchanged slide");
        byte[] source = RewriteStyles(created.ToBytes(), xml => xml.Descendants(OdfNamespaces.Style + "master-page").First().Add(
            new XElement(OdfNamespaces.Draw + "frame", new XAttribute(OdfNamespaces.Draw + "name", "MasterText"),
                new XElement(OdfNamespaces.Draw + "text-box", new XElement(OdfNamespaces.Text + "p", "Source")))));
        using var input = new MemoryStream(source); OdpPresentation document = OdpPresentation.Load(input);
        byte[] unchangedContent = document.GetPackageEntryBytes("content.xml");
        OdpTextBox text = document.MasterPages.First().Shapes.OfType<OdpTextBox>().Single();
        OdpParagraph paragraph = text.TextParagraphs[0]; paragraph.Bold = true;
        paragraph.AddRun("Styled run").Italic = true;
        OdpHyperlink link = paragraph.AddHyperlink("Target", "https://example.invalid/old"); link.Bold = true;
        link.Href = "https://example.invalid/edited";
        text.AddList(ordered: true).AddItem("Master item").Bold = true;
        using var saved = new MemoryStream(document.ToBytes()); OdpPresentation restored = OdpPresentation.Load(saved);
        OdpTextBox restoredText = restored.MasterPages.First().Shapes.OfType<OdpTextBox>().Single();
        Assert.True(restoredText.TextParagraphs[0].Bold);
        Assert.True(restoredText.TextParagraphs[0].Runs[0].Italic);
        OdpHyperlink restoredLink = restoredText.TextParagraphs[0].InlineNodes.Single(node => node.Hyperlink != null).Hyperlink!;
        Assert.Equal("https://example.invalid/edited", restoredLink.Href); Assert.True(restoredLink.Bold);
        Assert.Equal("Master item", Assert.Single(restoredText.Lists).Items[0].Text); Assert.True(restoredText.Lists[0].Items[0].Bold);
        Assert.Equal(unchangedContent, restored.GetPackageEntryBytes("content.xml"));
    }

    [Fact]
    public void DuplicateAutomaticStyleNamesResolveWithinEachTextPartsOwnership() {
        OdpPresentation created = OdpPresentation.Create(); created.AddSlide("Source slide");
        byte[] source = OdfTestPackageRewriter.Rewrite(created.ToBytes(), (name, bytes) => {
            if (name != "styles.xml" && name != "content.xml") return bytes;
            using var input = new MemoryStream(bytes); XDocument xml = XDocument.Load(input);
            XElement root = xml.Root!;
            XElement styles = root.Element(OdfNamespaces.Office + "automatic-styles")!;
            styles.Add(BackgroundStyle("P1", "paragraph", name == "styles.xml" ? "#224466" : "#AABBCC"));
            styles.Add(BackgroundStyle("T1", "text", name == "styles.xml" ? "#446688" : "#CCBBAA"));
            XElement textFrame = new(OdfNamespaces.Draw + "frame", new XAttribute(OdfNamespaces.Draw + "name", "StyledText"),
                new XElement(OdfNamespaces.Draw + "text-box", StyledParagraph(),
                    new XElement(OdfNamespaces.Text + "list", new XElement(OdfNamespaces.Text + "list-item", StyledParagraph()))));
            XElement owner = name == "styles.xml"
                ? root.Descendants(OdfNamespaces.Style + "master-page").First()
                : root.Descendants(OdfNamespaces.Draw + "page").First();
            owner.Add(textFrame);
            owner.Add(new XElement(OdfNamespaces.Draw + "frame", new XAttribute(OdfNamespaces.Draw + "name", "StyledTable"),
                new XElement(OdfNamespaces.Table + "table", new XElement(OdfNamespaces.Table + "table-row",
                    new XElement(OdfNamespaces.Table + "table-cell", StyledParagraph())))));
            using var output = new MemoryStream(); xml.Save(output); return output.ToArray();
        });
        using var input = new MemoryStream(source); OdpPresentation document = OdpPresentation.Load(input);
        AssertBackgrounds(document.MasterPages.First().Shapes, "#224466", "#446688");
        AssertBackgrounds(document.Slides[0].Shapes, "#AABBCC", "#CCBBAA");
        byte[] unchangedContent = document.GetPackageEntryBytes("content.xml");
        OdpParagraph masterParagraph = document.MasterPages.First().Shapes.OfType<OdpTextBox>().Single().TextParagraphs[0];
        masterParagraph.BackgroundColor = OdfColor.Parse("#6688AA");
        using var saved = new MemoryStream(document.ToBytes()); OdpPresentation restored = OdpPresentation.Load(saved);
        Assert.Equal(OdfColor.Parse("#6688AA"), restored.MasterPages.First().Shapes.OfType<OdpTextBox>().Single().TextParagraphs[0].BackgroundColor);
        AssertBackgrounds(restored.Slides[0].Shapes, "#AABBCC", "#CCBBAA");
        Assert.Equal(unchangedContent, restored.GetPackageEntryBytes("content.xml"));
    }

    private static XElement BackgroundStyle(string name, string family, string color) => new(OdfNamespaces.Style + "style",
        new XAttribute(OdfNamespaces.Style + "name", name), new XAttribute(OdfNamespaces.Style + "family", family),
        new XElement(OdfNamespaces.Style + "text-properties", new XAttribute(OdfNamespaces.Fo + "background-color", color)));

    private static XElement StyledParagraph() => new(OdfNamespaces.Text + "p", new XAttribute(OdfNamespaces.Text + "style-name", "P1"),
        new XElement(OdfNamespaces.Text + "span", new XAttribute(OdfNamespaces.Text + "style-name", "T1"), "Run"),
        new XElement(OdfNamespaces.Text + "a", new XAttribute(OdfNamespaces.Text + "style-name", "T1"),
            new XAttribute(OdfNamespaces.XLink + "href", "https://example.invalid/target"), "Link"));

    private static void AssertBackgrounds(System.Collections.Generic.IReadOnlyList<OdpShape> shapes, string paragraphColor, string inlineColor) {
        OdpTextBox text = shapes.OfType<OdpTextBox>().Single();
        OdpParagraph[] paragraphs = { text.TextParagraphs[0], text.Lists[0].Items[0], shapes.OfType<OdpTable>().Single().Cell(0, 0).Paragraphs[0] };
        foreach (OdpParagraph paragraph in paragraphs) {
            Assert.Equal(OdfColor.Parse(paragraphColor), paragraph.BackgroundColor);
            Assert.Equal(OdfColor.Parse(inlineColor), Assert.Single(paragraph.Runs).BackgroundColor);
            Assert.Equal(OdfColor.Parse(inlineColor), paragraph.InlineNodes.Single(node => node.Hyperlink != null).Hyperlink!.BackgroundColor);
        }
    }

    private static XElement Table() => new(OdfNamespaces.Table + "table", new XAttribute(OdfNamespaces.Table + "name", "SourceTable"),
        new XElement(OdfNamespaces.Table + "table-row", new XAttribute(OdfNamespaces.Table + "number-rows-repeated", 2),
            new XElement(OdfNamespaces.Table + "table-cell", new XAttribute(OdfNamespaces.Table + "number-columns-repeated", 2),
                new XElement(OdfNamespaces.Text + "p", new XElement(OdfNamespaces.Text + "span", "Old cell")))));

    private static byte[] RewriteStyles(byte[] package, Action<XDocument> mutate) => OdfTestPackageRewriter.Rewrite(package,
        (name, bytes) => {
            if (name != "styles.xml") return bytes;
            using var input = new MemoryStream(bytes); XDocument xml = XDocument.Load(input); mutate(xml);
            using var output = new MemoryStream(); xml.Save(output); return output.ToArray();
        });
}
