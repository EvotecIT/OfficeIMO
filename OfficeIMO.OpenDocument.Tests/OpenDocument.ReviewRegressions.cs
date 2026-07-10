using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentReviewRegressionTests {
    private static readonly byte[] TinyPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");

    [Fact]
    public void RejectsDecodedSpaceRunsBeyondTheTextSafetyLimit() {
        using OdtDocument document = OdtDocument.Create();
        document.AddParagraph("placeholder");
        XDocument content = document.Package.GetXml("content.xml");
        XElement paragraph = content.Descendants(OdfNamespaces.Text + "p").Single();
        paragraph.ReplaceNodes(new XElement(OdfNamespaces.Text + "s",
            new XAttribute(OdfNamespaces.Text + "c", int.MaxValue)));
        document.Package.MarkXmlDirty("content.xml");

        using OdtDocument reopened = OdtDocument.Open(new MemoryStream(document.ToBytes()));

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => reopened.Paragraphs.Single().Text);
        Assert.Contains("safety limit", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ExistingHeaderParagraphEditsRewriteStylesPart() {
        using OdtDocument source = OdtDocument.Create();
        source.PageLayout.Header.AddParagraph("Before");
        byte[] original = source.ToBytes();

        using OdtDocument edited = OdtDocument.Open(new MemoryStream(original));
        edited.PageLayout.Header.Paragraphs.Single().Text = "After";
        byte[] output = edited.ToBytes();

        Assert.Contains("styles.xml", edited.LastSaveReport!.RewrittenEntries);
        using OdtDocument reopened = OdtDocument.Open(new MemoryStream(output));
        Assert.Equal("After", reopened.PageLayout.Header.Paragraphs.Single().Text);
    }

    [Fact]
    public void DirectFormattingClonesSharedAutomaticStyles() {
        using OdtDocument document = OdtDocument.Create();
        OdtParagraph first = document.AddParagraph("First");
        OdtParagraph second = document.AddParagraph("Second");
        OdfStyle shared = document.Styles.CreateAutomatic(OdfStyleFamily.Paragraph, "shared");
        shared.Bold = false;
        first.StyleName = shared.Name;
        second.StyleName = shared.Name;

        first.Bold = true;

        Assert.True(first.Bold);
        Assert.False(second.Bold);
        Assert.NotEqual(first.StyleName, second.StyleName);
        using OdtDocument reopened = OdtDocument.Open(new MemoryStream(document.ToBytes()));
        Assert.True(reopened.Paragraphs[0].Bold);
        Assert.False(reopened.Paragraphs[1].Bold);
    }

    [Fact]
    public void ManifestValidationIgnoresUnlistedZipDirectoryEntries() {
        using OdtDocument document = OdtDocument.Create();
        document.Package.AddOrReplaceEntry("Configurations2/accelerator/", Array.Empty<byte>(), string.Empty);

        OdfValidationResult result = document.Validate();

        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Id == "ODF103");
    }

    [Fact]
    public void ImageBytesResolveRelativeAndEscapedPackageHrefs() {
        using OdtDocument document = OdtDocument.Create();
        OdtImage image = document.AddParagraph("Image").AddImage(TinyPng, "pixel.png",
            OdfLength.Centimeters(1), OdfLength.Centimeters(1));
        XElement imageElement = document.Package.GetXml("content.xml").Descendants(OdfNamespaces.Draw + "image").Single();
        imageElement.SetAttributeValue(OdfNamespaces.XLink + "href", "./" + image.Path.Replace(".", "%2E"));
        document.Package.MarkXmlDirty("content.xml");

        using OdtDocument reopened = OdtDocument.Open(new MemoryStream(document.ToBytes()));

        Assert.Equal(TinyPng, reopened.Paragraphs.Single().Images.Single().GetImageBytes());
        Assert.True(reopened.Validate().IsValid);
    }

    [Fact]
    public void HeaderStylesResolveWithinStylesPartWhenNamesCollide() {
        using OdtDocument document = OdtDocument.Create();
        OdtParagraph body = document.AddParagraph("Body");
        OdtParagraph header = document.PageLayout.Header.AddParagraph("Header");
        body.StyleName = "P1";
        header.StyleName = "P1";
        AddParagraphStyle(document.Package.GetXml("content.xml"), "P1", "normal");
        AddParagraphStyle(document.Package.GetXml("styles.xml"), "P1", "bold");
        document.Package.MarkXmlDirty("content.xml");
        document.Package.MarkXmlDirty("styles.xml");

        using OdtDocument reopened = OdtDocument.Open(new MemoryStream(document.ToBytes()));

        Assert.False(reopened.Paragraphs.Single().Bold);
        Assert.True(reopened.PageLayout.Header.Paragraphs.Single().Bold);
        Assert.True(reopened.Validate().IsValid);
    }

    [Fact]
    public void RepeatedOdtTableCellsAreLogicalCellsAndSplitWhenEdited() {
        using OdtDocument document = OdtDocument.Create();
        OdtTable table = document.AddTable(1, 1, "Repeated");
        table.Cell(0, 0).Text = "Same";
        XElement cell = table.Element.Descendants(OdfNamespaces.Table + "table-cell").Single();
        cell.SetAttributeValue(OdfNamespaces.Table + "number-columns-repeated", 3);
        document.Package.MarkXmlDirty("content.xml");

        using OdtDocument reopened = OdtDocument.Open(new MemoryStream(document.ToBytes()));
        OdtTableRow row = reopened.Tables.Single().Rows.Single();

        Assert.Equal(3, row.Cells.Count);
        Assert.Equal(new[] { "Same", "Same", "Same" }, row.Cells.Select(item => item.Text));
        row.Cells[1].Text = "Changed";
        Assert.Equal(new[] { "Same", "Changed", "Same" }, row.Cells.Select(item => item.Text));

        using OdtDocument roundTrip = OdtDocument.Open(new MemoryStream(reopened.ToBytes()));
        Assert.Equal(new[] { "Same", "Changed", "Same" }, roundTrip.Tables.Single().Rows.Single().Cells.Select(item => item.Text));
        Assert.True(roundTrip.Validate().IsValid);
    }

    [Fact]
    public void RepeatedCoveredOdtTableCellsAreLogicalCells() {
        using OdtDocument document = OdtDocument.Create();
        OdtTable table = document.AddTable(1, 3, "Merged");
        table.Merge(0, 0, 1, 3).Text = "Anchor";
        XElement[] covered = table.Element.Descendants(OdfNamespaces.Table + "covered-table-cell").ToArray();
        covered[0].SetAttributeValue(OdfNamespaces.Table + "number-columns-repeated", 2);
        covered[1].Remove();
        document.Package.MarkXmlDirty("content.xml");

        using OdtDocument reopened = OdtDocument.Open(new MemoryStream(document.ToBytes()));
        var cells = reopened.Tables.Single().Rows.Single().Cells;

        Assert.Equal(3, cells.Count);
        Assert.False(cells[0].IsCovered);
        Assert.True(cells[1].IsCovered);
        Assert.True(cells[2].IsCovered);
        Assert.True(reopened.Validate().IsValid);
    }

    [Fact]
    public void OrderedListsResolveCommonStylesFromStylesPart() {
        using OdtDocument document = OdtDocument.Create();
        document.AddList(ordered: true).AddItem("Numbered");
        XDocument content = document.Package.GetXml("content.xml");
        XDocument styles = document.Package.GetXml("styles.xml");
        XElement list = content.Descendants(OdfNamespaces.Text + "list").Single();
        string styleName = (string)list.Attribute(OdfNamespaces.Text + "style-name")!;
        XElement listStyle = content.Root!.Element(OdfNamespaces.Office + "automatic-styles")!
            .Elements(OdfNamespaces.Text + "list-style").Single();
        listStyle.Remove();
        styles.Root!.Element(OdfNamespaces.Office + "styles")!.Add(listStyle);
        document.Package.MarkXmlDirty("content.xml");
        document.Package.MarkXmlDirty("styles.xml");

        using OdtDocument reopened = OdtDocument.Open(new MemoryStream(document.ToBytes()));

        OdtContentBlock item = reopened.ContentBlocks.Single(block => block.IsListItem);
        Assert.Equal(styleName, (string?)reopened.TextBody.Descendants(OdfNamespaces.Text + "list").Single()
            .Attribute(OdfNamespaces.Text + "style-name"));
        Assert.True(item.IsOrderedList);
        Assert.True(reopened.Validate().IsValid);
    }

    private static void AddParagraphStyle(XDocument document, string name, string weight) {
        XElement automatic = document.Root!.Element(OdfNamespaces.Office + "automatic-styles")!;
        automatic.Add(new XElement(OdfNamespaces.Style + "style",
            new XAttribute(OdfNamespaces.Style + "name", name),
            new XAttribute(OdfNamespaces.Style + "family", "paragraph"),
            new XElement(OdfNamespaces.Style + "text-properties",
                new XAttribute(OdfNamespaces.Fo + "font-weight", weight))));
    }

}
