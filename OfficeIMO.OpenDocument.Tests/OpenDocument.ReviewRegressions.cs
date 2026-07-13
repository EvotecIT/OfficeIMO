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
        OdtDocument document = OdtDocument.Create();
        document.AddParagraph("placeholder");
        XDocument content = document.Package.GetXml("content.xml");
        XElement paragraph = content.Descendants(OdfNamespaces.Text + "p").Single();
        paragraph.ReplaceNodes(new XElement(OdfNamespaces.Text + "s",
            new XAttribute(OdfNamespaces.Text + "c", int.MaxValue)));
        document.Package.MarkXmlDirty("content.xml");

        OdtDocument reopened = OdtDocument.Load(new MemoryStream(document.ToBytes()));

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => reopened.Paragraphs.Single().Text);
        Assert.Contains("safety limit", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ExistingHeaderParagraphEditsRewriteStylesPart() {
        OdtDocument source = OdtDocument.Create();
        source.PageLayout.Header.AddParagraph("Before");
        byte[] original = source.ToBytes();

        OdtDocument edited = OdtDocument.Load(new MemoryStream(original));
        edited.PageLayout.Header.Paragraphs.Single().Text = "After";
        OdfSaveResult save = edited.Serialize();
        byte[] output = save.Value;

        Assert.Contains("styles.xml", save.Report.RewrittenEntries);
        OdtDocument reopened = OdtDocument.Load(new MemoryStream(output));
        Assert.Equal("After", reopened.PageLayout.Header.Paragraphs.Single().Text);
    }

    [Fact]
    public void ExistingHeaderImageEditsRewriteStylesPart() {
        OdtDocument source = OdtDocument.Create();
        source.PageLayout.Header.AddParagraph().AddImage(TinyPng, "header.png",
            OdfLength.Centimeters(1), OdfLength.Centimeters(1));

        OdtDocument edited = OdtDocument.Load(new MemoryStream(source.ToBytes()));
        edited.PageLayout.Header.Paragraphs.Single().Images.Single().Width = OdfLength.Centimeters(2);
        OdfSaveResult save = edited.Serialize();
        byte[] output = save.Value;

        Assert.Contains("styles.xml", save.Report.RewrittenEntries);
        OdtDocument reopened = OdtDocument.Load(new MemoryStream(output));
        Assert.Equal(OdfLength.Centimeters(2), reopened.PageLayout.Header.Paragraphs.Single().Images.Single().Width);
    }

    [Fact]
    public void DirectFormattingClonesSharedAutomaticStyles() {
        OdtDocument document = OdtDocument.Create();
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
        OdtDocument reopened = OdtDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.True(reopened.Paragraphs[0].Bold);
        Assert.False(reopened.Paragraphs[1].Bold);
    }

    [Fact]
    public void ManifestValidationIgnoresUnlistedZipDirectoryEntries() {
        OdtDocument document = OdtDocument.Create();
        document.Package.AddOrReplaceEntry("Configurations2/accelerator/", Array.Empty<byte>(), string.Empty);

        OdfValidationResult result = document.Validate();

        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Id == "ODF103");
    }

    [Fact]
    public void ImageBytesResolveRelativeAndEscapedPackageHrefs() {
        OdtDocument document = OdtDocument.Create();
        OdtImage image = document.AddParagraph("Image").AddImage(TinyPng, "pixel.png",
            OdfLength.Centimeters(1), OdfLength.Centimeters(1));
        XElement imageElement = document.Package.GetXml("content.xml").Descendants(OdfNamespaces.Draw + "image").Single();
        imageElement.SetAttributeValue(OdfNamespaces.XLink + "href", "./" + image.Path.Replace(".", "%2E"));
        document.Package.MarkXmlDirty("content.xml");

        OdtDocument reopened = OdtDocument.Load(new MemoryStream(document.ToBytes()));

        Assert.Equal(TinyPng, reopened.Paragraphs.Single().Images.Single().GetImageBytes());
        Assert.True(reopened.Validate().IsValid);
    }

    [Fact]
    public void HeaderStylesResolveWithinStylesPartWhenNamesCollide() {
        OdtDocument document = OdtDocument.Create();
        OdtParagraph body = document.AddParagraph("Body");
        OdtParagraph header = document.PageLayout.Header.AddParagraph("Header");
        body.StyleName = "P1";
        header.StyleName = "P1";
        AddParagraphStyle(document.Package.GetXml("content.xml"), "P1", "normal");
        AddParagraphStyle(document.Package.GetXml("styles.xml"), "P1", "bold");
        document.Package.MarkXmlDirty("content.xml");
        document.Package.MarkXmlDirty("styles.xml");

        OdtDocument reopened = OdtDocument.Load(new MemoryStream(document.ToBytes()));

        Assert.False(reopened.Paragraphs.Single().Bold);
        Assert.True(reopened.PageLayout.Header.Paragraphs.Single().Bold);
        Assert.True(reopened.Validate().IsValid);
    }

    [Fact]
    public void StyleEnumerationToleratesMissingOptionalStylesPart() {
        OdtDocument document = OdtDocument.Create();
        document.AddParagraph("Minimal");
        document.Package.RemoveEntry("styles.xml");

        Assert.Empty(document.Styles.Named);
        Assert.Empty(document.Styles.Automatic);
        Assert.Null(document.Styles.Find(OdfStyleFamily.Paragraph, "Missing"));
    }

    [Fact]
    public void PageDimensionsDoNotUseTheCommonMarginAsFallback() {
        OdtDocument document = OdtDocument.Create();
        _ = document.PageLayout;
        XElement properties = document.Package.GetXml("styles.xml")
            .Descendants(OdfNamespaces.Style + "page-layout-properties").Single();
        properties.Attribute(OdfNamespaces.Fo + "page-width")?.Remove();
        properties.Attribute(OdfNamespaces.Fo + "page-height")?.Remove();
        properties.SetAttributeValue(OdfNamespaces.Fo + "margin", "2cm");
        document.Package.MarkXmlDirty("styles.xml");

        Assert.Equal(OdfLength.Centimeters(21), document.PageLayout.Width);
        Assert.Equal(OdfLength.Centimeters(29.7), document.PageLayout.Height);
        Assert.Equal(OdfLength.Centimeters(2), document.PageLayout.MarginLeft);
    }

    [Fact]
    public void OdsHeaderRowsParticipateInTheLogicalRowModel() {
        OdsDocument document = OdsDocument.Create();
        OdsSheet sheet = document.AddSheet("Data");
        sheet.Cell(0, 0).SetString("Header");
        sheet.Cell(1, 0).SetString("Body");
        XElement table = document.Package.GetXml("content.xml").Descendants(OdfNamespaces.Table + "table").Single();
        XElement[] rows = table.Elements(OdfNamespaces.Table + "table-row").ToArray();
        rows[0].Remove();
        rows[1].AddBeforeSelf(new XElement(OdfNamespaces.Table + "table-header-rows", rows[0]));
        document.Package.MarkXmlDirty("content.xml");

        OdsDocument reopened = OdsDocument.Load(new MemoryStream(document.ToBytes()));
        OdsSheet actual = reopened.Sheets.Single();

        Assert.Equal(2, actual.RowRuns.Count);
        Assert.Equal("Header", actual.GetValue(0, 0).ToString());
        Assert.Equal("Body", actual.GetValue(1, 0).ToString());
        Assert.Equal(1, actual.UsedRange!.Value.LastRow);
    }

    [Fact]
    public void RepeatedOdtTableCellsAreLogicalCellsAndSplitWhenEdited() {
        OdtDocument document = OdtDocument.Create();
        OdtTable table = document.AddTable(1, 1, "Repeated");
        table.Cell(0, 0).Text = "Same";
        XElement cell = table.Element.Descendants(OdfNamespaces.Table + "table-cell").Single();
        cell.SetAttributeValue(OdfNamespaces.Table + "number-columns-repeated", 3);
        document.Package.MarkXmlDirty("content.xml");

        OdtDocument reopened = OdtDocument.Load(new MemoryStream(document.ToBytes()));
        OdtTableRow row = reopened.Tables.Single().Rows.Single();

        Assert.Equal(3, row.Cells.Count);
        Assert.Equal(new[] { "Same", "Same", "Same" }, row.Cells.Select(item => item.Text));
        row.Cells[1].Text = "Changed";
        Assert.Equal(new[] { "Same", "Changed", "Same" }, row.Cells.Select(item => item.Text));

        OdtDocument roundTrip = OdtDocument.Load(new MemoryStream(reopened.ToBytes()));
        Assert.Equal(new[] { "Same", "Changed", "Same" }, roundTrip.Tables.Single().Rows.Single().Cells.Select(item => item.Text));
        Assert.True(roundTrip.Validate().IsValid);
    }

    [Fact]
    public void RepeatedCoveredOdtTableCellsAreLogicalCells() {
        OdtDocument document = OdtDocument.Create();
        OdtTable table = document.AddTable(1, 3, "Merged");
        table.Merge(0, 0, 1, 3).Text = "Anchor";
        XElement[] covered = table.Element.Descendants(OdfNamespaces.Table + "covered-table-cell").ToArray();
        covered[0].SetAttributeValue(OdfNamespaces.Table + "number-columns-repeated", 2);
        covered[1].Remove();
        document.Package.MarkXmlDirty("content.xml");

        OdtDocument reopened = OdtDocument.Load(new MemoryStream(document.ToBytes()));
        var cells = reopened.Tables.Single().Rows.Single().Cells;

        Assert.Equal(3, cells.Count);
        Assert.False(cells[0].IsCovered);
        Assert.True(cells[1].IsCovered);
        Assert.True(cells[2].IsCovered);
        Assert.True(reopened.Validate().IsValid);
    }

    [Fact]
    public void OrderedListsResolveCommonStylesFromStylesPart() {
        OdtDocument document = OdtDocument.Create();
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

        OdtDocument reopened = OdtDocument.Load(new MemoryStream(document.ToBytes()));

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
