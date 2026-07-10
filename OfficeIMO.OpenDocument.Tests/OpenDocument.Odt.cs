using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public class OpenDocumentOdtTests {
    [Theory]
    [InlineData("libreoffice-writer-basic.odt")]
    [InlineData("microsoft-word-basic.odt")]
    public void PreservesAuthoredFixtureOutsideEditedContent(string fixtureName) {
        string path = Path.Combine(AppContext.BaseDirectory, "Fixtures", fixtureName);
        using OdtDocument document = OdtDocument.Open(path);
        var untouched = document.Package.Entries
            .Where(entry => entry.Name != "content.xml" && entry.Name != "META-INF/manifest.xml")
            .ToDictionary(entry => entry.Name, entry => entry.GetOriginalBytes());

        OdtParagraph paragraph = document.Paragraphs.First(item => item.Text.Length > 0);
        paragraph.Text = paragraph.Text + " [OfficeIMO]";
        byte[] output = document.ToBytes(new OdfSaveOptions { CompatibilityProfile = OdfCompatibilityProfile.PreserveSource });

        using OdtDocument reopened = OdtDocument.Open(new MemoryStream(output));
        Assert.Contains(reopened.Paragraphs, item => item.Text.EndsWith("[OfficeIMO]", StringComparison.Ordinal));
        foreach (var pair in untouched) {
            Assert.Equal(pair.Value, reopened.Package.GetRequiredEntry(pair.Key).GetOriginalBytes());
        }
    }

    [Fact]
    public void WritesAndReopensUsefulTextDocumentWithoutFlatteningStructure() {
        using OdtDocument document = OdtDocument.Create();
        document.Metadata.Title = "Native ODT";
        document.AddHeading("Quarterly report", 1);
        OdtParagraph paragraph = document.AddParagraph("Revenue  increased\t12%\nYear over year.");
        OdtSpan span = paragraph.AddSpan(" Important");
        span.Bold = true;
        span.Color = OdfColor.Parse("#B42318");
        paragraph.AddHyperlink("Details", "https://example.test/report");
        paragraph.AddBookmark("summary");

        OdtList list = document.AddList(ordered: true);
        list.AddItem("First");
        list.AddItem("Second");

        OdtTable table = document.AddTable(2, 3, "Results");
        table.Cell(0, 0).Text = "Region";
        table.Cell(0, 1).Text = "Actual";
        table.Cell(0, 2).Text = "Plan";
        table.Cell(1, 0).Text = "EMEA";
        table.Merge(1, 1, 1, 2).Text = "125";

        OdtSection section = document.AddSection("Appendix");
        section.AddHeading("Notes", 2);
        section.AddParagraph("Preserved section content.");
        document.AddPageBreak().AddText("Next page");
        document.PageLayout.MarginLeft = OdfLength.Centimeters(2.5);
        document.PageLayout.Header.AddParagraph("OfficeIMO");
        document.PageLayout.Footer.AddParagraph("Confidential");

        byte[] bytes = document.ToBytes();
        Assert.True(document.Validate().IsValid);
        using OdtDocument reopened = OdtDocument.Open(new MemoryStream(bytes));

        Assert.Equal("Native ODT", reopened.Metadata.Title);
        Assert.Contains(reopened.Paragraphs, item => item.IsHeading && item.Text == "Quarterly report");
        Assert.Contains(reopened.Paragraphs, item => item.Text.Contains("Revenue  increased\t12%\nYear over year.", System.StringComparison.Ordinal));
        Assert.Equal(2, reopened.Lists.Single().Items.Count);
        Assert.Equal("125", reopened.Tables.Single().Cell(1, 1).Text);
        Assert.True(reopened.Tables.Single().Cell(1, 2).IsCovered);
        Assert.Equal(2, reopened.Tables.Single().Cell(1, 1).ColumnSpan);
        Assert.Equal("OfficeIMO", reopened.PageLayout.Header.Paragraphs.Single().Text);
        Assert.Equal("Confidential", reopened.PageLayout.Footer.Paragraphs.Single().Text);
        Assert.Contains(reopened.Styles.Automatic, style => style.Family == OdfStyleFamily.Text && style.Bold == true);
    }

    [Fact]
    public void PreservesUnknownXmlAndEntriesDuringTargetedOdtEdit() {
        using OdtDocument source = OdtDocument.Create();
        source.AddParagraph("Before");
        XNamespace vendor = "urn:vendor:test";
        XElement foreign = new XElement(vendor + "payload", new XAttribute(vendor + "mode", "keep"), "opaque");
        source.TextBody.Add(foreign);
        source.MarkPartDirty("content.xml");
        source.Package.AddOrReplaceEntry("Vendor/data.dat", new byte[] { 4, 2, 1 }, "application/octet-stream");

        using OdtDocument edited = OdtDocument.Open(new MemoryStream(source.ToBytes()));
        edited.Paragraphs.Single().Text = "After";
        byte[] output = edited.ToBytes();

        using OdtDocument reopened = OdtDocument.Open(new MemoryStream(output));
        Assert.Equal("After", reopened.Paragraphs.Single().Text);
        Assert.Equal("opaque", reopened.TextBody.Element(vendor + "payload")?.Value);
        Assert.Equal(new byte[] { 4, 2, 1 }, reopened.Package.GetRequiredEntry("Vendor/data.dat").GetOriginalBytes());
    }

    [Fact]
    public void AddsEmbeddedImageAndManifestMediaType() {
        byte[] png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
        using OdtDocument document = OdtDocument.Create();
        OdtImage image = document.AddParagraph().AddImage(png, "pixel.png", OdfLength.Centimeters(1), OdfLength.Centimeters(1));

        byte[] bytes = document.ToBytes();
        Assert.StartsWith("Pictures/", image.Path);
        using var archive = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        Assert.NotNull(archive.GetEntry(image.Path));
        XDocument manifest;
        using (Stream stream = archive.GetEntry("META-INF/manifest.xml")!.Open()) manifest = XDocument.Load(stream);
        XNamespace ns = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
        XElement entry = manifest.Root!.Elements(ns + "file-entry")
            .Single(item => (string?)item.Attribute(ns + "full-path") == image.Path);
        Assert.Equal("image/png", (string?)entry.Attribute(ns + "media-type"));
    }
}
