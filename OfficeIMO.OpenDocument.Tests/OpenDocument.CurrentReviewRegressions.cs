using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentCurrentReviewRegressionTests {
    [Fact]
    public void TextAndPresentationTablesIncludeHeaderRowsInSourceOrder() {
        using OdtDocument text = OdtDocument.Create();
        OdtTable textTable = text.AddTable(2, 1, "TextTable");
        textTable.Cell(0, 0).Text = "Text header";
        textTable.Cell(1, 0).Text = "Text body";
        WrapFirstRowAsHeader(textTable.Element);

        using OdpPresentation presentation = OdpPresentation.Create();
        OdpTable presentationTable = presentation.AddSlide("Table").AddTable(
            OdfRect.FromCentimeters(1, 1, 8, 4), 2, 1, "PresentationTable");
        presentationTable.Cell(0, 0).Text = "Slide header";
        presentationTable.Cell(1, 0).Text = "Slide body";
        WrapFirstRowAsHeader(presentationTable.Element.Element(OdfNamespaces.Table + "table")!);

        Assert.Equal(new[] { "Text header", "Text body" }, textTable.Rows.Select(row => row.Cells[0].Text));
        Assert.Equal(new[] { "Slide header", "Slide body" }, presentationTable.Rows.Select(row => row.Cells[0].Text));
    }

    [Fact]
    public void NewOdsColumnsStayAtTableScopeWhenHeaderRowsComeFirst() {
        using OdsDocument document = OdsDocument.Create();
        OdsSheet sheet = document.AddSheet("Data");
        sheet.Cell(0, 0).SetString("Header");
        sheet.Cell(1, 0).SetString("Body");
        XElement table = sheet.Element;
        WrapFirstRowAsHeader(table);

        _ = sheet.Column(2);

        XElement[] columns = table.Elements(OdfNamespaces.Table + "table-column").ToArray();
        Assert.Equal(3, columns.Length);
        Assert.All(columns, column => Assert.Same(table, column.Parent));
        Assert.Empty(table.Element(OdfNamespaces.Table + "table-header-rows")!
            .Elements(OdfNamespaces.Table + "table-column"));
        Assert.True(document.Validate().IsValid);
    }

    [Fact]
    public void FlatExportTreatsMissingStylesPartAsEmptyStyles() {
        using OdtDocument document = OdtDocument.Create();
        document.AddParagraph("Minimal flat document");
        document.Package.RemoveEntry("styles.xml");
        using var stream = new MemoryStream();

        document.SaveFlatXml(stream);

        stream.Position = 0;
        using OdtDocument reopened = OdtDocument.OpenFlatXml(stream);
        Assert.Equal("Minimal flat document", reopened.Paragraphs.Single().Text);
        Assert.DoesNotContain("styles.xml", document.LastSaveReport!.RewrittenEntries);
    }

    [Fact]
    public void FlatImageExtractionPreservesSupportedMimeType() {
        byte[] webp = { 0x52, 0x49, 0x46, 0x46, 0x04, 0x00, 0x00, 0x00, 0x57, 0x45, 0x42, 0x50 };
        using OdtDocument document = OdtDocument.Create();
        document.AddParagraph().AddImage(webp, "pixel.webp",
            OdfLength.Centimeters(1), OdfLength.Centimeters(1));
        using var stream = new MemoryStream();

        document.SaveFlatXml(stream);
        stream.Position = 0;
        using OdtDocument reopened = OdtDocument.OpenFlatXml(stream);

        OdtImage image = reopened.Paragraphs.Single().Images.Single();
        Assert.EndsWith(".webp", image.Path, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(webp, image.GetImageBytes());
    }

    [Fact]
    public void PackageVersionFallsBackToContentPartWhenManifestVersionIsMissing() {
        using OdtDocument source = OdtDocument.Create();
        source.AddParagraph("ODF 1.3");
        byte[] package = RewriteWithoutManifestVersion(source.ToBytes(), "1.3");

        using OdtDocument reopened = OdtDocument.Open(new MemoryStream(package));

        Assert.Equal(OdfVersion.V1_3, reopened.Version);
        Assert.DoesNotContain(reopened.Diagnostics, diagnostic => diagnostic.Id == "ODF003");
    }

    [Fact]
    public void GraphicColorsAreHiddenWhenFillOrStrokeModeIsNone() {
        using OdpPresentation presentation = OdpPresentation.Create();
        OdpRectangle rectangle = presentation.AddSlide("Shape").AddRectangle(
            OdfRect.FromCentimeters(1, 1, 4, 2));
        rectangle.FillColor = OdfColor.Parse("#112233");
        rectangle.StrokeColor = OdfColor.Parse("#445566");
        string styleName = (string)rectangle.Element.Attribute(OdfNamespaces.Draw + "style-name")!;
        XElement properties = presentation.Package.GetXml("content.xml")
            .Descendants(OdfNamespaces.Style + "style")
            .Single(style => (string?)style.Attribute(OdfNamespaces.Style + "name") == styleName)
            .Element(OdfNamespaces.Style + "graphic-properties")!;
        properties.SetAttributeValue(OdfNamespaces.Draw + "fill", "none");
        properties.SetAttributeValue(OdfNamespaces.Draw + "stroke", "none");

        Assert.Null(rectangle.FillColor);
        Assert.Null(rectangle.StrokeColor);
    }

    private static void WrapFirstRowAsHeader(XElement table) {
        XElement firstRow = table.Elements(OdfNamespaces.Table + "table-row").First();
        XElement secondRow = firstRow.ElementsAfterSelf(OdfNamespaces.Table + "table-row").First();
        firstRow.Remove();
        secondRow.AddBeforeSelf(new XElement(OdfNamespaces.Table + "table-header-rows", firstRow));
    }

    private static byte[] RewriteWithoutManifestVersion(byte[] sourceBytes, string partVersion) {
        using var sourceStream = new MemoryStream(sourceBytes, writable: false);
        using var output = new MemoryStream();
        using (var source = new ZipArchive(sourceStream, ZipArchiveMode.Read, leaveOpen: false))
        using (var target = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            foreach (ZipArchiveEntry sourceEntry in source.Entries) {
                byte[] bytes;
                using (Stream entryStream = sourceEntry.Open())
                using (var copy = new MemoryStream()) {
                    entryStream.CopyTo(copy);
                    bytes = copy.ToArray();
                }
                if (sourceEntry.FullName == "META-INF/manifest.xml") {
                    XDocument manifest = XDocument.Parse(Encoding.UTF8.GetString(bytes));
                    manifest.Root!.Attribute(OdfNamespaces.Manifest + "version")?.Remove();
                    manifest.Root.Elements(OdfNamespaces.Manifest + "file-entry")
                        .First(element => (string?)element.Attribute(OdfNamespaces.Manifest + "full-path") == "/")
                        .Attribute(OdfNamespaces.Manifest + "version")?.Remove();
                    bytes = Encoding.UTF8.GetBytes(manifest.ToString(SaveOptions.DisableFormatting));
                } else if (sourceEntry.FullName == "content.xml") {
                    XDocument content = XDocument.Parse(Encoding.UTF8.GetString(bytes));
                    content.Root!.SetAttributeValue(OdfNamespaces.Office + "version", partVersion);
                    bytes = Encoding.UTF8.GetBytes(content.ToString(SaveOptions.DisableFormatting));
                }
                ZipArchiveEntry targetEntry = target.CreateEntry(sourceEntry.FullName,
                    sourceEntry.FullName == "mimetype" ? CompressionLevel.NoCompression : CompressionLevel.Optimal);
                using Stream targetStream = targetEntry.Open();
                targetStream.Write(bytes, 0, bytes.Length);
            }
        }
        return output.ToArray();
    }
}
