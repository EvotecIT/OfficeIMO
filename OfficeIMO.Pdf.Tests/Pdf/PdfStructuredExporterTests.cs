using System.Text.Json;
using System.Xml.Linq;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfStructuredExporterTests {
    [Fact]
    public void StructuredExporter_UsesOneLogicalModelAcrossFiveFormats() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Structured <export> & JSON \"proof\""))
            .ToBytes();
        PdfLogicalDocument logical = PdfLogicalDocument.Load(source);

        string json = PdfStructuredExporter.Export(logical, PdfStructuredExportFormat.Json);
        string markdown = PdfStructuredExporter.Export(logical, PdfStructuredExportFormat.Markdown);
        string alto = PdfStructuredExporter.Export(logical, PdfStructuredExportFormat.AltoXml);
        string hocr = PdfStructuredExporter.Export(logical, PdfStructuredExportFormat.Hocr);
        string pageXml = PdfStructuredExporter.Export(logical, PdfStructuredExportFormat.PageXml);

        using JsonDocument parsedJson = JsonDocument.Parse(json);
        Assert.Equal("officeimo.pdf.logical.v1", parsedJson.RootElement.GetProperty("schema").GetString());
        Assert.Contains("Structured <export> & JSON \"proof\"", parsedJson.RootElement.GetProperty("pages")[0].GetProperty("lines")[0].GetProperty("text").GetString(), StringComparison.Ordinal);
        Assert.Contains("Structured \\<export\\>", markdown, StringComparison.Ordinal);
        Assert.Equal("alto", XDocument.Parse(alto).Root!.Name.LocalName);
        Assert.Equal("html", XDocument.Parse(hocr).Root!.Name.LocalName);
        Assert.Equal("PcGts", XDocument.Parse(pageXml).Root!.Name.LocalName);
        Assert.Contains("Structured &lt;export&gt; &amp; JSON", alto, StringComparison.Ordinal);
        Assert.Contains("Structured &lt;export&gt; &amp; JSON", hocr, StringComparison.Ordinal);
        Assert.Contains("Structured &lt;export&gt; &amp; JSON", pageXml, StringComparison.Ordinal);
    }

    [Fact]
    public void PageXmlExporter_ReturnsOneSchemaRootPerPage() {
        byte[] page = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Page scoped export"))
            .ToBytes();
        PdfLogicalDocument logical = PdfLogicalDocument.Load(PdfMerger.Merge(page, page));

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            PdfStructuredExporter.Export(logical, PdfStructuredExportFormat.PageXml));
        IReadOnlyList<string> pages = PdfStructuredExporter.ExportPageXmlDocuments(logical);

        Assert.Contains("page scoped", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(2, pages.Count);
        Assert.All(pages, xml => Assert.Equal("PcGts", XDocument.Parse(xml).Root!.Name.LocalName));
    }

    [Fact]
    public void FluentReader_ExportsStructuredJsonWithoutAnotherExtractionSurface() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Fluent structured output"))
            .ToBytes();

        string json = PdfDocument.Open(source).Read.ExportStructured(PdfStructuredExportFormat.Json);

        Assert.Contains("Fluent structured output", json, StringComparison.Ordinal);
    }

    [Fact]
    public void FluentReader_UsesStoredCredentialsForStructuredExport() {
        byte[] source = PdfDocument.Create(new PdfOptions().SetEncryption("open", "owner"))
            .Paragraph(paragraph => paragraph.Text("Encrypted structured output"))
            .ToBytes();
        var readOptions = new PdfReadOptions { Password = "owner" };

        string json = PdfDocument.Open(source, readOptions).Read.ExportStructured(PdfStructuredExportFormat.Json);

        Assert.Contains("Encrypted structured output", json, StringComparison.Ordinal);
    }
}
