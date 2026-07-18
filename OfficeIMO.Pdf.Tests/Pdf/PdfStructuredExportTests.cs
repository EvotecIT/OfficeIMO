using System.Text.Json;
using System.Xml.Linq;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfStructuredExportTests {
    [Fact]
    public void StructuredExportEngineIsNotASecondPublicFacade() {
        Type[] exportedTypes = typeof(PdfDocument).Assembly.GetExportedTypes();

        Assert.Contains(exportedTypes, type => type == typeof(PdfLogicalDocumentStructuredExportExtensions));
        Assert.DoesNotContain(exportedTypes, type => type.Name == "PdfStructuredExporter");
        Assert.DoesNotContain(exportedTypes, type => type.Name == "PdfStructuredExportEngine");
    }

    [Fact]
    public void StructuredExporter_UsesOneLogicalModelAcrossFiveFormats() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Structured <export> & JSON \"proof\""))
            .ToBytes();
        PdfLogicalDocument logical = PdfLogicalDocument.Load(source);

        string json = logical.ExportStructured(PdfStructuredExportFormat.Json);
        string markdown = logical.ExportStructured(PdfStructuredExportFormat.Markdown);
        string alto = logical.ExportStructured(PdfStructuredExportFormat.AltoXml);
        string hocr = logical.ExportStructured(PdfStructuredExportFormat.Hocr);
        string pageXml = logical.ExportStructured(PdfStructuredExportFormat.PageXml);

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
            logical.ExportStructured(PdfStructuredExportFormat.PageXml));
        IReadOnlyList<string> pages = logical.ToPageXmlDocuments();

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
