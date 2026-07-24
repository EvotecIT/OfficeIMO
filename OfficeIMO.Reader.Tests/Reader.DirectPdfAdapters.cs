using System.IO.Compression;
using OfficeIMO.Pdf;
using OfficeIMO.Reader.All;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderDirectPdfAdapterTests {
    [Fact]
    public void EmailFacade_ProducesSearchablePdfWithEmailPolicyEvidence() {
        byte[] source = Encoding.UTF8.GetBytes(
            "From: sender@example.test\r\n" +
            "To: recipient@example.test\r\n" +
            "Subject: Delivery update\r\n" +
            "MIME-Version: 1.0\r\n" +
            "Content-Type: text/plain; charset=utf-8\r\n\r\n" +
            "The delivery is ready for review.\r\n");
        using var stream = new MemoryStream(source);

        PdfDocumentConversionResult conversion = OfficeDocumentPdfConverter.EmailToPdf(stream);
        byte[] pdf = conversion.ToBytes();

        Assert.Contains("delivery is ready", PdfReadDocument.Open(pdf).ExtractText(), StringComparison.OrdinalIgnoreCase);
        Assert.Contains(conversion.Warnings, static warning => warning.Code == "reader-email-policy");
    }

    [Fact]
    public void EpubFacade_PreservesChapterOrderAndReportsEpubPolicy() {
        using MemoryStream source = CreateEpub();

        PdfDocumentConversionResult conversion = OfficeDocumentPdfConverter.EpubToPdf(source);
        string text = PdfReadDocument.Open(conversion.ToBytes()).ExtractText();

        Assert.True(text.IndexOf("First chapter", StringComparison.Ordinal) <
                    text.IndexOf("Second chapter", StringComparison.Ordinal));
        Assert.Contains(conversion.Warnings, static warning => warning.Code == "reader-epub-policy");
    }

    [Fact]
    public void VisioFacade_ProducesSearchableSemanticFallbackWithEvidence() {
        VisioDocument source = VisioDocument.Create();
        VisioPage page = source.AddPage("Topology", 8, 5);
        page.Shapes.Add(new VisioShape("gateway") { Text = "Gateway service" });
        using var stream = new MemoryStream(source.ToBytes());

        PdfDocumentConversionResult conversion = OfficeDocumentPdfConverter.VisioToPdf(stream);
        string text = PdfReadDocument.Open(conversion.ToBytes()).ExtractText();

        Assert.Contains("Gateway service", text, StringComparison.Ordinal);
        Assert.Contains(conversion.Warnings, static warning => warning.Code == "reader-visio-semantic-fallback");
    }

    private static MemoryStream CreateEpub() {
        var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            WriteEntry(archive, "mimetype", "application/epub+zip", CompressionLevel.NoCompression);
            WriteEntry(
                archive,
                "META-INF/container.xml",
                "<container version=\"1.0\" xmlns=\"urn:oasis:names:tc:opendocument:xmlns:container\"><rootfiles>" +
                "<rootfile full-path=\"OEBPS/content.opf\" media-type=\"application/oebps-package+xml\"/>" +
                "</rootfiles></container>");
            WriteEntry(
                archive,
                "OEBPS/content.opf",
                "<package version=\"3.0\" xmlns=\"http://www.idpf.org/2007/opf\"><metadata xmlns:dc=\"http://purl.org/dc/elements/1.1/\"><dc:title>Adapter book</dc:title></metadata><manifest>" +
                "<item id=\"one\" href=\"one.xhtml\" media-type=\"application/xhtml+xml\"/>" +
                "<item id=\"two\" href=\"two.xhtml\" media-type=\"application/xhtml+xml\"/>" +
                "</manifest><spine><itemref idref=\"one\"/><itemref idref=\"two\"/></spine></package>");
            WriteEntry(
                archive,
                "OEBPS/one.xhtml",
                "<html xmlns=\"http://www.w3.org/1999/xhtml\"><body><h1>First chapter</h1><p>Alpha.</p></body></html>");
            WriteEntry(
                archive,
                "OEBPS/two.xhtml",
                "<html xmlns=\"http://www.w3.org/1999/xhtml\"><body><h1>Second chapter</h1><p>Beta.</p></body></html>");
        }
        output.Position = 0;
        return output;
    }

    private static void WriteEntry(
        ZipArchive archive,
        string path,
        string content,
        CompressionLevel compression = CompressionLevel.Optimal) {
        ZipArchiveEntry entry = archive.CreateEntry(path, compression);
        using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
        writer.Write(content);
    }
}
