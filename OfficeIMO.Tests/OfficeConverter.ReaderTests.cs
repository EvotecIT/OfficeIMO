using System.IO.Compression;
using OfficeIMO.Pdf;

using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class OfficeConverterReaderTests {
    [Fact]
    public void EmailConversion_RendersStyledMessageAndKeepsReaderTextOnlyOnRequest() {
        byte[] source = Encoding.UTF8.GetBytes(
            "From: Sender <sender@example.test>\r\n" +
            "To: recipient@example.test\r\n" +
            "Cc: reviewer@example.test\r\n" +
            "Date: Tue, 10 Nov 2026 09:30:00 +0000\r\n" +
            "Subject: Delivery update\r\n" +
            "MIME-Version: 1.0\r\n" +
            "Content-Type: multipart/mixed; boundary=\"part\"\r\n\r\n" +
            "--part\r\nContent-Type: text/html; charset=utf-8\r\n\r\n" +
            "<p>The delivery is <b>ready</b> for review.</p>\r\n" +
            "--part\r\nContent-Type: text/csv; name=\"manifest.csv\"\r\n" +
            "Content-Disposition: attachment; filename=\"manifest.csv\"\r\n\r\n" +
            "item,count\r\nboxes,4\r\n--part--\r\n");

        using var styled = new MemoryStream();
        OfficeConverter.Convert(new MemoryStream(source), OfficeDocumentFormat.Email, styled, OfficeDocumentFormat.Pdf);
        string text = PdfDocument.Load(new MemoryStream(styled.ToArray())).ExtractText();
        foreach (string expected in new[] { "Delivery update", "From", "sender@example.test", "To", "recipient@example.test",
                     "Cc", "reviewer@example.test", "Date", "2026-11-10", "delivery is", "ready", "Attachments (1)", "manifest.csv" })
            Assert.Contains(expected, text, StringComparison.Ordinal);

        using var textOnly = new MemoryStream();
        OfficeConversionReport archival = OfficeConverter.Convert(new MemoryStream(source), OfficeDocumentFormat.Email, textOnly,
            OfficeDocumentFormat.Pdf, new OfficeConversionOptions { OutputProfile = OfficeConversionProfile.TextOnly });
        Assert.Contains("delivery is", PdfDocument.Load(new MemoryStream(textOnly.ToArray())).ExtractText(), StringComparison.Ordinal);
        Assert.Contains(archival.Diagnostics, static warning => warning.Code == "pdf-projection-email-policy");
    }

    [Fact]
    public void EpubConversion_PreservesChapterOrderAndReportsEpubPolicy() {
        using MemoryStream source = CreateEpub();

        using var output = new MemoryStream();
        OfficeConversionReport conversion = OfficeConverter.Convert(source, OfficeDocumentFormat.Epub, output, OfficeDocumentFormat.Pdf);
        string text = PdfDocument.Load(new MemoryStream(output.ToArray())).ExtractText();
        Assert.True(source.CanRead);

        Assert.True(text.IndexOf("First chapter", StringComparison.Ordinal) <
                    text.IndexOf("Second chapter", StringComparison.Ordinal));
        Assert.Contains(conversion.Diagnostics, static warning => warning.Code == "pdf-projection-epub-policy");
    }

    [Fact]
    public void VisioConversion_ProducesSearchableSemanticFallbackWithEvidence() {
        VisioDocument source = VisioDocument.Create();
        VisioPage page = source.AddPage("Topology", OfficeLength.Inches(8), OfficeLength.Inches(5));
        page.Shapes.Add(new VisioShape("gateway") { Text = "Gateway service" });
        using var stream = new MemoryStream(source.ToBytes());

        using var output = new MemoryStream();
        OfficeConversionReport conversion = OfficeConverter.Convert(stream, OfficeDocumentFormat.Visio, output, OfficeDocumentFormat.Pdf,
            new OfficeConversionOptions { RouteId = "reader-visio-pdf" });
        string text = PdfDocument.Load(new MemoryStream(output.ToArray())).ExtractText();

        Assert.Contains("Gateway service", text, StringComparison.Ordinal);
        Assert.Contains(conversion.Diagnostics, static warning => warning.Code == "pdf-projection-visio-semantic-fallback");
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
