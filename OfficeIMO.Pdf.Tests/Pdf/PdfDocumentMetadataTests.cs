using System.IO;
using System.Text;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfDocumentMetadataTests {
    [Fact]
    public void Meta_EncodesTextStringsAndInspectorReadsOriginalValues() {
        const string title = "Quarterly (Q1) \\ Roadmap";
        const string author = "OfficeIMO\nTeam";
        const string subject = "PDF metadata (escaped) \\ subject";
        const string keywords = "alpha,beta\r\ngamma\tomega";

        byte[] bytes = PdfDocument.Create()
            .Meta(title: title, author: author, subject: subject, keywords: keywords)
            .Paragraph(p => p.Text("Metadata body."))
            .ToBytes();

        string pdfText = Encoding.ASCII.GetString(bytes);

        Assert.Contains("/Title " + PdfSyntaxEscaper.TextString(title), pdfText, StringComparison.Ordinal);
        Assert.Contains("/Author " + PdfSyntaxEscaper.TextString(author), pdfText, StringComparison.Ordinal);
        Assert.Contains("/Subject " + PdfSyntaxEscaper.TextString(subject), pdfText, StringComparison.Ordinal);
        Assert.Contains("/Keywords " + PdfSyntaxEscaper.TextString(keywords), pdfText, StringComparison.Ordinal);

        PdfDocumentInfo info = PdfInspector.Inspect(bytes);

        Assert.Equal(title, info.Metadata.Title);
        Assert.Equal(author, info.Metadata.Author);
        Assert.Equal(subject, info.Metadata.Subject);
        Assert.Equal(keywords, info.Metadata.Keywords);
    }

    [Fact]
    public void Meta_CanClearPreviouslySetFields() {
        var doc = PdfDocument.Create()
            .Meta(title: "Initial Title", author: "Initial Author", subject: "Initial Subject", keywords: "alpha, beta");

        doc.Meta(title: string.Empty, author: string.Empty, subject: string.Empty, keywords: string.Empty);

        byte[] bytes = doc.ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var info = pdf.Information;

        Assert.Null(info.Title);
        Assert.Null(info.Author);
        Assert.Null(info.Subject);
        Assert.Null(info.Keywords);
    }

    [Fact]
    public void Meta_ClearingSingleFieldRetainsOtherMetadata() {
        var doc = PdfDocument.Create()
            .Meta(title: "Retained Title", author: "Author To Clear");

        doc.Meta(author: string.Empty);

        byte[] bytes = doc.ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var info = pdf.Information;

        Assert.Equal("Retained Title", info.Title);
        Assert.Null(info.Author);
    }

    [Fact]
    public void MetadataReaders_DecodeUtf8AndUtf16HexTextStrings() {
        const string title = "Café UTF8";
        const string author = "Big Endian";
        const string subject = "Little Endian";
        byte[] bytes = BuildMetadataHexPdf(
            Utf8BomHex(title),
            Utf16BigEndianBomHex(author),
            Utf16LittleEndianBomHex(subject));

        PdfDocumentInfo inspected = PdfInspector.Inspect(bytes);
        var extracted = PdfTextExtractor.GetMetadata(bytes);

        Assert.Equal(title, inspected.Metadata.Title);
        Assert.Equal(author, inspected.Metadata.Author);
        Assert.Equal(subject, inspected.Metadata.Subject);
        Assert.Equal(title, extracted.Title);
        Assert.Equal(author, extracted.Author);
        Assert.Equal(subject, extracted.Subject);
    }

    [Fact]
    public void MetadataReaders_DecodeUtf8AndUtf16LiteralTextStrings() {
        const string title = "Café literal";
        const string author = "Big literal";
        const string subject = "Little literal";
        byte[] bytes = BuildMetadataLiteralPdf(
            LiteralFromBytes(Utf8BomBytes(title)),
            LiteralFromBytes(Utf16BigEndianBomBytes(author)),
            LiteralFromBytes(Utf16LittleEndianBomBytes(subject)));

        PdfDocumentInfo inspected = PdfInspector.Inspect(bytes);
        var extracted = PdfTextExtractor.GetMetadata(bytes);

        Assert.Equal(title, inspected.Metadata.Title);
        Assert.Equal(author, inspected.Metadata.Author);
        Assert.Equal(subject, inspected.Metadata.Subject);
        Assert.Equal(title, extracted.Title);
        Assert.Equal(author, extracted.Author);
        Assert.Equal(subject, extracted.Subject);
    }

    private static byte[] BuildMetadataHexPdf(string titleHex, string authorHex, string subjectHex) {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 200 200] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            $"<< /Title <{titleHex}> /Author <{authorHex}> /Subject <{subjectHex}> >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Info 5 0 R /Size 6 >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildMetadataLiteralPdf(string titleLiteral, string authorLiteral, string subjectLiteral) {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 200 200] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            $"<< /Title ({titleLiteral}) /Author ({authorLiteral}) /Subject ({subjectLiteral}) >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Info 5 0 R /Size 6 >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static string Utf8BomHex(string value) {
        return ToHex(Utf8BomBytes(value));
    }

    private static string Utf16BigEndianBomHex(string value) {
        return ToHex(Utf16BigEndianBomBytes(value));
    }

    private static string Utf16LittleEndianBomHex(string value) {
        return ToHex(Utf16LittleEndianBomBytes(value));
    }

    private static byte[] Utf8BomBytes(string value) {
        byte[] payload = Encoding.UTF8.GetBytes(value);
        var bytes = new byte[payload.Length + 3];
        bytes[0] = 0xEF;
        bytes[1] = 0xBB;
        bytes[2] = 0xBF;
        Buffer.BlockCopy(payload, 0, bytes, 3, payload.Length);
        return bytes;
    }

    private static byte[] Utf16BigEndianBomBytes(string value) {
        byte[] payload = Encoding.BigEndianUnicode.GetBytes(value);
        var bytes = new byte[payload.Length + 2];
        bytes[0] = 0xFE;
        bytes[1] = 0xFF;
        Buffer.BlockCopy(payload, 0, bytes, 2, payload.Length);
        return bytes;
    }

    private static byte[] Utf16LittleEndianBomBytes(string value) {
        byte[] payload = Encoding.Unicode.GetBytes(value);
        var bytes = new byte[payload.Length + 2];
        bytes[0] = 0xFF;
        bytes[1] = 0xFE;
        Buffer.BlockCopy(payload, 0, bytes, 2, payload.Length);
        return bytes;
    }

    private static string LiteralFromBytes(byte[] bytes) {
        var builder = new StringBuilder(bytes.Length * 4);
        for (int i = 0; i < bytes.Length; i++) {
            builder.Append('\\');
            builder.Append(Convert.ToString(bytes[i], 8).PadLeft(3, '0'));
        }

        return builder.ToString();
    }

    private static string ToHex(byte[] bytes) {
        var builder = new StringBuilder(bytes.Length * 2);
        for (int i = 0; i < bytes.Length; i++) {
            builder.Append(bytes[i].ToString("X2", System.Globalization.CultureInfo.InvariantCulture));
        }

        return builder.ToString();
    }
}
