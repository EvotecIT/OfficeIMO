using System.IO;
using System.Text;
using OfficeIMO.Pdf;
using UglyToad.PdfPig;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfDocMetadataTests {
    [Fact]
    public void Meta_EscapesLiteralStringsAndInspectorReadsOriginalValues() {
        const string title = "Quarterly (Q1) \\ Roadmap";
        const string author = "OfficeIMO\nTeam";
        const string subject = "PDF metadata (escaped) \\ subject";
        const string keywords = "alpha,beta\r\ngamma\tomega";

        byte[] bytes = PdfDoc.Create()
            .Meta(title: title, author: author, subject: subject, keywords: keywords)
            .Paragraph(p => p.Text("Metadata body."))
            .ToBytes();

        string pdfText = Encoding.ASCII.GetString(bytes);

        Assert.Contains("/Title (Quarterly \\(Q1\\) \\\\ Roadmap)", pdfText, StringComparison.Ordinal);
        Assert.Contains("/Author (OfficeIMO\\nTeam)", pdfText, StringComparison.Ordinal);
        Assert.Contains("/Subject (PDF metadata \\(escaped\\) \\\\ subject)", pdfText, StringComparison.Ordinal);
        Assert.Contains("/Keywords (alpha,beta\\r\\ngamma\\tomega)", pdfText, StringComparison.Ordinal);

        PdfDocumentInfo info = PdfInspector.Inspect(bytes);

        Assert.Equal(title, info.Metadata.Title);
        Assert.Equal(author, info.Metadata.Author);
        Assert.Equal(subject, info.Metadata.Subject);
        Assert.Equal(keywords, info.Metadata.Keywords);
    }

    [Fact]
    public void Meta_CanClearPreviouslySetFields() {
        var doc = PdfDoc.Create()
            .Meta(title: "Initial Title", author: "Initial Author", subject: "Initial Subject", keywords: "alpha, beta");

        doc.Meta(title: string.Empty, author: string.Empty, subject: string.Empty, keywords: string.Empty);

        byte[] bytes = doc.ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var info = pdf.Information;

        Assert.Null(info.Title);
        Assert.Null(info.Author);
        Assert.Null(info.Subject);
        Assert.Null(info.Keywords);
    }

    [Fact]
    public void Meta_ClearingSingleFieldRetainsOtherMetadata() {
        var doc = PdfDoc.Create()
            .Meta(title: "Retained Title", author: "Author To Clear");

        doc.Meta(author: string.Empty);

        byte[] bytes = doc.ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var info = pdf.Information;

        Assert.Equal("Retained Title", info.Title);
        Assert.Null(info.Author);
    }
}
