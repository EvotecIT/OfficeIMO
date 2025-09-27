using System.IO;
using OfficeIMO.Pdf;
using UglyToad.PdfPig;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfDocMetadataTests {
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
