using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfPortfolioTests {
    [Fact]
    public void Save_PortfolioBuildsCollectionAndRoundTripsMetadata() {
        var options = new PdfOptions()
            .AddEmbeddedFile("cover.pdf", new byte[] { 1, 2, 3 }, "application/pdf")
            .AddEmbeddedFile("data.csv", new byte[] { 4, 5 }, "text/csv")
            .SetPortfolio(new PdfPortfolioOptions {
                View = PdfPortfolioView.Tile,
                InitialDocumentFileName = "cover.pdf",
                SortBy = PdfPortfolioFieldKind.Size,
                SortAscending = false
            }.SetField(new PdfPortfolioField(PdfPortfolioFieldKind.Size, "Bytes", 1)));

        PdfDocument document = PdfDocument.Create(options);
        document.Paragraph(paragraph => paragraph.Text("Portfolio"));

        byte[] pdf = document.ToBytes();
        PdfReadDocument read = PdfReadDocument.Open(pdf);

        Assert.StartsWith("%PDF-1.7", PdfEncoding.Latin1GetString(pdf), StringComparison.Ordinal);
        Assert.NotNull(read.Portfolio);
        Assert.Equal("T", read.Portfolio!.View);
        Assert.Equal("cover.pdf", read.Portfolio.InitialDocumentFileName);
        Assert.Equal("Size", read.Portfolio.SortField);
        Assert.False(read.Portfolio.SortAscending);
        Assert.Equal(new[] { "FileName", "Size" }, read.Portfolio.Fields.Select(field => field.Key));
        Assert.Equal(2, read.Attachments.Count);
    }

    [Fact]
    public void Save_PortfolioRequiresEmbeddedFiles() {
        PdfDocument document = PdfDocument.Create(new PdfOptions {
            Portfolio = new PdfPortfolioOptions()
        });
        document.Paragraph(paragraph => paragraph.Text("Missing attachment"));

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() => document.ToBytes());
        Assert.Contains("requires at least one embedded file", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Options_PortfolioIsDefensivelyCloned() {
        var source = new PdfPortfolioOptions();
        var options = new PdfOptions { Portfolio = source };
        source.View = PdfPortfolioView.Hidden;

        Assert.Equal(PdfPortfolioView.Details, options.Portfolio!.View);
        PdfPortfolioOptions returned = options.Portfolio!;
        returned.ClearFields();
        Assert.Single(options.Portfolio!.Fields);
    }
}
