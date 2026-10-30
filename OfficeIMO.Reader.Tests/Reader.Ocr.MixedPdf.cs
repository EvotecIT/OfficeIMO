using OfficeIMO.Pdf;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Pdf;
using OfficeIMO.Tests.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderMixedPdfOcrTests {
    [Theory]
    [InlineData(false, 180, 120, false)]
    [InlineData(true, 180, 120, true)]
    [InlineData(true, 20, 10, false)]
    public void MixedPagePolicyRetainsNativeTextAndOffersSubstantialImagesForOcr(bool includeMixed, int width, int height, bool candidateExpected) {
        var pdf = PdfDocument.Create(new PdfCreateOptions { PageWidth = global::OfficeIMO.OfficeLength.Points(300), PageHeight = global::OfficeIMO.OfficeLength.Points(220), MarginLeft = global::OfficeIMO.OfficeLength.Points(24), MarginRight = global::OfficeIMO.OfficeLength.Points(24), MarginTop = global::OfficeIMO.OfficeLength.Points(24), MarginBottom = global::OfficeIMO.OfficeLength.Points(24) });
        pdf.Content.Image(PdfPngTestImages.CreateRgbPng(230, 230, 230), width, height);
        pdf.Content.Paragraph(paragraph => paragraph.Text("Footer 1"));
        var reader = new OfficeDocumentReaderBuilder().AddPdfHandler(new ReaderPdfOptions { IncludeMixedPageOcrCandidates = includeMixed }).Build();

        var result = reader.ReadDocument(new MemoryStream(pdf.ToBytes()), "mixed.pdf");

        Assert.Contains(result.Blocks, block => block.Text.Contains("Footer 1", StringComparison.Ordinal));
        Assert.Equal(candidateExpected ? 1 : 0, result.OcrCandidates.Count);
        Assert.Equal(candidateExpected, result.Diagnostics.Any(diagnostic => diagnostic.Code == "ocr-needed"));
        if (candidateExpected) {
            var candidate = Assert.Single(result.OcrCandidates);
            Assert.Equal(1, candidate.Location.Page);
            Assert.True(candidate.TextBlockCount > 0);
            Assert.Equal(Assert.Single(result.Assets).Id, candidate.AssetId);
        }
    }

    [Fact]
    public void MixedPagePolicyIsFrozenWhenThePdfHandlerIsRegistered() {
        var options = new ReaderPdfOptions { IncludeMixedPageOcrCandidates = true, MinimumMixedPageImageAreaRatio = 0 };
        var reader = new OfficeDocumentReaderBuilder().AddPdfHandler(options).Build();
        options.IncludeMixedPageOcrCandidates = false;
        var pdf = PdfDocument.Create();
        pdf.Content.Image(PdfPngTestImages.CreateRgbPng(230, 230, 230), 20, 10);
        pdf.Content.Paragraph(paragraph => paragraph.Text("Native text"));
        Assert.Single(reader.ReadDocument(new MemoryStream(pdf.ToBytes()), "mixed.pdf").OcrCandidates);
    }
}
