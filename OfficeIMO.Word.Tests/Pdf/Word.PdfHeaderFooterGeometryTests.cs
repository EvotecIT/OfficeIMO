using OfficeIMO.Drawing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_InlineRunningPictureRetainsItsSizeAndReservesBodySpace(bool footer) {
        using WordDocument document = WordDocument.Create();
        document.AddHeadersAndFooters();
        WordSection section = document.Sections[0];
        section.PageSettings.Width = 12240;
        section.PageSettings.Height = 15840;
        section.Margins.Top = 1440;
        section.Margins.Bottom = 1440;
        section.Margins.Left = section.Margins.Right = 1440;
        section.Margins.HeaderDistance = section.Margins.FooterDistance = 720;
        WordHeaderFooter story = footer ? section.Footer.Default! : section.Header.Default!;
        if (footer) section.Footer.Default!._footer!.RemoveAllChildren();
        else section.Header.Default!._header!.RemoveAllChildren();
        AddRunningGeometryPicture(story.AddParagraph(), 624, 416);
        var body = document.AddParagraph(string.Join("\n", Enumerable.Repeat("BodyMarker", 25)));
        body.FontFamily = "Arial";
        body.FontSize = 12;
        body.LineSpacingBeforePoints = body.LineSpacingAfterPoints = 0;
        body.LineSpacing = 240;

        byte[] bytes = document.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        var placements = PdfCore.PdfDocument.Load(bytes).Images.Placements();
        Assert.NotEmpty(placements);
        foreach (var page in pdf.GetPages()) {
            var image = Assert.Single(placements, placement => placement.PageNumber == page.Number);
            Assert.Equal(468D, image.Width, 2);
            Assert.Equal(312D, image.Height, 2);
            Assert.Equal(footer ? 36D : 444D, image.Y, 2);
            var words = page.GetWords().Where(word => word.Text == "BodyMarker").ToArray();
            Assert.NotEmpty(words);
            Assert.All(words, word => {
                if (footer) Assert.True(word.BoundingBox.Bottom >= image.Y + image.Height - .01D);
                else Assert.True(word.BoundingBox.Top <= image.Y + .01D);
            });
        }
    }

    [Fact]
    public void SaveAsPdf_InlineHeaderPictureReservesOnlyTheActiveVariantHeight() {
        using WordDocument document = WordDocument.Create();
        document.AddHeadersAndFooters();
        document.DifferentFirstPage = true;
        document.DifferentOddAndEvenPages = true;
        WordSection section = document.Sections[0];
        section.PageSettings.Width = 12240;
        section.PageSettings.Height = 15840;
        section.Margins.Top = section.Margins.Bottom = 1440;
        section.Margins.Left = section.Margins.Right = 1440;
        section.Margins.HeaderDistance = 720;
        WordHeader[] headers = { section.Header.First!, section.Header.Even!, section.Header.Default! };
        double[] heights = { 48D, 96D, 416D };
        for (int index = 0; index < headers.Length; index++) {
            headers[index]._header!.RemoveAllChildren();
            AddRunningGeometryPicture(headers[index].AddParagraph(), 96D, heights[index]);
            WordParagraph paragraph = document.AddParagraph("VariantBody");
            paragraph.PageBreakBefore = index > 0;
            paragraph.FontFamily = "Arial";
            paragraph.FontSize = 12;
            paragraph.LineSpacingBeforePoints = paragraph.LineSpacingAfterPoints = 0;
        }
        byte[] bytes = document.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(3, pdf.NumberOfPages);
        var placements = PdfCore.PdfDocument.Load(bytes).Images.Placements();
        double[] bodyBaselines = new double[3];
        for (int index = 0; index < heights.Length; index++) {
            var image = Assert.Single(placements, placement => placement.PageNumber == index + 1);
            Assert.Equal(heights[index] * .75D, image.Height, 2);
            var body = Assert.Single(pdf.GetPage(index + 1).GetWords(), word => word.Text == "VariantBody");
            Assert.True(body.BoundingBox.Top <= image.Y + .01D);
            bodyBaselines[index] = body.Letters[0].StartBaseLine.Y;
        }
        Assert.True(bodyBaselines[0] > bodyBaselines[1]);
        Assert.True(bodyBaselines[1] > bodyBaselines[2]);
    }

    [Fact]
    public void SaveAsPdf_ProjectedHeaderPictureReportsPhysicalPageClipping() {
        using WordDocument document = WordDocument.Create();
        document.AddHeadersAndFooters();
        AddRunningGeometryPicture(document.Sections[0].Header.Default!.AddParagraph(), 96, 1200);
        document.AddParagraph("BodyMarker");
        var result = document.ConvertToPdfWithEvidence(new WordPdfExportOptions {
            IncludePageNumbers = false,
            // An explicit projection typography override keeps the bounded
            // header-image path. Its oversized source box must report clipping.
            PdfOptions = new PdfCore.PdfCreateOptions { HeaderFontSize = 10D }
        });
        byte[] bytes = result.ToBytes();
        Assert.NotEmpty(bytes);
        Assert.Contains(result.Warnings, warning => warning.Code == "header-footer-page-bounds-clipped" && warning.Source == "Header");
    }

    private static void AddRunningGeometryPicture(WordParagraph paragraph, double width, double height) {
        byte[] png = OfficeRasterImageEncoder.Encode(new OfficeRasterImage(8, 8, OfficeColor.Black), OfficeImageExportFormat.Png);
        using var image = new MemoryStream(png);
        paragraph.LineSpacingBeforePoints = paragraph.LineSpacingAfterPoints = 0;
        paragraph.AddImage(image, "running-picture.png", OfficeIMO.OfficeLength.Pixels(width), OfficeIMO.OfficeLength.Pixels(height));
    }
}
