using OfficeIMO.Drawing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public sealed class PdfParagraphImageAnchorsTests {
    [Fact]
    public void LineRelativeImageReportsFlowFallbackRatherThanUsingTheFirstParagraphLine() {
        using WordDocument word = WordDocument.Create();
        var paragraph = word.AddParagraph("First line\nAnchor line");
        using var source = new MemoryStream(OfficeRasterImageEncoder.Encode(
            new OfficeRasterImage(16, 16, OfficeColor.Red), OfficeImageExportFormat.Png));
        var image = paragraph.AddImage(source, "line-anchor.png", OfficeIMO.OfficeLength.Pixels(32), OfficeIMO.OfficeLength.Pixels(32), options: new OfficeIMO.Word.WordImageOptions { Wrapping = WordImageTextWrapping.InFrontOfText });
        image.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Page;
        image.HorizontalPositionOffset = OfficeIMO.OfficeLength.Emus(30 * 12700);
        image.VerticalPositionRelativeFrom = WordVerticalRelativePosition.Line;
        image.VerticalPositionOffset = OfficeIMO.OfficeLength.Emus(0);
        var conversion = word.ConvertToPdfWithEvidence();
        Assert.Contains(conversion.Warnings, warning => warning.Code == "native-anchored-image-flowed");
        var pdf = PdfCore.PdfDocument.Load(conversion.ToBytes());
        Assert.Single(pdf.Images.Placements());
        Assert.Contains("Anchor line", pdf.Reader.Text());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NonWrappingImageRetainsMarginAlignmentAndParagraphOffset(bool behind) {
        using WordDocument word = WordDocument.Create();
        word.Sections[0].PageSettings.Width = 6000;
        word.Sections[0].PageSettings.Height = 6000;
        word.AddParagraph("Earlier text");
        var anchor = word.AddParagraph("Anchor text");
        using var source = new MemoryStream(OfficeRasterImageEncoder.Encode(
            new OfficeRasterImage(16, 16, OfficeColor.Red), OfficeImageExportFormat.Png));
        var image = anchor.AddImage(source, "anchor.png", OfficeIMO.OfficeLength.Pixels(40), OfficeIMO.OfficeLength.Pixels(40), options: new OfficeIMO.Word.WordImageOptions { Wrapping = behind ? WordImageTextWrapping.BehindText : WordImageTextWrapping.InFrontOfText });
        image.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Margin;
        image.HorizontalPositionOffset = null;
        image._Image.Anchor!.HorizontalPosition!.HorizontalAlignment =
            new DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalAlignment { Text = "right" };
        image.VerticalPositionRelativeFrom = WordVerticalRelativePosition.Paragraph;
        image.VerticalPositionOffset = OfficeIMO.OfficeLength.Emus(-5 * 12700);
        var conversion = word.ConvertToPdfWithEvidence(new WordPdfExportOptions {
            Margins = PdfCore.PdfPageMargins.Uniform(30),
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic()
        });
        var pdf = PdfCore.PdfDocument.Load(conversion.ToBytes());
        Assert.Equal(1, pdf.Inspect().PageCount);
        var placement = Assert.Single(pdf.Images.Placements());
        Assert.Equal(240D, placement.X, 3);
        Assert.DoesNotContain(conversion.Warnings, warning => warning.Code == "native-anchored-image-flowed");
        // Image top follows the actual paragraph, rather than entering document flow or using page Y=0.
        Assert.InRange(300D - placement.Y - placement.Height, 30D, 80D);
        var page = pdf.Render.Pages(PdfCore.PdfPageSelection.From(1), new PdfCore.PdfPageRenderOptions { Dpi = 72 })[0];
        Assert.True(OfficePngReader.TryDecode(page.Bytes!, out var bitmap));
        var pixels = bitmap!.GetPixels();
        Assert.Contains(Enumerable.Range(0, pixels.Length / 4), index =>
            pixels[index * 4] > 240 && pixels[index * 4 + 1] < 20 && pixels[index * 4 + 2] < 20);
    }
}
