using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordPdfImageTransformTests {
    [Theory]
    [InlineData("body", 0, false, false)]
    [InlineData("body", 90, false, false)]
    [InlineData("body", 27, true, false)]
    [InlineData("table", 90, true, false)]
    [InlineData("table", 0, false, true)]
    [InlineData("header", 90, false, false)]
    [InlineData("footer", 0, true, false)]
    [InlineData("simple-header", 27, false, true)]
    [InlineData("simple-footer", 90, true, false)]
    [InlineData("foreground", 90, false, false)]
    [InlineData("background", 27, true, false)]
    [InlineData("simple-position", 27, true, false)]
    public void SavedDocxImageRetainsSourceCropMirrorsAndClockwiseRotation(string context, int rotation, bool horizontal, bool vertical) {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Body marker");
        WordParagraph paragraph = context switch {
            "table" => document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0],
            "header" or "simple-header" => document.HeaderDefaultOrCreate.AddParagraph(),
            "footer" or "simple-footer" => document.FooterDefaultOrCreate.AddParagraph(),
            _ => document.AddParagraph()
        };
        using var pixels = new MemoryStream(OfficeIMO.Tests.Pdf.PdfPngTestImages.CreateRgbPng(4, 2));
        WordImage image = paragraph.AddImage(pixels, "image.png", 80, 40, description: "Transformed picture");
        image.CropLeft = 25000; image.CropTop = 12500;
        image.CropRight = 12500; image.CropBottom = 25000;
        image.Rotation = rotation;
        image.HorizontalFlip = horizontal;
        image.VerticalFlip = vertical;
        if (context is "foreground" or "background" or "simple-position") {
            image.WrapText = context == "background" ? WordImageTextWrapping.BehindText : WordImageTextWrapping.InFrontOfText;
            image.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Page;
            image.VerticalPositionRelativeFrom = WordVerticalRelativePosition.Page;
            image.HorizontalPositionOffset = 144L * 12700;
            image.VerticalPositionOffset = 180L * 12700;
            if (context == "simple-position") {
                image._Image.Anchor!.SimplePos = true;
                image._Image.Anchor.SimplePosition!.X = 144L * 12700;
                image._Image.Anchor.SimplePosition.Y = 180L * 12700;
            }
        }
        using var source = new MemoryStream(document.ToBytes());
        using WordDocument loaded = WordDocument.Load(source);
        string original = loaded._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        var options = new WordPdfExportOptions { IncludePageNumbers = false };
        if (context.StartsWith("simple-", StringComparison.Ordinal))
            options.PdfOptions = new PdfCore.PdfCreateOptions { HeaderFontSize = 10, FooterFontSize = 10 };
        PdfCore.PdfDocumentConversionResult conversion = loaded.ConvertToPdfWithEvidence(options);
        byte[] pdf = conversion.ToBytes();
        PdfCore.PdfImagePlacement placement = Assert.Single(PdfCore.PdfImageExtractor.ExtractImagePlacements(pdf));

        // 60 by 30 points is the authored visible frame; each crop axis leaves 62.5% of the source.
        double radians = -rotation * Math.PI / 180D;
        double scaleX = horizontal ? -96D : 96D;
        double scaleY = vertical ? -48D : 48D;
        Assert.Equal(Math.Cos(radians) * scaleX, placement.A, 3);
        Assert.Equal(Math.Sin(radians) * scaleX, placement.B, 3);
        Assert.Equal(-Math.Sin(radians) * scaleY, placement.C, 3);
        Assert.Equal(Math.Cos(radians) * scaleY, placement.D, 3);
        Assert.NotNull(placement.Clip);
        Assert.Equal(original, loaded._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
        Assert.DoesNotContain(conversion.Report.Warnings, warning => warning.Code == "native-anchored-image-flowed");
        if (context is "foreground" or "background" or "simple-position") {
            double centerX = 174D;
            PdfCore.PdfReadDocument read = PdfCore.PdfReadDocument.Open(pdf);
            double centerY = read.Pages[0].GetVisualPageSize().Height - 195D;
            // The retained crop center maps to the anchor frame center even when the source crop is asymmetric.
            double visibleCenterX = .25D + .625D / 2D;
            double visibleCenterY = .25D + .625D / 2D;
            Assert.Equal(centerX, placement.A * visibleCenterX + placement.C * visibleCenterY + placement.E, 3);
            Assert.Equal(centerY, placement.B * visibleCenterX + placement.D * visibleCenterY + placement.F, 3);
        }
    }
}
