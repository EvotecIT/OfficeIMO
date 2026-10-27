using System;
using System.IO;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(96D, 96D, .5D, 96D, 96D)]
    [InlineData(300D, 300D, 1D, 299.9994D, 299.9994D)]
    [InlineData(300D, 150D, 2D, 299.9994D, 150.0124D)]
    public void PhysicalImageCropUsesOriginalResolutionAndPreservesDisplayedScale(double dpiX, double dpiY,
        double displayScale, double storedDpiX, double storedDpiY) {
        using WordDocument document = WordDocument.Create();
        using var pixels = new MemoryStream(CreatePhysicalCropTestImage(dpiX, dpiY));
        var image = document.AddParagraph().AddImage(pixels, "crop.png", 360D * displayScale, 240D * displayScale);
        // PNG pHYs stores integer pixels per metre: 11811 and 5906 encode these exact DPIs.
        // The physical crop follows the stored source resolution, rather than the encoder's requested nominal DPI.
        double originalWidth = 360D / storedDpiX * 2.54D;
        double originalHeight = 240D / storedDpiY * 2.54D;
        image.CropLeftCentimeters = originalWidth * .1D;
        image.CropRightCentimeters = originalWidth * .2D;
        image.CropTopCentimeters = originalHeight * .1D;
        image.CropBottomCentimeters = originalHeight * .2D;
        using var input = new MemoryStream(document.ToBytes());
        using WordDocument reopened = WordDocument.Load(input);
        image = Assert.Single(reopened.Images);
        string xml = reopened._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        Assert.Equal(10000, image.CropLeft);
        Assert.Equal(20000, image.CropRight);
        Assert.Equal(10000, image.CropTop);
        Assert.Equal(20000, image.CropBottom);
        Assert.Equal(252D * displayScale, image.Width!.Value, 3);
        Assert.Equal(168D * displayScale, image.Height!.Value, 3);
        Assert.InRange(Math.Abs(image.CropLeftCentimeters!.Value - originalWidth * .1D), 0D, .0001D);
        Assert.InRange(Math.Abs(image.CropBottomCentimeters!.Value - originalHeight * .2D), 0D, .0001D);
        Assert.Equal(xml, reopened._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
        // Resizing later keeps the same source percentage and original-image crop length.
        image.Width *= 2D;
        Assert.Equal(10000, image.CropLeft);
        Assert.InRange(Math.Abs(image.CropLeftCentimeters!.Value - originalWidth * .1D), 0D, .0001D);
    }

    [Fact]
    public void PhysicalImageCropNegativeEdgesAndClearingRetainTheOriginalImageScale() {
        using WordDocument document = WordDocument.Create();
        using var pixels = new MemoryStream(CreatePhysicalCropTestImage());
        var image = document.AddParagraph().AddImage(pixels, "crop.png", 360, 240);
        image.CropLeftCentimeters = -.9525D;
        image.CropRightCentimeters = 1.905D;
        using var input = new MemoryStream(document.ToBytes());
        using WordDocument reopened = WordDocument.Load(input);
        image = Assert.Single(reopened.Images);
        Assert.Equal(-10000, image.CropLeft);
        Assert.Equal(20000, image.CropRight);
        Assert.Equal(-.9525D, image.CropLeftCentimeters!.Value, 5);
        Assert.Equal(324D, image.Width!.Value, 3);
        image.CropLeftCentimeters = null;
        Assert.Null(image.CropLeft);
        Assert.Equal(20000, image.CropRight);
        Assert.Equal(288D, image.Width!.Value, 3);
        image.CropRightCentimeters = null;
        Assert.Null(image.CropRight);
        Assert.Equal(360D, image.Width!.Value, 3);
    }

    [Fact]
    public void PhysicalImageCropEditsImportedPercentagesWithoutChangingRawCropSemantics() {
        using WordDocument document = WordDocument.Create();
        using var pixels = new MemoryStream(CreatePhysicalCropTestImage());
        var image = document.AddParagraph().AddImage(pixels, "crop.png", 100, 80);
        image.CropLeft = 25000;
        image.CropRight = 12500;
        Assert.Equal(100D, image.Width!.Value, 3);
        using var input = new MemoryStream(document.ToBytes());
        using WordDocument reopened = WordDocument.Load(input);
        image = Assert.Single(reopened.Images);
        Assert.Equal(2.38125D, image.CropLeftCentimeters!.Value, 5);
        image.CropLeftCentimeters = .9525D;
        Assert.Equal(10000, image.CropLeft);
        Assert.Equal(12500, image.CropRight);
        Assert.Equal(124D, image.Width!.Value, 3);
        Assert.Equal(80D, image.Height!.Value, 3);
    }

    [Fact]
    public void ImageCropValidatesTheWholeReplacementBeforeChangingTheDocument() {
        // The old right crop and new left crop would conflict if four setters were applied sequentially.
        using WordDocument document = WordDocument.Create();
        using var source = new MemoryStream(CreatePhysicalCropTestImage());
        var image = document.AddParagraph().AddImage(source, "crop.png");
        image.SetCropCentimeters(0D, 0D, 7.62D, 0D);
        image.SetCropCentimeters(7.62D, 0D, 0D, 0D);
        Assert.Same(image._Image, Assert.Single(document.Images)._Image);
        Assert.Equal(80000, image.CropLeft);
        Assert.Equal(0, image.CropRight);
        Assert.Equal(72D, image.Width!.Value, 3);
        string xml = document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        Assert.Throws<ArgumentOutOfRangeException>(() => image.SetCropCentimeters(7.62D, 6.35D, 0D, 0D));
        Assert.Equal(xml, document._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
        Assert.Throws<ArgumentOutOfRangeException>(() => image.CropLeftCentimeters = double.NaN);
        Assert.Equal(xml, document._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
    }

    [Theory]
    [InlineData(0, false, false, false, 110D, 100D)]
    [InlineData(27, false, false, false, 109.455D, 102.270D)]
    [InlineData(90, false, false, false, 105D, 105D)]
    [InlineData(0, true, false, false, 100D, 100D)]
    [InlineData(0, false, true, true, 100D, 100D)]
    public void PhysicalImageCropPreservesTheTransformedFloatingPictureOrigin(int rotation, bool flipH,
        bool flipV, bool cropTop, double leftPoints, double topPoints) {
        using WordDocument document = WordDocument.Create();
        using var pixels = new MemoryStream(CreatePhysicalCropTestImage());
        var image = document.AddParagraph().AddImage(pixels, "crop.png", 360, 240);
        image.WrapText = WordImageTextWrapping.InFrontOfText;
        image.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Page;
        image.VerticalPositionRelativeFrom = WordVerticalRelativePosition.Page;
        image.HorizontalPositionOffset = 1270000L;
        image.VerticalPositionOffset = 1270000L;
        image.Rotation = rotation; image.HorizontalFlip = flipH; image.VerticalFlip = flipV;
        if (cropTop) image.CropTopCentimeters = 10D / 72D * 2.54D;
        else image.CropLeftCentimeters = 10D / 72D * 2.54D;
        using var input = new MemoryStream(document.ToBytes());
        using WordDocument reopened = WordDocument.Load(input);
        image = Assert.Single(reopened.Images);
        Assert.InRange(Math.Abs(image.HorizontalPositionOffset!.Value / 12700D - leftPoints), 0D, .002D);
        Assert.InRange(Math.Abs(image.VerticalPositionOffset!.Value / 12700D - topPoints), 0D, .002D);
        Assert.InRange(Math.Abs(image.Width!.Value * .75D - (cropTop ? 270D : 260D)), 0D, .002D);
        Assert.InRange(Math.Abs(image.Height!.Value * .75D - (cropTop ? 170D : 180D)), 0D, .002D);
        Assert.Equal(rotation, image.Rotation);
        Assert.Equal(flipH, image.HorizontalFlip);
        Assert.Equal(flipV, image.VerticalFlip);
    }

    [Fact]
    public void PhysicalImageCropResolvesMarginAlignmentAndSimplePositionBeforeSaving() {
        using WordDocument document = WordDocument.Create();
        var page = document.Sections[0]._sectionProperties.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.PageSize>()!;
        page.Width = 14400U;
        document.Sections[0].Margins.Left = 1080U;
        document.Sections[0].Margins.Right = 1080U;
        using var pixels = new MemoryStream(CreatePhysicalCropTestImage());
        var image = document.AddParagraph().AddImage(pixels, "crop.png", 360, 240);
        image.WrapText = WordImageTextWrapping.InFrontOfText;
        image.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Margin;
        image.horizontalPosition.RemoveAllChildren();
        image.horizontalPosition.AppendChild(new DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalAlignment("center"));
        image.CropLeftCentimeters = 10D / 72D * 2.54D;
        Assert.Null(image.horizontalPosition.HorizontalAlignment);
        Assert.InRange(image.HorizontalPositionOffset!.Value / 12700D, 180.999D, 181.002D);
        image._Image.Anchor!.SimplePos = true;
        image._Image.Anchor.SimplePosition!.X = 1270000L;
        image._Image.Anchor.SimplePosition.Y = 1270000L;
        image.CropLeftCentimeters = null;
        using var input = new MemoryStream(document.ToBytes());
        using WordDocument reopened = WordDocument.Load(input);
        image = Assert.Single(reopened.Images);
        Assert.InRange(image._Image.Anchor!.SimplePosition!.X!.Value / 12700D, 89.998D, 90.001D);
        Assert.Equal(1270000L, image._Image.Anchor.SimplePosition.Y!.Value);
        Assert.Equal(360D, image.Width!.Value, 3);
    }

    [Fact]
    public void PhysicalImageCropRejectsUnresolvedOrMalformedAnchorsWithoutChangingThePackage() {
        using WordDocument document = WordDocument.Create();
        using var pixels = new MemoryStream(CreatePhysicalCropTestImage());
        var image = document.AddParagraph().AddImage(pixels, "crop.png", 360, 240);
        image.WrapText = WordImageTextWrapping.InFrontOfText;
        image.HorizontalPositionOffset = 0L;
        image.VerticalPositionOffset = 0L;
        image.horizontalPosition.PositionOffset!.Text = "bad";
        string xml = document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        Assert.Throws<InvalidOperationException>(() => image.CropLeftCentimeters = .5D);
        Assert.Equal(xml, document._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
        image.horizontalPosition.RemoveAllChildren();
        image.horizontalPosition.AppendChild(new DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalAlignment("center"));
        image.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Character;
        xml = document._wordprocessingDocument.MainDocumentPart.Document.OuterXml;
        Assert.Throws<InvalidOperationException>(() => image.CropLeftCentimeters = .5D);
        Assert.Equal(xml, document._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
    }

    private static byte[] CreatePhysicalCropTestImage(double dpiX = 96D, double dpiY = 96D) {
        var image = new OfficeRasterImage(360, 240, OfficeColor.Red);
        for (int y = 0; y < 240; y++)
            for (int x = 0; x < 360; x++)
                image.SetPixel(x, y, y < 120 ? x < 180 ? OfficeColor.Red : OfficeColor.Green : x < 180 ? OfficeColor.Blue : OfficeColor.Yellow);
        return OfficePngWriter.Encode(image, new OfficePngEncodeOptions {
            DpiX = dpiX, DpiY = dpiY, WritePhysicalResolution = dpiX != 96D || dpiY != 96D
        });
    }
}
