using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO;
using OfficeIMO.Word;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace OfficeIMO.Tests;

public sealed class WordImageLengthTests {
    private static byte[] Gif => System.Convert.FromBase64String("R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7");

    [Fact]
    public void LengthsKeepExactFrameEmusAndSignedFloatingOffsets() {
        using var document = WordDocument.Create();
        using var source = new MemoryStream(Gif);
        var image = document.AddParagraph().AddImage(source, "pixel.gif", OfficeLength.Points(12), OfficeLength.Points(8),
            options: new WordImageOptions { Wrapping = WordImageTextWrapping.InFrontOfText });
        image.Width = OfficeLength.Emus(12345);
        image.Height = OfficeLength.Emus(67891);
        image.HorizontalPositionOffset = OfficeLength.Millimeters(1);
        image.VerticalPositionOffset = OfficeLength.Points(-3);
        image.BlurRadius = OfficeLength.Points(1);
        image.GrayScale = true;
        image.LuminanceBrightness = 20;
        byte[] bytes = document.ToBytes();
        using (var package = WordprocessingDocument.Open(new MemoryStream(bytes), false)) {
            var anchor = package.MainDocumentPart!.Document.Descendants<DW.Anchor>().Single();
            Assert.Equal(12345L, anchor.Extent!.Cx!.Value);
            Assert.Equal(67891L, anchor.Extent.Cy!.Value);
            var picture = anchor.Descendants<DocumentFormat.OpenXml.Drawing.Pictures.Picture>().Single();
            Assert.Equal(12345L, picture.ShapeProperties!.Transform2D!.Extents!.Cx!.Value);
            Assert.Equal(67891L, picture.ShapeProperties.Transform2D.Extents.Cy!.Value);
            Assert.Equal("36000", anchor.HorizontalPosition!.PositionOffset!.Text);
            Assert.Equal("-38100", anchor.VerticalPosition!.PositionOffset!.Text);
            Assert.Equal(12700L, picture.BlipFill!.Blip!.GetFirstChild<A.Blur>()!.Radius!.Value);
            Assert.IsType<A.BlipExtensionList>(picture.BlipFill!.Blip!.LastChild);
            var errors = new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Microsoft365).Validate(package).ToArray();
            Assert.True(errors.Length == 0, string.Join(System.Environment.NewLine, errors.Select(error => error.Description)));
        }
        using var loaded = WordDocument.Load(new MemoryStream(bytes));
        var restored = Assert.Single(loaded.Images);
        Assert.Equal(OfficeLength.Emus(12345), restored.Width);
        Assert.Equal(OfficeLength.Emus(67891), restored.Height);
        Assert.Equal(OfficeLength.Points(-3), restored.VerticalPositionOffset);
        restored.BlurGrow = true;
        Assert.Equal(OfficeLength.Points(1), restored.BlurRadius);
    }

    [Fact]
    public void SourceCropRatiosRetainTheDisplayedFrameAndNativePercentagePrecision() {
        using var document = WordDocument.Create();
        using var source = new MemoryStream(Gif);
        var image = document.AddParagraph().AddImage(source, "pixel.gif", OfficeLength.Centimeters(2), OfficeLength.Centimeters(1));
        image.CropLeftRatio = -.05;
        image.CropRightRatio = .2;
        image.CropTopRatio = .1;
        image.CropBottomRatio = .15;
        image.CropTopRatio = null;
        Assert.Equal(OfficeLength.Centimeters(2), image.Width);
        Assert.Equal(OfficeLength.Centimeters(1), image.Height);
        byte[] bytes = document.ToBytes();
        using (var package = WordprocessingDocument.Open(new MemoryStream(bytes), false)) {
            var crop = package.MainDocumentPart!.Document.Descendants<A.SourceRectangle>().Single();
            Assert.Equal(-5000, crop.Left!.Value);
            Assert.Equal(20000, crop.Right!.Value);
            Assert.Null(crop.Top);
            Assert.Equal(15000, crop.Bottom!.Value);
        }
        using var loaded = WordDocument.Load(new MemoryStream(bytes));
        var restored = Assert.Single(loaded.Images);
        Assert.Equal(-.05, restored.CropLeftRatio);
        Assert.Null(restored.CropTopRatio);
        Assert.Equal(OfficeLength.Centimeters(2), restored.Width);
    }
}
