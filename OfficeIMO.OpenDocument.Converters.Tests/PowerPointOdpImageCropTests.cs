using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Drawing;
using OfficeIMO.OpenDocument;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.OpenDocument;
using A = DocumentFormat.OpenXml.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class PowerPointOdpImageCropTests {
    [Theory]
    [InlineData(96, 96, 30, 15)]
    [InlineData(192, 96, 15, 7.5)]
    [InlineData(192, 192, 15, 7.5)]
    public void PowerPointCropUsesIntrinsicPixelsAndDensityAtEveryFrameSize(double dpiX, double dpiY,
        double leftPoints, double rightPoints) {
        byte[] bytes = Grid(dpiX, dpiY);
        using PowerPointPresentation source = PowerPointPresentation.Create();
        source.SlideSize.WidthPoints = 400;
        source.SlideSize.HeightPoints = 300;
        PowerPointSlide slide = source.AddSlide();
        foreach (double width in new[] { 120D, 240D }) {
            using var stream = new MemoryStream(bytes);
            PowerPointPicture picture = slide.AddPicture(stream, OfficeImageFormat.Png,
                PowerPointLayoutBox.FromPoints(20, 30, width, 90));
            picture.Crop(25, 12.5, 12.5, 12.5);
        }

        OdfConversionResult<OdpPresentation> conversion = source.ToOpenDocumentResult();
        OdpPresentation reopened = OdpPresentation.Load(new MemoryStream(conversion.Value.ToBytes()));
        OdpImage[] images = reopened.Slides[0].Shapes.OfType<OdpImage>().ToArray();
        Assert.Equal(2, images.Length);
        foreach (OdpImage image in images) {
            Assert.Equal(bytes, image.GetImageBytes());
            // PNG pixels-per-meter and ODF absolute-length serialization both quantize the authored values.
            Assert.InRange(image.Crop!.Value.Left.ToPoints(), leftPoints - .005, leftPoints + .005);
            Assert.InRange(image.Crop.Value.Right.ToPoints(), rightPoints - .005, rightPoints + .005);
            double vertical = dpiY == 96 ? 7.5 : 3.75;
            Assert.InRange(image.Crop.Value.Top.ToPoints(), vertical - .003, vertical + .003);
            Assert.InRange(image.Crop.Value.Bottom.ToPoints(), vertical - .003, vertical + .003);
        }
        Assert.Equal(120, images[0].Bounds.Width.ToPoints());
        Assert.Equal(240, images[1].Bounds.Width.ToPoints());
        using PowerPointPresentation roundTrip = reopened.ToPowerPointPresentation();
        foreach (PowerPointPicture picture in roundTrip.Slides[0].Pictures) AssertCrop(picture, .25, .125, .125, .125);
        AssertNoCropLoss(conversion.Report);
    }

    [Theory]
    [InlineData(96, 96, 30, 15, 7.5)]
    [InlineData(192, 96, 15, 7.5, 7.5)]
    [InlineData(192, 192, 15, 7.5, 3.75)]
    public void OdpCropUsesIntrinsicPixelsAndDensityAfterResizing(double dpiX, double dpiY,
        double left, double right, double vertical) {
        byte[] bytes = Grid(dpiX, dpiY);
        OdpPresentation source = OdpPresentation.Create();
        OdpSlide slide = source.AddSlide();
        foreach (double width in new[] { 120D, 240D }) {
            OdpImage image = slide.AddImage(bytes, "grid.png", Bounds(width));
            image.Crop = Insets(vertical, right, vertical, left);
        }

        OdfConversionResult<PowerPointPresentation> conversion = source.ToPowerPointPresentationResult();
        using PowerPointPresentation target = conversion.Value;
        byte[] pptx = target.ToBytes();
        using PowerPointPresentation reopened = PowerPointPresentation.Load(new MemoryStream(pptx));
        PowerPointPicture[] pictures = reopened.Slides[0].Pictures.ToArray();
        Assert.Equal(2, pictures.Length);
        foreach (PowerPointPicture picture in pictures) {
            AssertCrop(picture, .25, .125, .125, .125);
            Assert.Equal(bytes, picture.GetImageBytes());
            Assert.Equal(20, picture.LeftPoints);
            Assert.Equal(30, picture.TopPoints);
            Assert.Equal(90, picture.HeightPoints);
        }
        Assert.Equal(120, pictures[0].WidthPoints);
        Assert.Equal(240, pictures[1].WidthPoints);
        AssertNoCropLoss(conversion.Report);
    }

    [Theory]
    [InlineData("negative")]
    [InlineData("collapsed")]
    [InlineData("rounded-collapse")]
    [InlineData("vector")]
    [InlineData("malformed")]
    public void UnsupportedOdpCropRetainsImageAndBoundsWithExplicitLoss(string kind) {
        byte[] bytes = kind == "vector" ? System.Text.Encoding.UTF8.GetBytes(
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"160\" height=\"80\"><rect width=\"160\" height=\"80\" fill=\"red\"/></svg>") : Grid();
        OdpPresentation source = OdpPresentation.Create();
        OdpImage image = source.AddSlide().AddImage(bytes, kind == "vector" ? "grid.svg" : "grid.png", Bounds(240));
        image.Crop = kind switch {
            "negative" => Insets(0, 15, 0, -30),
            "collapsed" => Insets(0, 90, 0, 90),
            "rounded-collapse" => Insets(0, 0, 0, OfficeImageReader.Identify(bytes).Width * 72D /
                OfficeImageReader.Identify(bytes).DpiX * .9999999),
            _ => Insets(0, 0, 0, 30)
        };
        if (kind == "malformed") image.EnsureGraphicStyle().SetProperty(OdfNamespaces.Style + "graphic-properties",
            OdfNamespaces.Fo + "clip", "rect(0pt 0pt 0pt 10%)");

        OdfConversionResult<PowerPointPresentation> conversion = source.ToPowerPointPresentationResult();
        using PowerPointPresentation target = conversion.Value;
        PowerPointPicture picture = Assert.Single(target.Slides[0].Pictures);
        Assert.Equal(bytes, picture.GetImageBytes());
        Assert.Equal(240, picture.WidthPoints);
        Assert.Equal(90, picture.HeightPoints);
        AssertCrop(picture, 0, 0, 0, 0);
        AssertCropLoss(conversion.Report);
        Assert.Throws<OdfConversionLossException>(() => source.ToPowerPointPresentationResult(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void CropLargerThanOnePixelIntrinsicImageIsExplicitLoss() {
        byte[] bytes = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        OdpPresentation source = OdpPresentation.Create();
        OdpImage image = source.AddSlide().AddImage(bytes, "pixel.png", OdfRect.FromCentimeters(1, 1, 2, 2));
        image.Crop = OdfInsets.FromCentimeters(0, 0, 0, .2);
        OdfConversionResult<PowerPointPresentation> conversion = source.ToPowerPointPresentationResult();
        using PowerPointPresentation target = conversion.Value;
        AssertCrop(Assert.Single(target.Slides[0].Pictures), 0, 0, 0, 0);
        AssertCropLoss(conversion.Report);
    }

    [Theory]
    [InlineData("-10000", "12500")]
    [InlineData("60000", "60000")]
    [InlineData("100000", "0")]
    [InlineData("10%", "0")]
    [InlineData("", "0")]
    public void UnsupportedPowerPointCropDoesNotBecomeAClampedOdpCrop(string left, string right) {
        byte[] bytes = Grid();
        using PowerPointPresentation source = PowerPointPresentation.Create();
        PowerPointSlide slide = source.AddSlide();
        using var stream = new MemoryStream(bytes);
        slide.AddPicture(stream, OfficeImageFormat.Png, PowerPointLayoutBox.FromPoints(20, 30, 240, 90));
        Picture raw = source.OpenXmlDocument.PresentationPart!.SlideParts.Single().Slide!.Descendants<Picture>().Single();
        raw.BlipFill!.SourceRectangle = new A.SourceRectangle();
        raw.BlipFill.SourceRectangle.SetAttribute(new DocumentFormat.OpenXml.OpenXmlAttribute("", "l", "", left));
        raw.BlipFill.SourceRectangle.SetAttribute(new DocumentFormat.OpenXml.OpenXmlAttribute("", "r", "", right));
        OdfConversionResult<OdpPresentation> conversion = source.ToOpenDocumentResult();
        OdpImage image = Assert.Single(conversion.Value.Slides[0].Shapes.OfType<OdpImage>());
        Assert.Null(image.Crop);
        Assert.Equal(bytes, image.GetImageBytes());
        Assert.Equal(240, image.Bounds.Width.ToPoints());
        AssertCropLoss(conversion.Report);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new PowerPointOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void CroppedPowerPointVectorRetainsItsPayloadAndReportsCropLoss() {
        byte[] bytes = System.Text.Encoding.UTF8.GetBytes(
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"160\" height=\"80\"><rect width=\"160\" height=\"80\" fill=\"red\"/></svg>");
        using PowerPointPresentation source = PowerPointPresentation.Create();
        using var stream = new MemoryStream(bytes);
        source.AddSlide().AddPicture(stream, OfficeImageFormat.Svg, PowerPointLayoutBox.FromPoints(20, 30, 240, 90)).Crop(25, 0, 0, 0);
        OdfConversionResult<OdpPresentation> conversion = source.ToOpenDocumentResult();
        OdpImage image = Assert.Single(conversion.Value.Slides[0].Shapes.OfType<OdpImage>());
        Assert.Null(image.Crop);
        Assert.Equal(bytes, image.GetImageBytes());
        Assert.Equal(240, image.Bounds.Width.ToPoints());
        AssertCropLoss(conversion.Report);
    }

    [Fact]
    public void SourceCropMappingRequiresKnownIntrinsicDimensions() {
        var unknownSize = new OfficeImageInfo(OfficeImageFormat.Png, 0, 0);
        Assert.Throws<NotSupportedException>(() => OdfImageCropCodec.ToSourceFractions(Insets(0, 0, 0, 1), unknownSize));
        Assert.Throws<NotSupportedException>(() => OdfImageCropCodec.FromSourceFractions(unknownSize, .25, 0, 0, 0));
    }

    [Fact]
    public void PowerPointCropThatCollapsesAtOdfLengthPrecisionIsExplicitLoss() {
        byte[] bytes = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.Red));
        using PowerPointPresentation source = PowerPointPresentation.Create();
        using var stream = new MemoryStream(bytes);
        source.AddSlide().AddPicture(stream, OfficeImageFormat.Png, PowerPointLayoutBox.FromPoints(20, 30, 240, 90)).Crop(99.999, 0, 0, 0);
        OdfConversionResult<OdpPresentation> conversion = source.ToOpenDocumentResult();
        OdpImage image = Assert.Single(conversion.Value.Slides[0].Shapes.OfType<OdpImage>());
        Assert.Null(image.Crop);
        Assert.Equal(bytes, image.GetImageBytes());
        AssertCropLoss(conversion.Report);
    }

    private static void AssertCrop(PowerPointPicture picture, double left, double top, double right, double bottom) {
        Assert.Equal(left, picture.CropLeftRatio, 4);
        Assert.Equal(top, picture.CropTopRatio, 4);
        Assert.Equal(right, picture.CropRightRatio, 4);
        Assert.Equal(bottom, picture.CropBottomRatio, 4);
    }
    private static void AssertCropLoss(OdfConversionReport report) => Assert.Contains(report.ForFeature("shape-appearance"),
        mapping => mapping.Status == OdfConversionMappingStatus.Unsupported);
    private static void AssertNoCropLoss(OdfConversionReport report) => Assert.DoesNotContain(report.ForFeature("shape-appearance"),
        mapping => mapping.Status == OdfConversionMappingStatus.Unsupported);
    private static OdfRect Bounds(double width) => new(OdfLength.Points(20), OdfLength.Points(30), OdfLength.Points(width), OdfLength.Points(90));
    private static OdfInsets Insets(double top, double right, double bottom, double left) => new(OdfLength.Points(top),
        OdfLength.Points(right), OdfLength.Points(bottom), OdfLength.Points(left));
    private static byte[] Grid(double dpiX = 96, double dpiY = 96) {
        var raster = new OfficeRasterImage(160, 80, OfficeColor.Red);
        return OfficePngWriter.Encode(raster, new OfficePngEncodeOptions { DpiX = dpiX, DpiY = dpiY });
    }
}
