using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfDocumentImageOptimizationTests {
    [Theory]
    [InlineData(false, 144, 72)]
    [InlineData(true, 144, 72)]
    public void ImageOptimizationCoversBothStretchedAndCroppedSourceAxes(bool crop, int width, int height) {
        var options = new PdfOptions { ImageOptimization = new() { Enabled = true, TargetDpi = 72, KeepOriginalWhenNotSmaller = false } };
        var document = PdfDocument.Create(options);
        document.Canvas(canvas => canvas.Image(CreateJpeg(400, 200), 40, 50, 72, crop ? 36 : 72,
            new PdfImageStyle { Fit = OfficeImageFit.Stretch, SourceCrop = crop ? new PdfImageSourceCrop(0.5, 0, 0, 0) : null }));
        var image = Assert.Single(PdfImageExtractor.ExtractImages(document.ToBytes()));
        Assert.Equal(width, image.Width);
        Assert.Equal(height, image.Height);
    }
    [Fact]
    public void ImageOptimization_RecompressesJpegWithoutReducingPixels() {
        byte[] jpeg = CreateJpeg(400, 200);
        byte[] pdf = PdfDocument.Create(new PdfOptions {
            ImageOptimization = new PdfImageOptimizationOptions {
                Enabled = true, Mode = OfficeImageOptimizationMode.Recompress,
                TargetDpi = 72, JpegQuality = 35
            }
        }).Image(jpeg, 72, 36).ToBytes();
        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));
        Assert.Equal(400, image.Width);
        Assert.Equal(200, image.Height);
        Assert.True(image.Bytes.Length < jpeg.Length);
    }

    [Fact]
    public void ImageOptimization_DownsamplesToLaidOutPlacementResolution() {
        byte[] jpeg = CreateJpeg(400, 200);
        var options = new PdfOptions {
            ImageOptimization = new PdfImageOptimizationOptions {
                Enabled = true,
                TargetDpi = 72,
                KeepOriginalWhenNotSmaller = false,
                JpegQuality = 80
            }
        };

        byte[] pdf = PdfDocument.Create(options)
            .Image(jpeg, 72, 36)
            .Image(jpeg, 72, 36)
            .ToBytes();

        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));
        Assert.Equal(72, image.Width);
        Assert.Equal(36, image.Height);
        Assert.Equal("DCTDecode", image.Filter);
        Assert.True(image.Bytes.Length < jpeg.Length);
    }

    [Fact]
    public void ImageOptimization_IsDisabledByDefault() {
        byte[] jpeg = CreateJpeg(400, 200);

        byte[] pdf = PdfDocument.Create()
            .Image(jpeg, 72, 36)
            .ToBytes();

        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));
        Assert.Equal(400, image.Width);
        Assert.Equal(200, image.Height);
        Assert.Equal(jpeg, image.Bytes);
    }

    [Fact]
    public void ImageOptimization_ConfiguredAfterExplicitPageAppliesAtWriteTime() {
        byte[] jpeg = CreateJpeg(400, 200);
        PdfDocument document = PdfDocument.Create(_ => { });
        document.Compose(compose => compose.Page(page => page
            .Content(content => content.Item(item => item.Image(jpeg, 72, 36)))));
        document.Compose(compose => compose.Settings(settings => settings.ImageOptimization = new PdfImageOptimizationOptions {
            Enabled = true,
            TargetDpi = 72,
            KeepOriginalWhenNotSmaller = false,
            JpegQuality = 80
        }));

        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(document.ToBytes()));

        Assert.Equal(72, image.Width);
        Assert.Equal(36, image.Height);
        Assert.True(image.Bytes.Length < jpeg.Length);
    }

    [Fact]
    public void ImageOptimizationOptions_AreSnapshottedAndCloned() {
        var policy = new PdfImageOptimizationOptions {
            Enabled = true,
            TargetDpi = 180,
            DownsampleThreshold = 1.25,
            JpegQuality = 78,
            ResamplingMode = OfficeRasterResamplingMode.NearestNeighbor,
            KeepOriginalWhenNotSmaller = false
        };
        var options = new PdfOptions { ImageOptimization = policy };

        policy.TargetDpi = 300;
        PdfImageOptimizationOptions readback = options.ImageOptimization!;
        readback.JpegQuality = 40;
        PdfOptions clone = options.Clone();

        Assert.Equal(180, options.ImageOptimization!.TargetDpi);
        Assert.Equal(78, options.ImageOptimization!.JpegQuality);
        Assert.Equal(1.25, clone.ImageOptimization!.DownsampleThreshold);
        Assert.Equal(OfficeRasterResamplingMode.NearestNeighbor, clone.ImageOptimization.ResamplingMode);
        Assert.False(clone.ImageOptimization.KeepOriginalWhenNotSmaller);
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfOptions {
            ImageOptimization = new() { MetadataPolicy = (OfficeImageMetadataPolicy)99 }
        });
    }

    private static byte[] CreateJpeg(int width, int height) {
        var image = new OfficeRasterImage(width, height);
        for (int y = 0; y < height; y++) {
            for (int x = 0; x < width; x++) {
                image.SetPixel(
                    x,
                    y,
                    OfficeColor.FromRgb(
                        (byte)(x % 256),
                        (byte)(y % 256),
                        (byte)((x + y) % 256)));
            }
        }

        return OfficeJpegCodec.Encode(image, new OfficeJpegEncodeOptions {
            Quality = 92,
            Subsampling = OfficeJpegSubsampling.Y444
        });
    }
}
