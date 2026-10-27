using System.IO;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class PowerPointRasterFormatExportTests {
        [Fact]
        public void PowerPointSlide_PublicRasterExportHonorsEncodedByteLimitAndCancellation() {
            using var package = new MemoryStream();
            using PowerPointDocument presentation = PowerPointDocument.Create(package);
            presentation.SlideSize.SetSizePoints(160, 90);
            PowerPointSlide slide = presentation.AddSlide();
            slide.BackgroundColor = "123456";
            var options = new PowerPointImageExportOptions {
                IncludeSlideContent = false,
                MaximumTotalEncodedBytes = 8L
            };

            OfficeImageExportBatchLimitException exception = Assert.Throws<OfficeImageExportBatchLimitException>(() =>
                slide.ExportImage(OfficeImageExportFormat.Png, options));

            Assert.Equal(nameof(OfficeImageExportOptions.MaximumTotalEncodedBytes), exception.LimitName);
            Assert.Throws<OperationCanceledException>(() =>
                slide.ToImage().AsPng().Export(new System.Threading.CancellationToken(canceled: true)));
        }

        [Theory]
        [InlineData(OfficeImageExportFormat.Jpeg, OfficeImageFormat.Jpeg)]
        [InlineData(OfficeImageExportFormat.Tiff, OfficeImageFormat.Tiff)]
        [InlineData(OfficeImageExportFormat.Webp, OfficeImageFormat.Webp)]
        public void PowerPointSlide_ExportsSharedRasterFormats(
            OfficeImageExportFormat format,
            OfficeImageFormat expected) {
            using var package = new MemoryStream();
            using PowerPointDocument presentation = PowerPointDocument.Create(package);
            presentation.SlideSize.SetSizePoints(160, 90);
            PowerPointSlide slide = presentation.AddSlide();
            slide.BackgroundColor = "123456";

            OfficeImageExportResult result = slide.ExportImage(format, new PowerPointImageExportOptions {
                IncludeSlideContent = false
            });

            OfficeImageInfo info = OfficeImageReader.Identify(result.Bytes);
            Assert.Equal(format, result.Format);
            Assert.Equal(expected, info.Format);
            Assert.Equal(160, info.Width);
            Assert.Equal(90, info.Height);
        }

        [Fact]
        public void PowerPointDocument_BatchSaveUsesWebpExtension() {
            using var package = new MemoryStream();
            using PowerPointDocument presentation = PowerPointDocument.Create(package);
            presentation.SlideSize.SetSizePoints(160, 90);
            presentation.AddSlide().BackgroundColor = "123456";
            string folder = Path.Combine(Path.GetTempPath(), "OfficeIMO-" + System.Guid.NewGuid().ToString("N"));
            try {
                presentation.SaveAsImages(folder, OfficeImageExportFormat.Webp);

                Assert.True(File.Exists(Path.Combine(folder, "Slide 1.webp")));
            } finally {
                if (Directory.Exists(folder)) Directory.Delete(folder, recursive: true);
            }
        }
    }

    public partial class PowerPointImageExportTests {
        [Fact]
        public void PowerPointSlide_RendersMediaPosterInRasterAndSvg() {
            using var stream = new MemoryStream();
            using PowerPointDocument presentation = PowerPointDocument.Create(stream);
            presentation.SlideSize.SetSizePoints(160, 100);
            PowerPointSlide slide = presentation.AddSlide();
            using MemoryStream video = new(new byte[] { 0, 0, 0, 24, 102, 116, 121, 112, 109, 112, 52, 50 });
            slide.AddVideo(
                video,
                "video/mp4",
                ".mp4",
                PowerPointUnits.FromPoints(20),
                PowerPointUnits.FromPoints(20),
                PowerPointUnits.FromPoints(80),
                PowerPointUnits.FromPoints(45));

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);

            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.Equal(OfficeColor.FromRgb(31, 41, 55), rendered!.GetPixel(24, 24));
            Assert.True(CountPixelsNear(rendered, OfficeColor.FromRgb(249, 250, 251)) > 0);

            Assert.Contains("<svg", Encoding.UTF8.GetString(svg.Bytes), StringComparison.Ordinal);
            OfficeRasterImage svgPoster = ReadSvgOutput(svg);
            Assert.Equal(OfficeColor.FromRgb(31, 41, 55), svgPoster.GetPixel(24, 24));
            Assert.True(CountPixelsNear(svgPoster, OfficeColor.FromRgb(249, 250, 251)) > 0);
        }

        [Fact]
        public void PowerPointSlide_RendersBmpPicturesThroughSharedRasterDecoder() {
            using var stream = new MemoryStream();
            using PowerPointDocument presentation = PowerPointDocument.Create(stream);
            presentation.SlideSize.SetSizePoints(120, 80);
            PowerPointSlide slide = presentation.AddSlide();
            byte[] bmp = CreateBmp24(2, 2, new[] {
                OfficeColor.FromRgb(18, 52, 86), OfficeColor.FromRgb(18, 52, 86),
                OfficeColor.FromRgb(18, 52, 86), OfficeColor.FromRgb(18, 52, 86)
            });
            slide.AddImage(
                new MemoryStream(bmp),
                OfficeImageFormat.Bmp,
                PowerPointUnits.FromPoints(20),
                PowerPointUnits.FromPoints(20),
                PowerPointUnits.FromPoints(24),
                PowerPointUnits.FromPoints(18));

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);

            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.Equal(OfficeColor.FromRgb(18, 52, 86), rendered!.GetPixel(30, 28));
            Assert.Contains("<svg", Encoding.UTF8.GetString(svg.Bytes), StringComparison.Ordinal);
        }

        [Fact]
        public void PowerPointSlide_RendersTopDownBmpPicturesThroughSharedRasterDecoder() {
            using var stream = new MemoryStream();
            using PowerPointDocument presentation = PowerPointDocument.Create(stream);
            presentation.SlideSize.SetSizePoints(120, 80);
            PowerPointSlide slide = presentation.AddSlide();
            byte[] bmp = CreateBmp24(2, 2, new[] {
                OfficeColor.FromRgb(24, 96, 144), OfficeColor.FromRgb(24, 96, 144),
                OfficeColor.FromRgb(24, 96, 144), OfficeColor.FromRgb(24, 96, 144)
            }, topDown: true);
            slide.AddImage(
                new MemoryStream(bmp),
                OfficeImageFormat.Bmp,
                PowerPointUnits.FromPoints(20),
                PowerPointUnits.FromPoints(20),
                PowerPointUnits.FromPoints(24),
                PowerPointUnits.FromPoints(18));

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);

            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.Equal(OfficeColor.FromRgb(24, 96, 144), rendered!.GetPixel(30, 28));
            Assert.Contains("<svg", Encoding.UTF8.GetString(svg.Bytes), StringComparison.Ordinal);
        }

        [Fact]
        public void PowerPointSlide_RendersBmp32AlphaPicturesThroughSharedRasterDecoder() {
            using var stream = new MemoryStream();
            using PowerPointDocument presentation = PowerPointDocument.Create(stream);
            presentation.SlideSize.SetSizePoints(120, 80);
            PowerPointSlide slide = presentation.AddSlide();
            byte[] bmp = CreateBmp32(2, 2, new[] {
                OfficeColor.FromRgba(255, 0, 0, 128), OfficeColor.FromRgba(255, 0, 0, 128),
                OfficeColor.FromRgba(255, 0, 0, 128), OfficeColor.FromRgba(255, 0, 0, 128)
            });
            slide.AddImage(
                new MemoryStream(bmp),
                OfficeImageFormat.Bmp,
                PowerPointUnits.FromPoints(20),
                PowerPointUnits.FromPoints(20),
                PowerPointUnits.FromPoints(24),
                PowerPointUnits.FromPoints(18));

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png, new PowerPointImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg, new PowerPointImageExportOptions { BackgroundColor = OfficeColor.White });

            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            OfficeColor blended = rendered!.GetPixel(30, 28);
            Assert.True(blended.R >= 252, $"Expected red channel to stay near full after BMP alpha blend, got {blended.R}.");
            Assert.InRange(blended.G, 124, 130);
            Assert.InRange(blended.B, 124, 130);
            Assert.Equal(255, blended.A);
            Assert.Contains("<svg", Encoding.UTF8.GetString(svg.Bytes), StringComparison.Ordinal);
        }

        [Fact]
        public void PowerPointSlide_RendersGifPicturesThroughSharedRasterDecoder() {
            using var stream = new MemoryStream();
            using PowerPointDocument presentation = PowerPointDocument.Create(stream);
            presentation.SlideSize.SetSizePoints(120, 80);
            PowerPointSlide slide = presentation.AddSlide();
            slide.BackgroundColor = "000000";
            slide.AddImage(
                new MemoryStream(CreateSinglePixelGif()),
                OfficeImageFormat.Gif,
                PowerPointUnits.FromPoints(20),
                PowerPointUnits.FromPoints(20),
                PowerPointUnits.FromPoints(24),
                PowerPointUnits.FromPoints(18));

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);

            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.Equal(OfficeColor.White, rendered!.GetPixel(30, 28));
            Assert.Contains("<svg", Encoding.UTF8.GetString(svg.Bytes), StringComparison.Ordinal);
        }
    }
}
