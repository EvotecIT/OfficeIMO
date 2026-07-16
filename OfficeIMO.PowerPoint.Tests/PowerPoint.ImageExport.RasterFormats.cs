using System.IO;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class PowerPointRasterFormatExportTests {
        [Theory]
        [InlineData(OfficeImageExportFormat.Jpeg, OfficeImageFormat.Jpeg)]
        [InlineData(OfficeImageExportFormat.Tiff, OfficeImageFormat.Tiff)]
        [InlineData(OfficeImageExportFormat.Webp, OfficeImageFormat.Webp)]
        public void PowerPointSlide_ExportsSharedRasterFormats(
            OfficeImageExportFormat format,
            OfficeImageFormat expected) {
            using var package = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(package);
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
        public void PowerPointPresentation_BatchSaveUsesWebpExtension() {
            using var package = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(package);
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
}
