using System.IO;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class WordRasterFormatExportTests {
        [Theory]
        [InlineData(OfficeImageExportFormat.Jpeg, OfficeImageFormat.Jpeg)]
        [InlineData(OfficeImageExportFormat.Tiff, OfficeImageFormat.Tiff)]
        [InlineData(OfficeImageExportFormat.Webp, OfficeImageFormat.Webp)]
        public void WordDocument_ExportsSharedRasterFormats(
            OfficeImageExportFormat format,
            OfficeImageFormat expected) {
            using var package = new MemoryStream();
            using WordDocument document = WordDocument.Create(package);
            document.AddParagraph("Dependency-free raster export");

            OfficeImageExportResult result = document.ExportImage(format, new WordImageExportOptions {
                RasterEncoding = new OfficeRasterEncodingOptions {
                    Tiff = new OfficeTiffEncodeOptions { Compression = OfficeTiffCompression.None }
                }
            });

            OfficeImageInfo info = OfficeImageReader.Identify(result.Bytes);
            Assert.Equal(format, result.Format);
            Assert.Equal(expected, info.Format);
            Assert.Equal(result.Width, info.Width);
            Assert.Equal(result.Height, info.Height);
        }

        [Fact]
        public void WordDocument_ThinRasterWrappersUseSharedEncoder() {
            using var package = new MemoryStream();
            using WordDocument document = WordDocument.Create(package);
            document.AddParagraph("Thin wrappers");

            using var output = new MemoryStream();
            OfficeImageExportResult saved = document.SaveAsWebp(output);

            Assert.Equal(OfficeImageExportFormat.Webp, saved.Format);
            Assert.Equal(OfficeImageFormat.Jpeg, OfficeImageReader.Identify(document.ToJpeg()).Format);
            Assert.Equal(OfficeImageFormat.Tiff, OfficeImageReader.Identify(document.ToTiff()).Format);
            Assert.Equal(OfficeImageFormat.Webp, OfficeImageReader.Identify(output.ToArray()).Format);
        }
    }
}
