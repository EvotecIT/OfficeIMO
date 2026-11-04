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
        public void WordDocument_OptionsChooseTheSharedEncoder() {
            using var package = new MemoryStream();
            using WordDocument document = WordDocument.Create(package);
            document.AddParagraph("Thin wrappers");

            using var output = new MemoryStream();
            OfficeConversionReport saved = document.SaveAsImage(output, new WordImageExportOptions { Format = OfficeImageExportFormat.Webp });

            Assert.Equal("Webp", saved.Target);
            Assert.Equal(OfficeImageFormat.Jpeg, OfficeImageReader.Identify(document.ToImage(new WordImageExportOptions { Format = OfficeImageExportFormat.Jpeg })).Format);
            Assert.Equal(OfficeImageFormat.Tiff, OfficeImageReader.Identify(document.ToImage(new WordImageExportOptions { Format = OfficeImageExportFormat.Tiff })).Format);
            Assert.Equal(OfficeImageFormat.Webp, OfficeImageReader.Identify(output.ToArray()).Format);
        }
    }
}
