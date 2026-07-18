using System.IO;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class ExcelRasterFormatExportTests {
        [Theory]
        [InlineData(OfficeImageExportFormat.Jpeg, OfficeImageFormat.Jpeg)]
        [InlineData(OfficeImageExportFormat.Tiff, OfficeImageFormat.Tiff)]
        [InlineData(OfficeImageExportFormat.Webp, OfficeImageFormat.Webp)]
        public void ExcelRange_ExportsSharedRasterFormats(
            OfficeImageExportFormat format,
            OfficeImageFormat expected) {
            string path = Path.Combine(Path.GetTempPath(), System.Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(path);
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Shared raster export");

            OfficeImageExportResult result = sheet.Range("A1:B2").ExportImage(format);

            OfficeImageInfo info = OfficeImageReader.Identify(result.Bytes);
            Assert.Equal(format, result.Format);
            Assert.Equal(expected, info.Format);
            Assert.Equal(result.Width, info.Width);
            Assert.Equal(result.Height, info.Height);
        }

        [Theory]
        [InlineData(OfficeImageExportFormat.Jpeg, OfficeImageFormat.Jpeg)]
        [InlineData(OfficeImageExportFormat.Tiff, OfficeImageFormat.Tiff)]
        [InlineData(OfficeImageExportFormat.Webp, OfficeImageFormat.Webp)]
        public void ExcelWorksheet_ComposesInPngThenEncodesRequestedRasterFormat(
            OfficeImageExportFormat format,
            OfficeImageFormat expected) {
            string path = Path.Combine(Path.GetTempPath(), System.Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(path);
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Composed raster export");

            OfficeImageExportResult result = Assert.Single(sheet.ExportImages(format, new ExcelWorksheetImageExportOptions {
                SplitByManualPageBreaks = true
            }));

            Assert.Equal(expected, OfficeImageReader.Identify(result.Bytes).Format);
            Assert.Equal(format, result.Format);
        }

        [Fact]
        public void ExcelWorksheet_CompositionUsesRequestedFormatDensityInsteadOfWorkingPngDensity() {
            string path = Path.Combine(Path.GetTempPath(), System.Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(path);
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Format-specific density");
            sheet.SetPageSetup(scale: 100, paperSize: ExcelPaperSize.Letter);

            OfficeImageExportResult result = Assert.Single(sheet.ExportImages(
                OfficeImageExportFormat.Jpeg,
                new ExcelWorksheetImageExportOptions {
                    SplitByManualPageBreaks = true,
                    RasterEncoding = new OfficeRasterEncodingOptions {
                        Png = new OfficePngEncodeOptions {
                            DpiX = 96D,
                            DpiY = 96D
                        },
                        Jpeg = new OfficeJpegEncodeOptions {
                            DpiX = 300D,
                            DpiY = 240D
                        }
                    }
                }));

            Assert.Equal(300D, result.DpiX);
            Assert.Equal(240D, result.DpiY);
        }

        [Fact]
        public void ExcelWorksheet_PageCompositionPlansSafetyReductionBeforeAllocatingFinalCanvas() {
            string path = Path.Combine(Path.GetTempPath(), System.Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(path);
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Safety-limited page");
            sheet.SetPageSetup(scale: 100, paperSize: ExcelPaperSize.Letter);

            OfficeImageExportResult result = Assert.Single(sheet.ExportImages(
                OfficeImageExportFormat.Png,
                new ExcelWorksheetImageExportOptions {
                    SplitByManualPageBreaks = true,
                    TargetDpi = 192D,
                    MaximumRasterPixels = 100_000L
                }));

            Assert.True((long)result.Width * result.Height <= 100_000L);
            Assert.Contains(
                result.Diagnostics,
                diagnostic => diagnostic.Code == OfficeImageExportDiagnosticCodes.RasterScaleReduced &&
                    diagnostic.Source == "Data!pageSetup");
            Assert.InRange(result.PhysicalWidthInches, 8.4D, 8.6D);
            Assert.InRange(result.PhysicalHeightInches, 10.9D, 11.1D);
        }

        [Fact]
        public void ExcelWorkbook_BatchSaveUsesTiffExtension() {
            string path = Path.Combine(Path.GetTempPath(), System.Guid.NewGuid() + ".xlsx");
            string folder = Path.Combine(Path.GetTempPath(), "OfficeIMO-" + System.Guid.NewGuid().ToString("N"));
            using ExcelDocument document = ExcelDocument.Create(path);
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Batch raster export");
            try {
                document.SaveAsImages(folder, OfficeImageExportFormat.Tiff);

                Assert.True(File.Exists(Path.Combine(folder, "Data.tiff")));
            } finally {
                if (Directory.Exists(folder)) Directory.Delete(folder, recursive: true);
            }
        }
    }
}
