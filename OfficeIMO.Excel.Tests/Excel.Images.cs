using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using Xunit;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_ExcelImage_FromFile_ScalesAndSetsMetadata() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelImage.FromFile.Scale.xlsx");
            string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
            OfficeImageInfo info = OfficeImageReader.Identify(imagePath);

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Images");
                ExcelImage image = sheet.AddImageFromFile(2, 3, imagePath, scalePercent: 20, offsetX: OfficeIMO.OfficeLength.Pixels(5), offsetY: OfficeIMO.OfficeLength.Pixels(7),
                    name: "ScaledLogo", altText: "Evotec logo", title: "Logo", rotationDegrees: 15);

                Assert.Equal("ScaledLogo", image.Name);
                Assert.Equal("Evotec logo", image.Description);
                Assert.Equal("Logo", image.Title);
                Assert.Equal((OfficeIMO.OfficeLength.Pixels(info.Width) * 0.20).ToPixels(), image.Width.ToPixels());
                Assert.Equal((OfficeIMO.OfficeLength.Pixels(info.Height) * 0.20).ToPixels(), image.Height.ToPixels());
                Assert.Equal(15, image.RotationDegrees);

                document.Save();
            }

            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false);
            var drawingPart = spreadsheet.WorkbookPart!.WorksheetParts.First().DrawingsPart!;
            Xdr.OneCellAnchor anchor = drawingPart.WorksheetDrawing!.Elements<Xdr.OneCellAnchor>().Single();
            Assert.Equal("2", anchor.FromMarker!.ColumnId!.Text);
            Assert.Equal("1", anchor.FromMarker.RowId!.Text);
            Assert.Equal((5 * 9525).ToString(), anchor.FromMarker.ColumnOffset!.Text);
            Assert.Equal((7 * 9525).ToString(), anchor.FromMarker.RowOffset!.Text);

            var picture = anchor.GetFirstChild<Xdr.Picture>()!;
            Assert.Equal("ScaledLogo", picture.NonVisualPictureProperties!.NonVisualDrawingProperties!.Name!.Value);
            Assert.Equal("Evotec logo", picture.NonVisualPictureProperties.NonVisualDrawingProperties.Description!.Value);
            Assert.Equal("Logo", picture.NonVisualPictureProperties.NonVisualDrawingProperties.Title!.Value);
            Assert.Equal(15 * 60000, picture.ShapeProperties!.Transform2D!.Rotation!.Value);
        }

        [Fact]
        public void Test_ExcelImage_ToRange_UsesTwoCellAnchor() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelImage.RangeAnchor.xlsx");
            string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Images");
                ExcelImage image = sheet.AddImageFromFileToRange("A1:C15", imagePath, name: "RangeLogo", altText: "Logo pinned to report header",
                    title: "Pinned logo", placement: ExcelImagePlacement.MoveAndSize);
                Assert.True(image.Width.ToPixels() > 0);
                Assert.True(image.Height.ToPixels() > 0);
                document.Save();
            }

            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false);
            var drawingPart = spreadsheet.WorkbookPart!.WorksheetParts.First().DrawingsPart!;
            Xdr.TwoCellAnchor anchor = drawingPart.WorksheetDrawing!.Elements<Xdr.TwoCellAnchor>().Single();

            Assert.Equal(Xdr.EditAsValues.TwoCell, anchor.EditAs!.Value);
            Assert.Equal("0", anchor.FromMarker!.ColumnId!.Text);
            Assert.Equal("0", anchor.FromMarker.RowId!.Text);
            Assert.Equal("3", anchor.ToMarker!.ColumnId!.Text);
            Assert.Equal("15", anchor.ToMarker.RowId!.Text);

            var picture = anchor.GetFirstChild<Xdr.Picture>()!;
            Assert.Equal("RangeLogo", picture.NonVisualPictureProperties!.NonVisualDrawingProperties!.Name!.Value);
            Assert.Equal("Logo pinned to report header", picture.NonVisualPictureProperties.NonVisualDrawingProperties.Description!.Value);
            Assert.Equal("Pinned logo", picture.NonVisualPictureProperties.NonVisualDrawingProperties.Title!.Value);
            Assert.True(picture.ShapeProperties!.Transform2D!.Extents!.Cx!.Value > 0);
            Assert.True(picture.ShapeProperties.Transform2D.Extents.Cy!.Value > 0);
        }

        [Fact]
        public void Test_ExcelImage_ToRange_UsesCustomCellDimensions() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelImage.RangeAnchor.CustomDimensions.xlsx");
            string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Images");
                sheet.SetColumnWidth(1, ExcelColumnWidth.Characters(20));
                sheet.SetColumnWidth(2, ExcelColumnWidth.Characters(18));
                sheet.SetRowHeight(1, ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(30)));
                sheet.SetRowHeight(2, ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(24)));

                ExcelImage image = sheet.AddImageFromFileToRange("A1:B2", imagePath, name: "CustomSizedRange");

                Assert.True(image.Width.ToPixels() > 128);
                Assert.True(image.Height.ToPixels() > 40);
                document.Save();
            }

            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false);
            var picture = spreadsheet.WorkbookPart!.WorksheetParts.First().DrawingsPart!.WorksheetDrawing!
                .Elements<Xdr.TwoCellAnchor>()
                .Single()
                .GetFirstChild<Xdr.Picture>()!;
            Assert.True(picture.ShapeProperties!.Transform2D!.Extents!.Cx!.Value > 128L * 9525L);
            Assert.True(picture.ShapeProperties.Transform2D.Extents.Cy!.Value > 40L * 9525L);
        }

        [Fact]
        public void Test_ExcelImage_ToRange_ScalesTwoCellEndMarker() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelImage.RangeAnchor.ScaleMarker.xlsx");
            string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Images");
                ExcelImage image = sheet.AddImageFromFileToRange("A1:C15", imagePath, name: "ScaledRange");
                double originalWidth = image.Width.ToPixels();
                double originalHeight = image.Height.ToPixels();

                image.SetSizePercent(50);

                Assert.True(image.Width.ToPixels() < originalWidth);
                Assert.True(image.Height.ToPixels() < originalHeight);
                document.Save();
            }

            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false);
            Xdr.TwoCellAnchor anchor = spreadsheet.WorkbookPart!.WorksheetParts.First().DrawingsPart!.WorksheetDrawing!
                .Elements<Xdr.TwoCellAnchor>()
                .Single();
            Assert.Equal("0", anchor.FromMarker!.ColumnId!.Text);
            Assert.Equal("0", anchor.FromMarker.RowId!.Text);
            Assert.NotEqual("3", anchor.ToMarker!.ColumnId!.Text);
            Assert.NotEqual("15", anchor.ToMarker.RowId!.Text);
        }

        [Fact]
        public void Test_ExcelImage_TwoCellSizeReflectsCurrentCellDimensions() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelImage.RangeAnchor.DynamicDimensions.xlsx");
            string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");

            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddSheet("Images");
            ExcelImage image = sheet.AddImageFromFileToRange("A1:B2", imagePath, name: "DynamicRange");
            double originalWidth = image.Width.ToPixels();
            double originalHeight = image.Height.ToPixels();

            sheet.SetColumnWidth(1, ExcelColumnWidth.Characters(30));
            sheet.SetColumnWidth(2, ExcelColumnWidth.Characters(30));
            sheet.SetRowHeight(1, ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(45)));
            sheet.SetRowHeight(2, ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(45)));

            Assert.True(image.Width.ToPixels() > originalWidth);
            Assert.True(image.Height.ToPixels() > originalHeight);
        }

        [Fact]
        public void Test_ExcelImage_TwoCellSizeIgnoresHiddenColumns() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelImage.RangeAnchor.HiddenColumns.xlsx");
            string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");

            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddSheet("Images");
            sheet.SetColumnWidth(1, ExcelColumnWidth.Characters(20));
            sheet.SetColumnWidth(2, ExcelColumnWidth.Characters(20));
            sheet.SetColumnWidth(3, ExcelColumnWidth.Characters(20));
            ExcelImage image = sheet.AddImageFromFileToRange("A1:C1", imagePath, name: "HiddenColumnRange");
            double visibleWidth = image.Width.ToPixels();

            sheet.SetColumnHidden(2, true);

            Assert.True(image.Width.ToPixels() < visibleWidth);
        }

        [Fact]
        public void Test_ExcelImage_TwoCellSizeIgnoresHiddenRows() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelImage.RangeAnchor.HiddenRows.xlsx");
            string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");

            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddSheet("Images");
            sheet.SetRowHeight(1, ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(15)));
            sheet.SetRowHeight(2, ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(15)));
            sheet.SetRowHeight(3, ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(15)));
            ExcelImage image = sheet.AddImageFromFileToRange("A1:A3", imagePath, name: "HiddenRowRange");

            sheet.SetRowHidden(2, true);

            Assert.Equal(40, image.Height.ToPixels());
        }

        [Fact]
        public void Test_ExcelImage_MoveOnlyRangeKeepsStoredSizeWhenCellsResize() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelImage.RangeAnchor.MoveOnlySize.xlsx");
            string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");

            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddSheet("Images");
            ExcelImage image = sheet.AddImageFromFileToRange("A1:B2", imagePath, name: "MoveOnlyRange", placement: ExcelImagePlacement.MoveOnly);
            double originalWidth = image.Width.ToPixels();
            double originalHeight = image.Height.ToPixels();

            sheet.SetColumnWidth(1, ExcelColumnWidth.Characters(30));
            sheet.SetColumnWidth(2, ExcelColumnWidth.Characters(30));
            sheet.SetRowHeight(1, ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(45)));
            sheet.SetRowHeight(2, ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(45)));

            Assert.Equal(originalWidth, image.Width.ToPixels());
            Assert.Equal(originalHeight, image.Height.ToPixels());
        }

        [Fact]
        public void Test_ExcelImage_MoveOnlyRangeKeepsEndMarkerWhenTemplateRowsExpandInsideRange() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelImage.RangeAnchor.MoveOnlyTemplateRows.xlsx");
            string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Images");
                sheet.CellAt(3, 1).SetValue("{{Name}}");
                sheet.AddImageFromFileToRange("A2:B5", imagePath, name: "MoveOnlyTemplateRange", placement: ExcelImagePlacement.MoveOnly);

                sheet.ApplyTemplateRows(3, new[] {
                    new Dictionary<string, object?> { ["Name"] = "First" },
                    new Dictionary<string, object?> { ["Name"] = "Second" }
                });
                document.Save();
            }

            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false);
            Xdr.TwoCellAnchor anchor = spreadsheet.WorkbookPart!.WorksheetParts.First().DrawingsPart!.WorksheetDrawing!
                .Elements<Xdr.TwoCellAnchor>()
                .Single();
            Assert.Equal(Xdr.EditAsValues.OneCell, anchor.EditAs!.Value);
            Assert.Equal("1", anchor.FromMarker!.RowId!.Text);
            Assert.Equal("5", anchor.ToMarker!.RowId!.Text);
        }

        [Fact]
        public void Test_ExcelImage_MoveAndSizeRangeKeepsImageWhenFirstRowIsRemoved() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelImage.RangeAnchor.MoveAndSizePartialDelete.xlsx");
            string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Images");
                sheet.CellAt(2, 1).SetValue("Remove");
                sheet.AddImageFromFileToRange("A2:B5", imagePath, name: "MoveAndSizePartialDelete", placement: ExcelImagePlacement.MoveAndSize);

                sheet.RemoveTemplateOptionalRows(2, 1);
                document.Save();
            }

            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false);
            Xdr.TwoCellAnchor anchor = spreadsheet.WorkbookPart!.WorksheetParts.First().DrawingsPart!.WorksheetDrawing!
                .Elements<Xdr.TwoCellAnchor>()
                .Single();
            Assert.Equal(Xdr.EditAsValues.TwoCell, anchor.EditAs!.Value);
            Assert.Equal("1", anchor.FromMarker!.RowId!.Text);
            Assert.Equal("4", anchor.ToMarker!.RowId!.Text);
        }

        [Fact]
        public void Test_ExcelImage_MoveAndSizeRangeDoesNotGrowWhenRowsAreInsertedBelowRange() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelImage.RangeAnchor.MoveAndSizeInsertBelow.xlsx");
            string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Images");
                sheet.CellAt(5, 1).SetValue("{{Name}}");
                sheet.AddImageFromFileToRange("A2:B5", imagePath, name: "MoveAndSizeInsertBelow", placement: ExcelImagePlacement.MoveAndSize);

                sheet.ApplyTemplateRows(5, new[] {
                    new Dictionary<string, object?> { ["Name"] = "First" },
                    new Dictionary<string, object?> { ["Name"] = "Second" }
                });
                document.Save();
            }

            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false);
            Xdr.TwoCellAnchor anchor = spreadsheet.WorkbookPart!.WorksheetParts.First().DrawingsPart!.WorksheetDrawing!
                .Elements<Xdr.TwoCellAnchor>()
                .Single();
            Assert.Equal(Xdr.EditAsValues.TwoCell, anchor.EditAs!.Value);
            Assert.Equal("1", anchor.FromMarker!.RowId!.Text);
            Assert.Equal("5", anchor.ToMarker!.RowId!.Text);
        }

        [Fact]
        public void Test_ExcelImage_TwoCellSizeUsesWorkbookDefaultFontWidth() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelImage.RangeAnchor.DefaultFontWidth.xlsx");
            string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");

            using ExcelDocument document = ExcelDocument.Create(filePath);
            SetDefaultWorkbookFont(document._spreadSheetDocument, "Consolas", 11);
            ExcelSheet sheet = document.AddSheet("Images");

            ExcelImage image = sheet.AddImageFromFileToRange("A1:B1", imagePath, name: "FontSizedRange");
            double regularWidth = image.Width.ToPixels();
            SetDefaultWorkbookFont(document._spreadSheetDocument, "Consolas", 44);
            ExcelImage large = sheet.AddImageFromFileToRange("A2:B2", imagePath, name: "LargeFontSizedRange");
            OfficeTrueTypeFont? font = OfficeTrueTypeFont.TryLoadFontFamilyForText("Consolas", OfficeFontStyle.Regular, "0123456789", out _);
            if (font?.MatchesFamilyName("Consolas") == true) {
                Assert.True(large.Width.ToPixels() > regularWidth);
            } else {
                // An unavailable authored face keeps the portable seven-pixel
                // digit fallback at either font size.
                Assert.Equal(128, regularWidth);
                Assert.Equal(regularWidth, large.Width.ToPixels());
            }
        }

        private static void SetDefaultWorkbookFont(SpreadsheetDocument document, string fontName, double fontSize) {
            WorkbookStylesPart stylesPart = document.WorkbookPart!.WorkbookStylesPart ?? document.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            Stylesheet stylesheet = stylesPart.Stylesheet ??= new Stylesheet(new Fonts(new Font()), new Fills(new Fill()), new Borders(new Border()), new CellFormats(new CellFormat()));
            stylesheet.Fonts ??= new Fonts(new Font());
            Font font = stylesheet.Fonts.Elements<Font>().FirstOrDefault() ?? new Font();
            if (!stylesheet.Fonts.Elements<Font>().Any()) {
                stylesheet.Fonts.Append(font);
            }

            font.FontName = new FontName { Val = fontName };
            font.FontSize = new FontSize { Val = fontSize };
            stylesheet.Fonts.Count = (uint)stylesheet.Fonts.Count();
            stylesheet.Save();
        }

        [Fact]
        public void Test_ExcelImage_FromFile_MapsTiffContentType() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelImage.FromFile.Tiff.xlsx");
            string imagePath = Path.Combine(_directoryWithImages, "saturn.tif");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Images");
                ExcelImage image = sheet.AddImageFromFile(2, 2, imagePath, width: OfficeIMO.OfficeLength.Pixels(32), height: OfficeIMO.OfficeLength.Pixels(32));

                Assert.Equal("image/tiff", image.ContentType);
                document.Save();
            }

            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false);
            var imagePart = spreadsheet.WorkbookPart!.WorksheetParts.First().DrawingsPart!.ImageParts.Single();
            Assert.Equal("image/tiff", imagePart.ContentType);
        }

        [Fact]
        public void Test_ExcelImage_FromFile_RejectsBytesThatOnlyMatchTheExtension() {
            string workbookPath = Path.Combine(_directoryWithFiles, "ExcelImage.InvalidExtension.xlsx");
            string imagePath = Path.Combine(_directoryWithFiles, "not-an-image.png");
            File.WriteAllText(imagePath, "not an image");

            using ExcelDocument document = ExcelDocument.Create(workbookPath);
            ExcelSheet sheet = document.AddSheet("Images");

            Assert.Throws<ArgumentException>(() =>
                sheet.AddImageFromFile(2, 2, imagePath, width: OfficeIMO.OfficeLength.Pixels(32), height: OfficeIMO.OfficeLength.Pixels(32)));
            Assert.Empty(sheet.Images);
        }

        [Fact]
        public void Test_ExcelImage_FromFile_RejectsRecognizableButTruncatedGif() {
            string workbookPath = Path.Combine(_directoryWithFiles, "ExcelImage.TruncatedGif.xlsx");
            string imagePath = Path.Combine(_directoryWithFiles, "truncated.gif");
            File.WriteAllBytes(imagePath, new byte[] {
                (byte)'G', (byte)'I', (byte)'F', (byte)'8', (byte)'9', (byte)'a',
                1, 0, 1, 0, 0, 0, 0
            });

            using ExcelDocument document = ExcelDocument.Create(workbookPath);
            ExcelSheet sheet = document.AddSheet("Images");

            Assert.Throws<ArgumentException>(() =>
                sheet.AddImageFromFile(2, 2, imagePath, width: OfficeIMO.OfficeLength.Pixels(32), height: OfficeIMO.OfficeLength.Pixels(32)));
            Assert.Empty(sheet.Images);
        }

        [Fact]
        public void Test_ExcelImage_FromFile_RejectsJpegWithoutScanData() {
            string workbookPath = Path.Combine(_directoryWithFiles, "ExcelImage.MarkerOnlyJpeg.xlsx");
            string imagePath = Path.Combine(_directoryWithFiles, "marker-only.jpg");
            File.WriteAllBytes(imagePath, new byte[] {
                0xFF, 0xD8,
                0xFF, 0xC0, 0x00, 0x0B, 0x08, 0x00, 0x01, 0x00, 0x01, 0x01, 0x01, 0x11, 0x00,
                0xFF, 0xDA, 0x00, 0x08, 0x01, 0x01, 0x00, 0x00, 0x3F, 0x00,
                0xFF, 0xD9
            });

            using ExcelDocument document = ExcelDocument.Create(workbookPath);
            ExcelSheet sheet = document.AddSheet("Images");

            Assert.Throws<ArgumentException>(() =>
                sheet.AddImageFromFile(2, 2, imagePath, width: OfficeIMO.OfficeLength.Pixels(32), height: OfficeIMO.OfficeLength.Pixels(32)));
            Assert.Empty(sheet.Images);
        }

        [Theory]
        [InlineData("Sample.svg", "image/svg+xml")]
        [InlineData("sample.emf", "image/x-emf")]
        public void Test_ExcelImage_FromFile_MapsDetectedVectorContentType(string imageName, string expectedContentType) {
            string filePath = Path.Combine(_directoryWithFiles, $"ExcelImage.FromFile.{imageName}.xlsx");
            string imagePath = Path.Combine(_directoryWithImages, imageName);
            if (string.Equals(Path.GetExtension(imageName), ".emf", StringComparison.OrdinalIgnoreCase)) {
                imagePath = Path.Combine(_directoryWithFiles, imageName);
                File.WriteAllBytes(imagePath, CreateCompleteImageTestEmf(2, 2));
            }

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Images");
                ExcelImage image = sheet.AddImageFromFile(2, 2, imagePath, width: OfficeIMO.OfficeLength.Pixels(32), height: OfficeIMO.OfficeLength.Pixels(32));

                Assert.Equal(expectedContentType, image.ContentType);
                document.Save();
            }

            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false);
            var imagePart = spreadsheet.WorkbookPart!.WorksheetParts.First().DrawingsPart!.ImageParts.Single();
            Assert.Equal(expectedContentType, imagePart.ContentType);
        }

        private static byte[] CreateCompleteImageTestEmf(int width, int height) {
            var bytes = new byte[108];
            WriteImageTestInt32LittleEndian(bytes, 0, 1);
            WriteImageTestInt32LittleEndian(bytes, 4, 88);
            WriteImageTestInt32LittleEndian(bytes, 16, width);
            WriteImageTestInt32LittleEndian(bytes, 20, height);
            WriteImageTestInt32LittleEndian(bytes, 40, 0x464D4520);
            WriteImageTestInt32LittleEndian(bytes, 44, 0x00010000);
            WriteImageTestInt32LittleEndian(bytes, 48, bytes.Length);
            WriteImageTestInt32LittleEndian(bytes, 52, 2);
            bytes[56] = 1;
            WriteImageTestInt32LittleEndian(bytes, 88, 14);
            WriteImageTestInt32LittleEndian(bytes, 92, 20);
            WriteImageTestInt32LittleEndian(bytes, 104, 20);
            return bytes;
        }

        private static void WriteImageTestInt32LittleEndian(byte[] bytes, int offset, int value) {
            bytes[offset] = (byte)value;
            bytes[offset + 1] = (byte)(value >> 8);
            bytes[offset + 2] = (byte)(value >> 16);
            bytes[offset + 3] = (byte)(value >> 24);
        }

        [Fact]
        public async Task Test_ExcelImage_FromUrl_PropagatesCancellation() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelImage.FromUrl.ScaleOnly.xlsx");

            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddSheet("Images");

            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
                sheet.AddImageFromUrlAsync(2, 2, "http://127.0.0.1:1/not-found.png", scalePercent: 25,
                    cancellationToken: cancellation.Token));
        }
    }
}
