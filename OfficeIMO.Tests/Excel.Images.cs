using System.IO;
using System.Linq;
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
                ExcelSheet sheet = document.AddWorkSheet("Images");
                ExcelImage image = sheet.AddImageFromFile(2, 3, imagePath, scalePercent: 20, offsetXPixels: 5, offsetYPixels: 7,
                    name: "ScaledLogo", altText: "Evotec logo", title: "Logo", rotationDegrees: 15);

                Assert.Equal("ScaledLogo", image.Name);
                Assert.Equal("Evotec logo", image.Description);
                Assert.Equal("Logo", image.Title);
                Assert.Equal(Math.Max(1, (int)System.Math.Round(info.Width * 0.20)), image.WidthPixels);
                Assert.Equal(Math.Max(1, (int)System.Math.Round(info.Height * 0.20)), image.HeightPixels);
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
                ExcelSheet sheet = document.AddWorkSheet("Images");
                ExcelImage image = sheet.AddImageFromFileToRange("A1:C15", imagePath, name: "RangeLogo", altText: "Logo pinned to report header",
                    title: "Pinned logo", placement: ExcelImagePlacement.MoveAndSize);
                Assert.True(image.WidthPixels > 0);
                Assert.True(image.HeightPixels > 0);
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
                ExcelSheet sheet = document.AddWorkSheet("Images");
                sheet.SetColumnWidth(1, 20);
                sheet.SetColumnWidth(2, 18);
                sheet.SetRowHeight(1, 30);
                sheet.SetRowHeight(2, 24);

                ExcelImage image = sheet.AddImageFromFileToRange("A1:B2", imagePath, name: "CustomSizedRange");

                Assert.True(image.WidthPixels > 128);
                Assert.True(image.HeightPixels > 40);
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
                ExcelSheet sheet = document.AddWorkSheet("Images");
                ExcelImage image = sheet.AddImageFromFileToRange("A1:C15", imagePath, name: "ScaledRange");
                int originalWidth = image.WidthPixels;
                int originalHeight = image.HeightPixels;

                image.SetSizePercent(50);

                Assert.True(image.WidthPixels < originalWidth);
                Assert.True(image.HeightPixels < originalHeight);
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
            ExcelSheet sheet = document.AddWorkSheet("Images");
            ExcelImage image = sheet.AddImageFromFileToRange("A1:B2", imagePath, name: "DynamicRange");
            int originalWidth = image.WidthPixels;
            int originalHeight = image.HeightPixels;

            sheet.SetColumnWidth(1, 30);
            sheet.SetColumnWidth(2, 30);
            sheet.SetRowHeight(1, 45);
            sheet.SetRowHeight(2, 45);

            Assert.True(image.WidthPixels > originalWidth);
            Assert.True(image.HeightPixels > originalHeight);
        }

        [Fact]
        public void Test_ExcelImage_MoveOnlyRangeKeepsStoredSizeWhenCellsResize() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelImage.RangeAnchor.MoveOnlySize.xlsx");
            string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");

            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("Images");
            ExcelImage image = sheet.AddImageFromFileToRange("A1:B2", imagePath, name: "MoveOnlyRange", placement: ExcelImagePlacement.MoveOnly);
            int originalWidth = image.WidthPixels;
            int originalHeight = image.HeightPixels;

            sheet.SetColumnWidth(1, 30);
            sheet.SetColumnWidth(2, 30);
            sheet.SetRowHeight(1, 45);
            sheet.SetRowHeight(2, 45);

            Assert.Equal(originalWidth, image.WidthPixels);
            Assert.Equal(originalHeight, image.HeightPixels);
        }

        [Fact]
        public void Test_ExcelImage_TwoCellSizeUsesWorkbookDefaultFontWidth() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelImage.RangeAnchor.DefaultFontWidth.xlsx");
            string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");

            using ExcelDocument document = ExcelDocument.Create(filePath);
            SetDefaultWorkbookFont(document._spreadSheetDocument, "Consolas", 11);
            ExcelSheet sheet = document.AddWorkSheet("Images");

            ExcelImage image = sheet.AddImageFromFileToRange("A1:B1", imagePath, name: "FontSizedRange");

            Assert.True(image.WidthPixels > 128);
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
                ExcelSheet sheet = document.AddWorkSheet("Images");
                ExcelImage image = sheet.AddImageFromFile(2, 2, imagePath, widthPixels: 32, heightPixels: 32);

                Assert.Equal("image/tiff", image.ContentType);
                document.Save();
            }

            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false);
            var imagePart = spreadsheet.WorkbookPart!.WorksheetParts.First().DrawingsPart!.ImageParts.Single();
            Assert.Equal("image/tiff", imagePart.ContentType);
        }

        [Theory]
        [InlineData("Sample.svg", "image/svg+xml")]
        [InlineData("sample.emf", "image/x-emf")]
        public void Test_ExcelImage_FromFile_MapsDetectedVectorContentType(string imageName, string expectedContentType) {
            string filePath = Path.Combine(_directoryWithFiles, $"ExcelImage.FromFile.{imageName}.xlsx");
            string imagePath = Path.Combine(_directoryWithImages, imageName);

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Images");
                ExcelImage image = sheet.AddImageFromFile(2, 2, imagePath, widthPixels: 32, heightPixels: 32);

                Assert.Equal(expectedContentType, image.ContentType);
                document.Save();
            }

            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false);
            var imagePart = spreadsheet.WorkbookPart!.WorksheetParts.First().DrawingsPart!.ImageParts.Single();
            Assert.Equal(expectedContentType, imagePart.ContentType);
        }

        [Fact]
        public void Test_ExcelImage_FromUrl_AllowsScaleOnlyCallShape() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelImage.FromUrl.ScaleOnly.xlsx");

            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("Images");

            ExcelImage? image = sheet.AddImageFromUrl(2, 2, "http://127.0.0.1:1/not-found.png", scalePercent: 25);

            Assert.Null(image);
        }
    }
}
