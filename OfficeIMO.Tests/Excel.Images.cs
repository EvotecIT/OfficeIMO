using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
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
                sheet.AddImageFromFileToRange("A1:C15", imagePath, name: "RangeLogo", altText: "Logo pinned to report header",
                    title: "Pinned logo", placement: ExcelImagePlacement.MoveAndSize);
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
        }
    }
}
