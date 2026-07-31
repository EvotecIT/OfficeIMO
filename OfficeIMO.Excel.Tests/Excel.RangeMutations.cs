using System.Linq;
using OfficeIMO.Excel;
using Xunit;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_RangeMutations_CopyAndTransposeTranslateRelativeFormulas() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, 1);
            sheet.CellFormula(1, 2, "A1+1");
            sheet.CellValue(2, 1, 2);
            sheet.CellFormula(2, 2, "$A$1+A2");

            ExcelMutationResult copy = sheet.CopyRange("A1:B2", "D1");
            Assert.True(copy.PackageIsValid);
            Assert.Equal(1d, sheet.CellAt(1, 4).GetValue<double>());
            Assert.Equal("D1+1", sheet.GetFormulaCells().Single(item => item.CellReference == "E1").Formula);
            Assert.Equal("$A$1+D2", sheet.GetFormulaCells().Single(item => item.CellReference == "E2").Formula);

            sheet.TransposeRange("A1:B2", "A4");
            Assert.Equal(1d, sheet.CellAt(4, 1).GetValue<double>());
            Assert.Equal(2d, sheet.CellAt(4, 2).GetValue<double>());
            Assert.Equal("A4+1", sheet.GetFormulaCells().Single(item => item.CellReference == "A5").Formula);
        }

        [Fact]
        public void Test_RangeMutations_MoveUpdatesDependentReferences() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, 1);
            sheet.CellValue(2, 2, 2);
            sheet.CellFormula(3, 3, "A1+B2");

            sheet.MoveRange("A1:B2", "G1");

            Assert.Null(sheet.CellAt(1, 1).GetValue().Value);
            Assert.Equal(1d, sheet.CellAt(1, 7).GetValue<double>());
            Assert.Equal(2d, sheet.CellAt(2, 8).GetValue<double>());
            Assert.Equal("G1+H2", sheet.GetFormulaCells().Single().Formula);
        }

        [Fact]
        public void Test_RangeMutations_MovePreservesOutsideReferencesInsideMovedFormula() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, 1);
            sheet.CellFormula(1, 2, "A1");

            sheet.MoveRange("B1", "C1");

            Assert.Equal("A1", Assert.Single(sheet.GetFormulaCells()).Formula);
        }

        [Fact]
        public void Test_RangeMutations_InsertAndDeleteCellsRoundTripReferences() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, 1);
            sheet.CellValue(1, 2, 2);
            sheet.CellFormula(1, 3, "A1+B1");

            sheet.InsertCells("B1", ExcelCellShiftDirection.Right);
            Assert.Equal(2d, sheet.CellAt(1, 3).GetValue<double>());
            Assert.Equal("A1+C1", sheet.GetFormulaCells().Single().Formula);

            sheet.DeleteCells("B1", ExcelCellShiftDirection.Left);
            Assert.Equal(2d, sheet.CellAt(1, 2).GetValue<double>());
            Assert.Equal("A1+B1", sheet.GetFormulaCells().Single().Formula);
        }

        [Fact]
        public void Test_RangeMutations_DeleteCellsShrinksPartiallyDeletedFormulaRange() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 2, 2);
            sheet.CellValue(1, 3, 3);
            sheet.CellFormula(1, 5, "SUM(B1:C1)");

            sheet.DeleteCells("B1", ExcelCellShiftDirection.Left);

            Assert.Equal("SUM(B1:B1)", sheet.GetFormulaCells().Single().Formula);
        }

        [Theory]
        [InlineData("B1")]
        [InlineData("C1")]
        public void Test_RangeMutations_DeleteCellsShrinksReversedFormulaRange(string deletedCell) {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 2, 2);
            sheet.CellValue(1, 3, 3);
            sheet.CellFormula(1, 5, "SUM(C1:B1)");

            sheet.DeleteCells(deletedCell, ExcelCellShiftDirection.Left);

            Assert.Equal("SUM(B1:B1)", sheet.GetFormulaCells().Single().Formula);
        }

        [Theory]
        [InlineData("A2")]
        [InlineData("A3")]
        public void Test_RangeMutations_DeleteCellsUpShrinksReversedFormulaRange(string deletedCell) {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(2, 1, 2);
            sheet.CellValue(3, 1, 3);
            sheet.CellFormula(1, 5, "SUM(A3:A2)");

            sheet.DeleteCells(deletedCell, ExcelCellShiftDirection.Up);

            Assert.Equal("SUM(A2:A2)", sheet.GetFormulaCells().Single().Formula);
        }

        [Fact]
        public void Test_RangeMutations_CopyRejectsOwnedDestinationAndRevalidatesPlan() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Source");
            ExcelStructuralMutationPlan plan = sheet.PlanCopyRange("A1", "D1");
            sheet.CellValue(1, 4, "One");
            sheet.CellValue(1, 5, "Two");
            sheet.CellValue(2, 4, 1);
            sheet.CellValue(2, 5, 2);
            sheet.AddTable("D1:E2", true, "Destination", TableStyle.TableStyleMedium2);

            Assert.Throws<System.InvalidOperationException>(() => plan.Apply());
            Assert.Throws<System.InvalidOperationException>(() => sheet.CopyRange("A1", "D1"));
            Assert.Equal("Source", sheet.CellAt(1, 1).GetValue<string>());
            Assert.Equal("One", sheet.CellAt(1, 4).GetValue<string>());
        }

        [Fact]
        public void Test_RangeMutations_TransposeMapsAbsoluteAxes() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellFormula(2, 2, "$A2");

            sheet.TransposeRange("B2", "D4");

            Assert.Equal("D$1", sheet.GetFormulaCells().Single(item => item.CellReference == "D4").Formula);
        }

        [Fact]
        public void Test_RangeMutations_CopyPreservesImageMetadataAndTwoCellGeometry() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            ExcelImage source = sheet.AddImageToRange(
                "A1:B2",
                TinyPng,
                name: "Logo",
                altText: "Accessible logo",
                title: "Brand",
                lockAspectRatio: false,
                placement: ExcelImagePlacement.MoveOnly,
                rotationDegrees: 12.5);
            source.SetCropRatio(0.1, 0.2, 0.05, 0.15).SetFlip(true, false);

            sheet.CopyRange("A1:B2", "D4");

            ExcelImage copy = Assert.Single(sheet.Images, image => image.RowIndex == 4 && image.ColumnIndex == 4);
            Assert.True(copy.HasTwoCellAnchor);
            Assert.Equal(6, copy.ToRowIndex);
            Assert.Equal(6, copy.ToColumnIndex);
            Assert.Equal("Logo", copy.Name);
            Assert.Equal("Brand", copy.Title);
            Assert.Equal("Accessible logo", copy.Description);
            Assert.False(copy.IsAspectRatioLocked);
            Assert.Equal(12.5, copy.RotationDegrees, 3);
            Assert.True(copy.FlipHorizontal);
            Assert.False(copy.FlipVertical);
            Assert.Equal(0.1, copy.CropLeftRatio, 3);
            Assert.Equal(0.2, copy.CropTopRatio, 3);
            Assert.Equal(0.05, copy.CropRightRatio, 3);
            Assert.Equal(0.15, copy.CropBottomRatio, 3);
            Xdr.TwoCellAnchor copiedAnchor = sheet.WorksheetPart.DrawingsPart!.WorksheetDrawing!
                .Elements<Xdr.TwoCellAnchor>().Last();
            Assert.Equal(Xdr.EditAsValues.OneCell, copiedAnchor.EditAs!.Value);
        }
    }
}
