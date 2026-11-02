using System.Linq;
using OfficeIMO.Excel;
using Xunit;
using S = DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_RangeMutations_CopyAndTransposeTranslateRelativeFormulas() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(1, 1).SetValue(1);
            sheet.Cell(1, 2).SetFormula("A1+1");
            sheet.Cell(2, 1).SetValue(2);
            sheet.Cell(2, 2).SetFormula("$A$1+A2");

            ExcelMutationResult copy = sheet.Range("A1:B2").CopyTo("D1");
            Assert.True(copy.PackageIsValid);
            Assert.Equal(1d, sheet.Cell(1, 4).GetValue<double>());
            Assert.Equal("D1+1", sheet.GetFormulaCells().Single(item => item.CellReference == "E1").Formula);
            Assert.Equal("$A$1+D2", sheet.GetFormulaCells().Single(item => item.CellReference == "E2").Formula);

            sheet.Range("A1:B2").TransposeTo("A4");
            Assert.Equal(1d, sheet.Cell(4, 1).GetValue<double>());
            Assert.Equal(2d, sheet.Cell(4, 2).GetValue<double>());
            Assert.Equal("A4+1", sheet.GetFormulaCells().Single(item => item.CellReference == "A5").Formula);
        }

        [Fact]
        public void Test_RangeMutations_MoveUpdatesDependentReferences() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(1, 1).SetValue(1);
            sheet.Cell(2, 2).SetValue(2);
            sheet.Cell(3, 3).SetFormula("A1+B2");

            sheet.Range("A1:B2").MoveTo("G1");

            Assert.Null(sheet.Cell(1, 1).GetValue().Value);
            Assert.Equal(1d, sheet.Cell(1, 7).GetValue<double>());
            Assert.Equal(2d, sheet.Cell(2, 8).GetValue<double>());
            Assert.Equal("G1+H2", sheet.GetFormulaCells().Single().Formula);
        }

        [Fact]
        public void Test_RangeMutations_MovePreservesOutsideReferencesInsideMovedFormula() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(1, 1).SetValue(1);
            sheet.Cell(1, 2).SetFormula("A1");

            sheet.Range("B1").MoveTo("C1");

            Assert.Equal("A1", Assert.Single(sheet.GetFormulaCells()).Formula);
        }

        [Fact]
        public void Test_RangeMutations_InsertAndDeleteCellsRoundTripReferences() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(1, 1).SetValue(1);
            sheet.Cell(1, 2).SetValue(2);
            sheet.Cell(1, 3).SetFormula("A1+B1");

            sheet.InsertCells("B1", ExcelCellShiftDirection.Right);
            Assert.Equal(2d, sheet.Cell(1, 3).GetValue<double>());
            Assert.Equal("A1+C1", sheet.GetFormulaCells().Single().Formula);

            sheet.DeleteCells("B1", ExcelCellShiftDirection.Left);
            Assert.Equal(2d, sheet.Cell(1, 2).GetValue<double>());
            Assert.Equal("A1+B1", sheet.GetFormulaCells().Single().Formula);
        }

        [Fact]
        public void Test_RangeMutations_DeleteCellsShrinksPartiallyDeletedFormulaRange() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(1, 2).SetValue(2);
            sheet.Cell(1, 3).SetValue(3);
            sheet.Cell(1, 5).SetFormula("SUM(B1:C1)");

            sheet.DeleteCells("B1", ExcelCellShiftDirection.Left);

            Assert.Equal("SUM(B1:B1)", sheet.GetFormulaCells().Single().Formula);
        }

        [Theory]
        [InlineData("B1")]
        [InlineData("C1")]
        public void Test_RangeMutations_DeleteCellsShrinksReversedFormulaRange(string deletedCell) {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(1, 2).SetValue(2);
            sheet.Cell(1, 3).SetValue(3);
            sheet.Cell(1, 5).SetFormula("SUM(C1:B1)");

            sheet.DeleteCells(deletedCell, ExcelCellShiftDirection.Left);

            Assert.Equal("SUM(B1:B1)", sheet.GetFormulaCells().Single().Formula);
        }

        [Theory]
        [InlineData("A2")]
        [InlineData("A3")]
        public void Test_RangeMutations_DeleteCellsUpShrinksReversedFormulaRange(string deletedCell) {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(2, 1).SetValue(2);
            sheet.Cell(3, 1).SetValue(3);
            sheet.Cell(1, 5).SetFormula("SUM(A3:A2)");

            sheet.DeleteCells(deletedCell, ExcelCellShiftDirection.Up);

            Assert.Equal("SUM(A2:A2)", sheet.GetFormulaCells().Single().Formula);
        }

        [Fact]
        public void Test_RangeMutations_CopyRejectsOwnedDestinationAndRevalidatesPlan() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(1, 1).SetValue("Source");
            ExcelStructuralMutationPlan plan = sheet.PlanCopyRange("A1", "D1");
            sheet.Cell(1, 4).SetValue("One");
            sheet.Cell(1, 5).SetValue("Two");
            sheet.Cell(2, 4).SetValue(1);
            sheet.Cell(2, 5).SetValue(2);
            sheet.AddTable("D1:E2", true, "Destination", ExcelTableStyle.TableStyleMedium2);

            Assert.Throws<System.InvalidOperationException>(() => plan.Apply());
            Assert.Throws<System.InvalidOperationException>(() => sheet.Range("A1").CopyTo("D1"));
            Assert.Equal("Source", sheet.Cell(1, 1).GetValue<string>());
            Assert.Equal("One", sheet.Cell(1, 4).GetValue<string>());
        }

        [Fact]
        public void Test_RangeMutations_TransposeMapsAbsoluteAxes() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(2, 2).SetFormula("$A2");

            sheet.Range("B2").TransposeTo("D4");

            Assert.Equal("D$1", sheet.GetFormulaCells().Single(item => item.CellReference == "D4").Formula);
        }

        [Fact]
        public void Test_RangeMutations_CopyPreservesImageMetadataAndTwoCellGeometry() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            ExcelImage source = sheet.AddImage("A1:B2", TinyPng, new OfficeIMO.Excel.ExcelImageOptions {
                Name = "Logo",
                Description = "Accessible logo",
                Title = "Brand",
                LockAspectRatio = false,
                Placement = ExcelImagePlacement.MoveOnly,
                RotationDegrees = 12.5
            });
            source.SetCropRatio(0.1, 0.2, 0.05, 0.15).SetFlip(true, false);

            sheet.Range("A1:B2").CopyTo("D4");

            ExcelImage copy = Assert.Single(sheet.Images, image => image.RowIndex == 4 && image.ColumnIndex == 4);
            Assert.True(copy.HasTwoCellAnchor);
            Assert.Equal(6, copy.ToRowIndex);
            Assert.Equal(6, copy.ToColumnIndex);
            Assert.Equal("Logo Copy", copy.Name);
            Assert.Equal(1, sheet.GetImage("Logo")!.RowIndex);
            Assert.NotNull(sheet.GetImage("Logo Copy"));
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

        [Fact]
        public void Test_RangeMutations_VerticalCellShiftsReparentCellsToTargetRows() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(2, 1).SetValue("shifted");
            sheet.Cell(2, 2).SetValue("fixed");

            sheet.InsertCells("A1", ExcelCellShiftDirection.Down);

            Assert.False(sheet.TryGetCellValueSnapshot(2, 1, out _));
            Assert.Equal("shifted", sheet.Cell(3, 1).GetValue<string>());
            Assert.Equal("fixed", sheet.Cell(2, 2).GetValue<string>());
            S.Cell moved = sheet.WorksheetPart.Worksheet.Descendants<S.Cell>()
                .Single(cell => cell.CellReference?.Value == "A3");
            Assert.Equal(3U, Assert.IsType<S.Row>(moved.Parent).RowIndex!.Value);

            sheet.DeleteCells("A1", ExcelCellShiftDirection.Up);

            Assert.Equal("shifted", sheet.Cell(2, 1).GetValue<string>());
            moved = sheet.WorksheetPart.Worksheet.Descendants<S.Cell>()
                .Single(cell => cell.CellReference?.Value == "A2");
            Assert.Equal(2U, Assert.IsType<S.Row>(moved.Parent).RowIndex!.Value);
        }

        [Theory]
        [InlineData(ExcelCellShiftDirection.Down, true)]
        [InlineData(ExcelCellShiftDirection.Up, false)]
        [InlineData(ExcelCellShiftDirection.Right, true)]
        [InlineData(ExcelCellShiftDirection.Left, false)]
        public void Test_RangeMutations_RejectOwnedStructuresAnywhereInShiftedBand(
            ExcelCellShiftDirection direction,
            bool inserting) {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            bool vertical = direction == ExcelCellShiftDirection.Down || direction == ExcelCellShiftDirection.Up;
            string tableRange = vertical ? "A3:B4" : "C1:D2";
            if (vertical) {
                sheet.Cell(3, 1).SetValue("A");
                sheet.Cell(3, 2).SetValue("B");
                sheet.Cell(4, 1).SetValue(1);
                sheet.Cell(4, 2).SetValue(2);
            } else {
                sheet.Cell(1, 3).SetValue("A");
                sheet.Cell(1, 4).SetValue("B");
                sheet.Cell(2, 3).SetValue(1);
                sheet.Cell(2, 4).SetValue(2);
            }
            sheet.AddTable(tableRange, true, "Owned", ExcelTableStyle.TableStyleMedium2);

            Assert.Throws<System.InvalidOperationException>(() => {
                if (inserting) sheet.InsertCells("A1", direction);
                else sheet.DeleteCells("A1", direction);
            });

            Assert.Equal(tableRange, Assert.Single(document.GetTables()).Range);
        }
    }
}
