using System.Linq;
using System.Threading;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests
{
    public partial class Excel
    {
        [Fact]
        public void Test_StructuralColumns_PlanCommitRemapsCellsFormulasTablesAndSecurityRegions() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(1, 1).SetValue("Region");
            sheet.Cell(1, 2).SetValue("Amount");
            sheet.Cell(1, 3).SetValue("Tax");
            sheet.Cell(2, 1).SetValue("EU");
            sheet.Cell(2, 2).SetValue(10);
            sheet.Cell(2, 3).SetValue(2);
            sheet.AddTable("A1:C2", true, "Sales", ExcelTableStyle.TableStyleMedium2);
            sheet.Cell(4, 1).SetFormula("SUM(Sales[Amount])+C2");
            sheet.Protect(new ExcelSheetProtectionOptions { Password = "secret" });
            sheet.SetAllowedEditRange("Inputs", new[] { "B2:C2" });
            sheet.AddIgnoredErrorRegion(new[] { "B2:C2" }, ExcelIgnoredErrorKind.NumberStoredAsText);

            ExcelStructuralMutationPlan plan = sheet.PlanInsertColumns(2);
            Assert.Equal("B:B", plan.SourceRange);
            Assert.Contains(plan.Impacts, impact => impact.Category == "tables");

            ExcelMutationResult result = plan.Apply();

            Assert.True(plan.IsApplied);
            Assert.True(result.PackageIsValid);
            ExcelTableInfo table = Assert.Single(document.GetTables());
            Assert.Equal("A1:D2", table.Range);
            Assert.Equal(4, table.Columns.Count);
            Assert.Equal("SUM(Sales[Amount])+D2", sheet.GetFormulaCells().Single().Formula);
            Assert.Equal(new[] { "C2:D2" }, Assert.Single(sheet.GetAllowedEditRanges()).Ranges);
            Assert.Equal(new[] { "C2:D2" }, Assert.Single(sheet.GetIgnoredErrorRegions()).Ranges);

            sheet.DeleteColumns(2);
            Assert.Equal("A1:C2", Assert.Single(document.GetTables()).Range);
            Assert.Equal("SUM(Sales[Amount])+C2", sheet.GetFormulaCells().Single().Formula);
            Assert.Equal(new[] { "B2:C2" }, Assert.Single(sheet.GetAllowedEditRanges()).Ranges);
        }

        [Fact]
        public void Test_StructuralColumns_RejectsDeletingEntireTableBeforeMutation() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(1, 1).SetValue("A");
            sheet.Cell(1, 2).SetValue("B");
            sheet.Cell(2, 1).SetValue(1);
            sheet.Cell(2, 2).SetValue(2);
            sheet.AddTable("A1:B2", true, "DataTable", ExcelTableStyle.TableStyleMedium2);

            Assert.Throws<InvalidOperationException>(() => sheet.PlanDeleteColumns(1, 2));
            Assert.Equal("A1:B2", Assert.Single(document.GetTables()).Range);
        }

        [Fact]
        public void Test_StructuralColumns_PlanRevalidatesNewTableBeforeApply() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(1, 1).SetValue("A");
            sheet.Cell(1, 2).SetValue("B");
            sheet.Cell(2, 1).SetValue(1);
            sheet.Cell(2, 2).SetValue(2);
            ExcelStructuralMutationPlan plan = sheet.PlanDeleteColumns(1, 2);
            sheet.AddTable("A1:B2", true, "DataTable", ExcelTableStyle.TableStyleMedium2);

            Assert.Throws<System.InvalidOperationException>(() => plan.Apply());
            Assert.Equal("A1:B2", Assert.Single(document.GetTables()).Range);
        }

        [Fact]
        public void Test_StructuralColumns_ResultUsesRevalidatedAffectedCellCount() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(1, 1).SetValue("One");
            ExcelStructuralMutationPlan plan = sheet.PlanInsertColumns(1);
            sheet.Cell(2, 1).SetValue("Two");
            sheet.Cell(3, 2).SetValue("Three");

            ExcelMutationResult result = plan.Apply();

            Assert.Equal(3, result.AffectedCells);
        }

        [Fact]
        public void Test_StructuralMutation_RollbackRestoresDeletedCalculationChain() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            CalculationChainPart chainPart = document.WorkbookPartRoot.AddNewPart<CalculationChainPart>();
            chainPart.CalculationChain = new DocumentFormat.OpenXml.Spreadsheet.CalculationChain(
                new DocumentFormat.OpenXml.Spreadsheet.CalculationCell { CellReference = "A1", SheetId = 0 });
            using var cancellation = new CancellationTokenSource();

            Assert.Throws<OperationCanceledException>(() => sheet.ApplyTransactionalMutation(_ => {
                document.WorkbookPartRoot.DeletePart(document.WorkbookPartRoot.CalculationChainPart!);
                cancellation.Cancel();
            }, 0, new ExcelMutationPlanOptions(), cancellation.Token));

            DocumentFormat.OpenXml.Spreadsheet.CalculationCell restored = Assert.Single(document.WorkbookPartRoot.CalculationChainPart!
                .CalculationChain!.Elements<DocumentFormat.OpenXml.Spreadsheet.CalculationCell>());
            Assert.Equal("A1", restored.CellReference!.Value);
        }

        [Fact]
        public void Test_StructuralMutation_RollbackRestoresConnections() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            ConnectionsPart connectionsPart = document.WorkbookPartRoot.AddNewPart<ConnectionsPart>();
            connectionsPart.Connections = new S.Connections(
                new S.Connection { Id = 1U, Name = "Original" });
            using var cancellation = new CancellationTokenSource();

            Assert.Throws<OperationCanceledException>(() => sheet.ApplyTransactionalMutation(_ => {
                Assert.Single(connectionsPart.Connections.Elements<S.Connection>()).Name = "Changed";
                cancellation.Cancel();
            }, 0, new ExcelMutationPlanOptions(), cancellation.Token));

            Assert.Equal("Original", Assert.Single(connectionsPart.Connections.Elements<S.Connection>()).Name!.Value);
        }

        [Fact]
        public void Test_StructuralColumns_InsertionPreservesHeadersAndRemapsFilterColumn() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(1, 1).SetValue("A");
            sheet.Cell(1, 2).SetValue("B");
            sheet.Cell(1, 3).SetValue("C");
            sheet.Cell(2, 1).SetValue(1);
            sheet.Cell(2, 2).SetValue(2);
            sheet.Cell(2, 3).SetValue(3);
            sheet.AddTable("A1:C2", true, "DataTable", ExcelTableStyle.TableStyleMedium2);
            S.AutoFilter tableFilter = Assert.Single(sheet.WorksheetPart.TableDefinitionParts).Table!.AutoFilter!;
            tableFilter.Append(new S.FilterColumn(new S.Filters { Blank = true }) { ColumnId = 1U });

            sheet.InsertColumns(2);

            Assert.Equal(new[] { "A", "Column2", "B", "C" }, Enumerable.Range(1, 4).Select(column => sheet.Cell(1, column).GetValue<string>()));
            Assert.Equal(2U, Assert.Single(Assert.Single(sheet.GetAutoFilters(), filter => filter.IsTableFilter).Columns).ColumnOffset);
        }

        [Fact]
        public void Test_StructuralColumns_DeletionInvalidatesRemovedStructuredColumn() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(1, 1).SetValue("A");
            sheet.Cell(1, 2).SetValue("B");
            sheet.Cell(2, 1).SetValue(1);
            sheet.Cell(2, 2).SetValue(2);
            sheet.AddTable("A1:B2", true, "DataTable", ExcelTableStyle.TableStyleMedium2);
            sheet.Cell(4, 1).SetFormula("SUM(DataTable[B])");

            sheet.DeleteColumns(2);

            Assert.Equal("SUM(#REF!)", Assert.Single(sheet.GetFormulaCells()).Formula);
        }

        [Fact]
        public void Test_StructuralColumns_RejectsPivotOutputIntersection()
        {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(1, 1).SetValue("Category");
            sheet.Cell(1, 2).SetValue("Value");
            sheet.Cell(2, 1).SetValue("A");
            sheet.Cell(2, 2).SetValue(1);
            sheet.AddPivotTable("A1:B2", "E5", options: new ExcelPivotTableOptions { RowFields = new[] { "Category" }, DataFields = new[] { new ExcelPivotDataField("Value", ExcelPivotDataFunction.Sum) } });
            S.Location location = Assert.Single(sheet.WorksheetPart.PivotTableParts).PivotTableDefinition!.Location!;
            Assert.Throws<System.InvalidOperationException>(() => sheet.PlanDeleteColumns(5));
            Assert.Equal("E5", location.Reference!.Value.Split(':')[0]);
        }

        [Fact]
        public void Test_StructuralMutation_RollbackRemovesAddedImageRelationship() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            int baselineParts = sheet.WorksheetPart.Parts.Count();
            using var cancellation = new CancellationTokenSource();

            Assert.Throws<OperationCanceledException>(() => sheet.ApplyTransactionalMutation(_ => {
                sheet.AddImage(1, 1, new byte[] { 0x89, 0x50, 0x4E, 0x47 }, new OfficeIMO.Excel.ExcelImageOptions {
                    ContentType = "image/png",
                    Width = OfficeIMO.OfficeLength.Pixels(8),
                    Height = OfficeIMO.OfficeLength.Pixels(8)
                });
                cancellation.Cancel();
            }, 0, new ExcelMutationPlanOptions(), cancellation.Token));

            Assert.Equal(baselineParts, sheet.WorksheetPart.Parts.Count());
            Assert.Empty(sheet.Images);
        }

        [Fact]
        public void Test_AutoFilter_RejectsColumnOffsetOutsideRange() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");

            ExcelAutoFilter filter = sheet.Range("A1:B2").AutoFilter();
            Assert.Throws<System.ArgumentOutOfRangeException>(() => filter.Column(2));
            Assert.Empty(Assert.Single(sheet.GetAutoFilters()).Columns);
        }
    }
}