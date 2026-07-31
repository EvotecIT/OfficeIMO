using System.Linq;
using System.Threading;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_StructuralColumns_PlanCommitRemapsCellsFormulasTablesAndSecurityRegions() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Region");
            sheet.CellValue(1, 2, "Amount");
            sheet.CellValue(1, 3, "Tax");
            sheet.CellValue(2, 1, "EU");
            sheet.CellValue(2, 2, 10);
            sheet.CellValue(2, 3, 2);
            sheet.AddTable("A1:C2", true, "Sales", TableStyle.TableStyleMedium2);
            sheet.CellFormula(4, 1, "SUM(Sales[Amount])+C2");
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
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "A");
            sheet.CellValue(1, 2, "B");
            sheet.CellValue(2, 1, 1);
            sheet.CellValue(2, 2, 2);
            sheet.AddTable("A1:B2", true, "DataTable", TableStyle.TableStyleMedium2);

            Assert.Throws<InvalidOperationException>(() => sheet.PlanDeleteColumns(1, 2));
            Assert.Equal("A1:B2", Assert.Single(document.GetTables()).Range);
        }

        [Fact]
        public void Test_StructuralColumns_PlanRevalidatesNewTableBeforeApply() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "A");
            sheet.CellValue(1, 2, "B");
            sheet.CellValue(2, 1, 1);
            sheet.CellValue(2, 2, 2);
            ExcelStructuralMutationPlan plan = sheet.PlanDeleteColumns(1, 2);
            sheet.AddTable("A1:B2", true, "DataTable", TableStyle.TableStyleMedium2);

            Assert.Throws<System.InvalidOperationException>(() => plan.Apply());
            Assert.Equal("A1:B2", Assert.Single(document.GetTables()).Range);
        }

        [Fact]
        public void Test_StructuralMutation_RollbackRestoresDeletedCalculationChain() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
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
    }
}
