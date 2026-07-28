using System;
using System.Data;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_StructuralRows_MutationPlanIsNonMutatingAndAppliesOnce() {
            using var document = ExcelDocument.Create();
            ExcelSheet data = document.AddWorksheet("Data");
            ExcelSheet summary = document.AddWorksheet("Summary");
            data.CellValue(1, 1, "Header");
            data.CellValue(2, 1, 10);
            data.CellValue(3, 1, 20);
            data.CellFormula(1, 3, "SUM(A2:A3)");
            summary.CellFormula(1, 1, "'Data'!A3");
            document.SetNamedRange("DataRows", "'Data'!A2:A3", save: false);

            ExcelRowMutationPlan plan = data.PlanInsertRows(2);

            Assert.Equal(ExcelRowMutationKind.Insert, plan.Kind);
            Assert.Equal("Data", plan.SheetName);
            Assert.Equal(2, plan.FirstRow);
            Assert.Equal(1, plan.Count);
            Assert.True(plan.RequiresFullRecalculation);
            Assert.True(plan.ScannedElements > 0);
            Assert.Contains(plan.Impacts, impact =>
                impact.Category == "worksheet-cells" && impact.ItemCount == 2);
            Assert.Contains(plan.Impacts, impact =>
                impact.Category == "formula-references" && impact.ItemCount >= 2);
            Assert.Contains(plan.Impacts, impact =>
                impact.Category == "defined-names" && impact.ItemCount == 1);
            Assert.Equal(10, data.CellAt(2, 1).GetValue<int>());

            plan.Apply();

            Assert.True(plan.IsApplied);
            Assert.True(data.CellAt(2, 1).GetValue().IsBlank);
            Assert.Equal(10, data.CellAt(3, 1).GetValue<int>());
            Assert.Equal("SUM(A3:A4)", data.GetFormulaText(1, 3));
            Assert.Equal("'Data'!A4", summary.GetFormulaText(1, 1));
            Assert.Throws<InvalidOperationException>(() => plan.Apply());
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanHonorsInspectionBudgetWithoutMutation() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            for (int row = 1; row <= 10; row++) {
                sheet.CellValue(row, 1, row);
            }

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                sheet.PlanDeleteRows(
                    2,
                    options: new ExcelMutationPlanOptions { MaximumScannedElements = 3 }));

            Assert.Contains("exceeded its limit", exception.Message, StringComparison.Ordinal);
            Assert.Equal(2, sheet.CellAt(2, 1).GetValue<int>());
            Assert.Equal(10, sheet.CellAt(10, 1).GetValue<int>());
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanRejectsDeferredWritesWithoutMaterializing() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            var table = new DataTable("Rows");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Score", typeof(int));
            table.Rows.Add("North", 10);
            table.Rows.Add("South", 20);
            sheet.InsertDataTable(table);
            Assert.True(document.HasDeferredDirectDataSetImport);

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                sheet.PlanInsertRows(2));

            Assert.Contains("pending deferred", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.True(document.HasDeferredDirectDataSetImport);
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanPreservesPendingDirectCellBuffer() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, 1);
            sheet.CellValue(2, 1, 2);
            sheet.CellValue(3, 1, 3);
            Assert.True(document.HasPendingDirectCellValues);

            ExcelRowMutationPlan plan = sheet.PlanInsertRows(2);

            Assert.True(document.HasPendingDirectCellValues);
            Assert.Contains(plan.Impacts, impact =>
                impact.Category == "worksheet-cells" && impact.ItemCount == 2);
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanBudgetsMalformedSheetEntriesBeforeRelationshipLookup() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            Sheets sheets = Assert.IsType<Sheets>(document.WorkbookRoot.Sheets);
            for (uint index = 0; index < 10; index++) {
                sheets.Append(new Sheet {
                    Name = "Malformed" + index,
                    SheetId = 100U + index
                });
            }

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                sheet.PlanDeleteRows(
                    1,
                    options: new ExcelMutationPlanOptions { MaximumScannedElements = 5 }));

            Assert.Contains("exceeded its limit", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanRevalidatesBeforeApply() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, 1);
            sheet.CellValue(2, 1, 2);
            sheet.CellValue(3, 1, 3);
            ExcelRowMutationPlan plan = sheet.PlanInsertRows(2);
            sheet.SetArrayFormula("B1:B3", "A1:A3*2");

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() => plan.Apply());

            Assert.Contains("array formula", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.False(plan.IsApplied);
            Assert.Equal(2, sheet.CellAt(2, 1).GetValue<int>());
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanRespectsDefinedNameScope() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet data = document.AddWorksheet("Data");
            document.AddWorksheet("Other");
            var globalName = new DefinedName("'Data'!$A$5:$B$7") {
                Name = "DataRows"
            };
            var otherLocalName = new DefinedName("$A$2:$B$3") {
                Name = "OtherRows",
                LocalSheetId = 1U
            };
            document.WorkbookRoot.DefinedNames = new DefinedNames(
                globalName,
                otherLocalName);

            ExcelRowMutationPlan plan = data.PlanDeleteRows(2);

            ExcelMutationImpact names = Assert.Single(
                plan.Impacts,
                impact => impact.Category == "defined-names");
            Assert.Equal(1, names.ItemCount);
            Assert.Equal("'Data'!$A$5:$B$7", globalName.Text);
            Assert.Equal("$A$2:$B$3", otherLocalName.Text);
        }
    }
}
