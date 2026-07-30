using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_StructuralRows_MutationPlanIncludesUnchangedFormulaRecalculation() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet data = document.AddWorksheet("Data");
            ExcelSheet summary = document.AddWorksheet("Summary");
            summary.CellFormula(1, 1, "1+1");
            CellFormula formula = Assert.Single(
                summary.WorksheetPart.Worksheet.Descendants<CellFormula>());
            formula.CalculateCell = null;

            ExcelRowMutationPlan plan = data.PlanInsertRows(100);

            ExcelMutationImpact formulas = Assert.Single(
                plan.Impacts,
                impact => impact.Category == "formula-references");
            Assert.Equal(1, formulas.ItemCount);

            plan.Apply();

            Assert.True(formula.CalculateCell?.Value);
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanIgnoresUnchangedHyperlinksAndTables() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(1, 1).SetValue("Name");
            sheet.CellAt(1, 2).SetValue("Value");
            sheet.CellAt(2, 1).SetValue("One");
            sheet.CellAt(2, 2).SetValue(1);
            sheet.SetInternalLink(3, 1, sheet, "A1", display: "Top");
            sheet.AddTable(
                "A1:B2",
                hasHeader: true,
                name: "DataTable",
                OfficeIMO.Excel.TableStyle.TableStyleMedium2);

            ExcelRowMutationPlan plan = sheet.PlanInsertRows(100);

            Assert.DoesNotContain(
                plan.Impacts,
                impact => impact.Category == "hyperlinks");
            Assert.DoesNotContain(
                plan.Impacts,
                impact => impact.Category == "tables");
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanIgnoresUnchangedPivotOutputAndSource() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = CreatePivotSheet(document);

            ExcelRowMutationPlan plan = sheet.PlanInsertRows(100);

            Assert.DoesNotContain(
                plan.Impacts,
                impact => impact.Category == "pivots");
        }
    }
}
