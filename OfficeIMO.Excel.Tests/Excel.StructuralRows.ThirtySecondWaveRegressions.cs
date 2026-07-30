using System.IO;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using Xm = DocumentFormat.OpenXml.Office.Excel;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_StructuralRows_MutationPlanCountsFormulaRemovedWithDeletedRow() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellFormula(5, 1, "1+1");
            CellFormula formula = Assert.Single(
                sheet.WorksheetPart.Worksheet.Descendants<CellFormula>());
            formula.CalculateCell = true;

            ExcelRowMutationPlan plan = sheet.PlanDeleteRows(5);

            ExcelMutationImpact formulas = Assert.Single(
                plan.Impacts,
                impact => impact.Category == "formula-references");
            Assert.Equal(1, formulas.ItemCount);
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanClassifiesSparklineGroupDateAxisFormula() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet data = document.AddWorksheet("Data");
            ExcelSheet summary = document.AddWorksheet("Summary");
            data.CellValue(5, 1, 1);
            data.CellValue(6, 1, 2);
            summary.CellValue(1, 1, 1);
            summary.CellValue(2, 1, 2);
            summary.AddSparklines("'Data'!A5:A6", "B1");
            X14.SparklineGroup group = Assert.Single(
                summary.WorksheetPart.Worksheet.Descendants<X14.SparklineGroup>());
            X14.Sparkline sparkline = Assert.Single(group.Descendants<X14.Sparkline>());
            sparkline.Formula!.Text = "'Summary'!A1:A2";
            group.Formula = new Xm.Formula("'Data'!A5:A6");

            ExcelRowMutationPlan plan = data.PlanInsertRows(5);

            ExcelMutationImpact sparklines = Assert.Single(
                plan.Impacts,
                impact => impact.Category == "sparklines");
            Assert.Equal(1, sparklines.ItemCount);
        }
    }
}
