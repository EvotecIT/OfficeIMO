using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_StructuralRows_ChartCachePlanningRemainsLinearForSharedFormulaParent() {
            const int formulaCount = 32_000;
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet data = document.AddWorksheet("Data");
            ExcelSheet summary = CreateChartOwner(document, "Summary");
            ChartPart chartPart = Assert.Single(summary.WorksheetPart.DrawingsPart!.ChartParts);
            C.Formula template = chartPart.ChartSpace.Descendants<C.Formula>().First();
            OpenXmlCompositeElement reference = Assert.IsAssignableFrom<OpenXmlCompositeElement>(template.Parent);
            reference.RemoveAllChildren();
            for (int index = 0; index < formulaCount; index++) {
                reference.Append(new C.Formula("Summary!A1:A2"));
            }
            Assert.Equal(formulaCount, reference.Elements<C.Formula>().Count());

            var stopwatch = Stopwatch.StartNew();
            ExcelRowMutationPlan plan = data.PlanInsertRows(
                5,
                options: new ExcelMutationPlanOptions {
                    MaximumScannedElements = formulaCount + 1_000
                });
            stopwatch.Stop();

            Assert.True(
                stopwatch.Elapsed < TimeSpan.FromSeconds(5),
                $"Chart mutation planning exceeded the linear-time budget: {stopwatch.Elapsed}.");
            Assert.True(plan.ScannedElements <= formulaCount + 1_000);
        }
    }
}
