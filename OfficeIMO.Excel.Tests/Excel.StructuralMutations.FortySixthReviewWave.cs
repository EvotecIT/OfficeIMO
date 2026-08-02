using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_PartialFormulaRecalculation_KeepsFailedCachedFormulasDirty() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Calc");
            sheet.CellValue(1, 1, 1d);
            sheet.CellFormula(1, 2, "A1+1");
            sheet.CellFormula(1, 3, "UNSUPPORTED(A1)");
            Cell supported = sheet.WorksheetPart.Worksheet.Descendants<Cell>()
                .Single(cell => cell.CellReference?.Value == "B1");
            Cell unsupported = sheet.WorksheetPart.Worksheet.Descendants<Cell>()
                .Single(cell => cell.CellReference?.Value == "C1");
            supported.CellValue = new CellValue("2");
            supported.DataType = CellValues.Number;
            supported.CellFormula!.CalculateCell = false;
            unsupported.CellValue = new CellValue("99");
            unsupported.DataType = CellValues.Number;
            unsupported.CellFormula!.CalculateCell = false;
            document.MarkFormulaSheetRecalculated(
                sheet.WorksheetPart,
                document.CaptureFormulaInputMutationVersion());
            sheet.CellValue(1, 1, 2d);
            Assert.All(sheet.GetFormulaCells(), formula =>
                Assert.True(formula.State.HasFlag(ExcelFormulaState.Dirty)));

            Assert.Equal(1, sheet.RecalculateSupportedFormulas());

            ExcelFormulaCellInfo recalculated = sheet.GetFormulaCells().Single(item => item.CellReference == "B1");
            ExcelFormulaCellInfo failed = sheet.GetFormulaCells().Single(item => item.CellReference == "C1");
            Assert.Equal("3", recalculated.CachedValue);
            Assert.False(recalculated.State.HasFlag(ExcelFormulaState.Dirty));
            Assert.True(recalculated.State.HasFlag(ExcelFormulaState.Evaluated));
            Assert.Equal("99", failed.CachedValue);
            Assert.True(failed.State.HasFlag(ExcelFormulaState.Dirty));
            Assert.True(failed.State.HasFlag(ExcelFormulaState.Deferred));
            Assert.False(failed.State.HasFlag(ExcelFormulaState.Evaluated));
        }

        [Fact]
        public async Task Test_ModernChartSnapshots_WaitForWorkbookWriter() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Dashboard");
            ExcelModernChart authored = sheet.AddModernChart(
                new ExcelChartData(
                    new[] { "A" },
                    new[] { new ExcelChartSeries("Value", new[] { 1d }) }),
                1,
                1,
                ExcelModernChartType.Funnel);

            IReadOnlyList<ExcelModernChart> charts = await AssertWorkbookReadWaitsForWriter(
                document,
                () => sheet.ModernCharts.ToArray(),
                "modern chart snapshot");
            IReadOnlyList<ExcelModernChart> lookup = await AssertWorkbookReadWaitsForWriter(
                document,
                () => {
                    ExcelModernChart? chart = sheet.GetModernChart(authored.Name);
                    return chart == null ? Array.Empty<ExcelModernChart>() : new[] { chart };
                },
                "modern chart lookup");

            Assert.Single(charts);
            Assert.Single(lookup);
            Assert.Equal(authored.Name, lookup[0].Name);
        }
    }
}
