using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void FormulaEvaluator_StandalonePassReleasesSharedStringsForRetainedWorksheetHandles() {
            using var document = ExcelDocument.Create();
            var source = document.AddSheet("Source");
            source.Cell(1, 1).SetValue("Before");
            ExcelSheet? retained = null;
            document.Calculation.RegisterCustomFunction("ECHO", (context, arguments) => {
                retained = context.Worksheet;
                return arguments.Count == 1 ? arguments[0] : null;
            });
            source.Cell(1, 3).SetFormula("ECHO(A1)");
            var report = document.AddSheet("Report");
            report.Cell(1, 1).SetFormula("Source!C1");
            report.Cell(2, 1).SetFormula("LEN(Source!A1)");
            Assert.Equal(2, report.RecalculateSupportedFormulas());
            AssertFormulaCache(report, 1, 1, "Before");
            AssertFormulaCache(report, 2, 1, "6");
            Assert.NotNull(retained);

            // Direct Open XML edits do not advance the typed input mutation counter.
            var item = Assert.Single(document.OpenXmlDocument.WorkbookPart!.SharedStringTablePart!.SharedStringTable!.Elements<SharedStringItem>());
            item.Text!.Text = "After";
            Assert.Equal("After", retained!.Cell(1, 1).GetValue().Value);
            Assert.Equal(2, report.RecalculateSupportedFormulas());
            AssertFormulaCache(report, 1, 1, "After");
            AssertFormulaCache(report, 2, 1, "5");
            Assert.Equal(3, document.Calculate());
            AssertFormulaCache(source, 1, 3, "After");
        }
    }
}
