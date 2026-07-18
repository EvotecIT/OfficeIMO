using System.IO;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData(ExcelFileFormat.Xlsx)]
        [InlineData(ExcelFileFormat.Xls)]
        [InlineData(ExcelFileFormat.Xlsb)]
        public void Test_FormulaDepth_RoundTripsAcrossPhysicalFormats(ExcelFileFormat format) {
            byte[] bytes;
            using (ExcelDocument document = ExcelDocument.Create()) {
                ExcelSheet sheet = document.AddWorksheet("Depth");
                sheet.CellValue(4, 1, 1d);
                sheet.CellFormula(3, 1, "A4+1");
                sheet.CellFormula(2, 1, "A3+1");
                sheet.CellFormula(1, 1, "A2+1");

                Assert.Equal(3, document.Calculate());
                Assert.Equal(3, document.InspectFormulas().DependencyGraph.MaximumDependencyDepth);
                bytes = document.ToBytes(format);
            }

            using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(bytes, writable: false));
            ExcelSheet reloadedSheet = Assert.Single(reloaded.Sheets);
            Assert.Equal(format, reloaded.SourceFormat);
            Assert.Equal("A2+1", reloadedSheet.GetFormulaText(1, 1));
            Assert.True(reloadedSheet.TryGetCachedFormulaValue(1, 1, out string? cached));
            Assert.Equal("4", cached);
            Assert.Equal(3, reloaded.InspectFormulas().DependencyGraph.MaximumDependencyDepth);
        }

        [Fact]
        public void Test_FormatCapabilityReport_StatesParityAndFallbacksExplicitly() {
            ExcelFormatCapabilityReport report = ExcelFormatCapabilityReport.Current;

            ExcelFormatCapabilityEntry formulas = Assert.IsType<ExcelFormatCapabilityEntry>(report.Find("Formula authoring and cached results"));
            Assert.Equal(ExcelFormatCapabilityStatus.Native, formulas.GetStatus(ExcelFileFormat.Xlsx));
            Assert.Equal(ExcelFormatCapabilityStatus.NativeSubset, formulas.GetStatus(ExcelFileFormat.Xls));
            Assert.Equal(ExcelFormatCapabilityStatus.NativeSubset, formulas.GetStatus(ExcelFileFormat.Xlsb));

            ExcelFormatCapabilityEntry modernCharts = Assert.IsType<ExcelFormatCapabilityEntry>(report.Find("Histogram, Pareto, funnel, and waterfall recipes"));
            Assert.Equal(ExcelFormatCapabilityStatus.CompatibleRecipe, modernCharts.Xlsx);
            Assert.Equal(ExcelFormatCapabilityStatus.Unsupported, modernCharts.Xls);
            Assert.Equal(ExcelFormatCapabilityStatus.Unsupported, modernCharts.Xlsb);

            ExcelFormatCapabilityEntry slicers = Assert.IsType<ExcelFormatCapabilityEntry>(report.Find("Slicer cache metadata"));
            Assert.Equal(ExcelFormatCapabilityStatus.MetadataOnly, slicers.Xlsx);
            Assert.Equal(ExcelFormatCapabilityStatus.PreserveOnly, slicers.Xls);
            Assert.Equal(ExcelFormatCapabilityStatus.PreserveOnly, slicers.Xlsb);

            string markdown = report.ToMarkdown();
            Assert.Contains("| Feature | XLSX | XLS | XLSB |", markdown);
            Assert.Contains("Pivot table authoring", markdown);
            Assert.Contains("Slicer cache metadata", markdown);
        }
    }
}
