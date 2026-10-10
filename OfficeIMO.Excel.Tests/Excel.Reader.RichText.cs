using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void Reader_SharedStrings_KeepMixedRunValuesIndependent(bool editable) {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderMixedSharedStringRuns.xlsx");
            try {
                CreateSharedStringWorkbook(filePath,
                    "<si><r><t>Alpha</t></r><r><t>Beta</t></r><r><t>Gamma</t></r></si>" +
                    "<si><t>Plain</t></si><si/>" +
                    "<si><r><t>東京</t></r><rPh sb=\"0\" eb=\"2\"><t>Ignored</t></rPh><r><t>😀</t></r></si>" +
                    "<si><r><t>A</t></r><r><t>B</t></r></si>", "5", "5");
                using var spreadsheet = SpreadsheetDocument.Open(filePath, editable);
                using SharedStringCache cache = SharedStringCache.Build(spreadsheet);

                Assert.Equal(new[] { "AlphaBetaGamma", "Plain", "", "東京😀", "AB" },
                    Enumerable.Range(0, cache.Count).Select(cache.Get));
                Assert.Equal("AlphaBetaGamma", cache.Get(0));
            } finally {
                if (File.Exists(filePath)) File.Delete(filePath);
            }
        }
    }
}
