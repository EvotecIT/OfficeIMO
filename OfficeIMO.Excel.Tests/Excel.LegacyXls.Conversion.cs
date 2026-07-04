using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_Convert_XlsxToXlsAndBack_RoundTripsSupportedContent() {
            string xlsxPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xls");
            string roundTripPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsx");

            using (ExcelDocument document = ExcelDocument.Create(xlsxPath, autoSave: false)) {
                ExcelSheet sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(2, 1, "Alice");
                sheet.CellValue(2, 2, 42);
                sheet.CellValue(3, 1, true);
                document.Save();
            }

            ExcelDocument.Convert(xlsxPath, xlsPath);

            AssertNativeXlsRoundTrip(xlsPath, expectedRow2Name: "Alice");

            ExcelDocument.Convert(xlsPath, roundTripPath);

            using ExcelDocument roundTrip = ExcelDocument.Load(roundTripPath);
            Assert.False(roundTrip.WasLoadedFromLegacyXls);
            Assert.True(roundTrip.Sheets[0].TryGetCellText(1, 1, out string? header));
            Assert.Equal("Name", header);
            Assert.True(roundTrip.Sheets[0].TryGetCellText(2, 1, out string? name));
            Assert.Equal("Alice", name);
            Assert.True(roundTrip.Sheets[0].TryGetCellText(2, 2, out string? amount));
            Assert.Equal("42", amount);
        }

        [Fact]
        public void LegacyXls_Convert_BlocksUnsupportedLegacyContentUnlessLossIsAllowed() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateUnsupportedFeatureWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);
            string xlsPath = WriteTempWorkbook(compound, ".xls");
            string blockedPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsx");
            string allowedPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsx");

            try {
                NotSupportedException exception = Assert.Throws<NotSupportedException>(() => ExcelDocument.Convert(xlsPath, blockedPath));

                Assert.Contains("unsupported or preserve-only", exception.Message, StringComparison.OrdinalIgnoreCase);
                Assert.False(File.Exists(blockedPath));

                ExcelDocument.Convert(xlsPath, allowedPath, new ExcelDocumentConversionOptions {
                    AllowLossyLegacyConversion = true
                });

                using ExcelDocument converted = ExcelDocument.Load(allowedPath);
                Assert.False(converted.WasLoadedFromLegacyXls);
                Assert.True(converted.Sheets[0].TryGetCellText(1, 1, out string? header));
                Assert.Equal("Feature", header);
            } finally {
                TryDelete(xlsPath);
            }
        }
    }
}
