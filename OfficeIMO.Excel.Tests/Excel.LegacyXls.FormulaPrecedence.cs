using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData(0x09, "<")]
        [InlineData(0x0A, "<=")]
        [InlineData(0x0B, "=")]
        [InlineData(0x0C, ">=")]
        [InlineData(0x0D, ">")]
        [InlineData(0x0E, "<>")]
        public void LegacyXls_FormulaComparisonBeforeConcatenationPreservesGroupingAndCache(int token, string comparison) {
            byte[] expression = { 0x1E, 1, 0, 0x1E, 2, 0, (byte)token, 0x17, 1, 0, (byte)'x', 0x08 };
            AssertSavedFormula(expression, "(1" + comparison + "2)&\"x\"");
        }

        [Theory]
        [InlineData(new byte[] { 0x1E, 1, 0, 0x1E, 2, 0, 0x08, 0x1E, 3, 0, 0x0B }, "1&2=3")]
        [InlineData(new byte[] { 0x1E, 1, 0, 0x1E, 2, 0, 0x1E, 3, 0, 0x04, 0x04 }, "1-(2-3)")]
        [InlineData(new byte[] { 0x1E, 1, 0, 0x1E, 2, 0, 0x03, 0x1E, 3, 0, 0x05 }, "(1+2)*3")]
        [InlineData(new byte[] { 0x1E, 1, 0, 0x1E, 2, 0, 0x1E, 3, 0, 0x06, 0x06 }, "1/(2/3)")]
        [InlineData(new byte[] { 0x1E, 1, 0, 0x1E, 2, 0, 0x1E, 3, 0, 0x07, 0x07 }, "1^(2^3)")]
        [InlineData(new byte[] { 0x1E, 1, 0, 0x1E, 2, 0, 0x07, 0x1E, 3, 0, 0x07 }, "1^2^3")]
        [InlineData(new byte[] { 0x1E, 1, 0, 0x1E, 2, 0, 0x03, 0x13 }, "-(1+2)")]
        [InlineData(new byte[] { 0x1E, 1, 0, 0x1E, 2, 0, 0x03, 0x14 }, "(1+2)%")]
        public void LegacyXls_FormulaOperatorFamiliesPreservePostfixGrouping(byte[] expression, string expected) {
            AssertSavedFormula(expression, expected);
        }

        private static void AssertSavedFormula(byte[] expression, string expected) {
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(
                LegacyXlsTestWorkbookBuilder.CreateOperatorPrecedenceWorkbookStream(expression));
            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound),
                new LegacyXlsImportOptions { ReportUnsupportedContent = false });
            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument saved = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            Cell cell = Assert.Single(saved.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<Cell>());
            Assert.Equal(expected, cell.CellFormula!.Text);
            Assert.Equal("77", cell.CellValue!.Text);
            Assert.Equal(CellValues.Number, cell.DataType!.Value);
        }

        private static partial class LegacyXlsTestWorkbookBuilder {
            internal static byte[] CreateOperatorPrecedenceWorkbookStream(byte[] expression) {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0, 6, 5, 0, 0xDB, 0x0B, 0xCC, 7 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Operators"));
                WriteRecord(stream, 0x000A, Array.Empty<byte>());
                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0, 6, 0x10, 0, 0xDB, 0x0B, 0xCC, 7 });
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(0, 0, 77d, formulaTokens: expression));
                WriteRecord(stream, 0x000A, Array.Empty<byte>());
                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }
        }
    }
}
