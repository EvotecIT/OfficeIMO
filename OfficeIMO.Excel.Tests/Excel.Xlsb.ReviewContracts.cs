using OfficeIMO.Excel;
using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Write;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Xlsb_FormulaReader_PreservesComparisonAndConcatenationPrecedence() {
            byte[] comparisonThenConcatenation = {
                0x44, 0, 0, 0, 0, 0, 0, // $A$1
                0x44, 0, 0, 0, 0, 1, 0, // $B$1
                0x0B,                    // =
                0x44, 0, 0, 0, 0, 2, 0, // $C$1
                0x08                     // &
            };
            byte[] concatenationThenComparison = {
                0x44, 0, 0, 0, 0, 0, 0, // $A$1
                0x44, 0, 0, 0, 0, 1, 0, // $B$1
                0x44, 0, 0, 0, 0, 2, 0, // $C$1
                0x08,                    // &
                0x0B                     // =
            };

            Assert.True(XlsbFormulaTextReader.TryRead(comparisonThenConcatenation, out string? first));
            Assert.True(XlsbFormulaTextReader.TryRead(concatenationThenComparison, out string? second));
            Assert.Equal("($A$1=$B$1)&$C$1", first);
            Assert.Equal("$A$1=$B$1&$C$1", second);
        }

        [Fact]
        public void Xlsb_FormulaReader_GroupsUnionUsedAsSingleFunctionArgument() {
            byte[] unionAreas = {
                0x44, 0, 0, 0, 0, 0, 0, // $A$1
                0x44, 0, 0, 0, 0, 1, 0, // $B$1
                0x10,                    // union
                0x21, 0x4B, 0x00         // AREAS, one fixed argument
            };

            Assert.True(XlsbFormulaTextReader.TryRead(unionAreas, out string? formula));
            Assert.Equal("AREAS(($A$1,$B$1))", formula);
        }

        [Fact]
        public void Xlsb_NewWorkbook_ProjectsComparisonBeforeConcatenationWithoutSemanticLoss() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Formula");
            sheet.CellValue(1, 1, "A");
            sheet.CellValue(1, 2, "A");
            sheet.CellValue(1, 3, " suffix");
            sheet.CellFormula(2, 1, "(A1=B1)&C1");

            byte[] package = document.ToBytes(ExcelFileFormat.Xlsb);

            using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(package, writable: false));
            Assert.Equal("(A1=B1)&C1", reloaded.Sheets[0].CellAt(2, 1).GetValue().Formula);
        }

        [Fact]
        public void Xlsb_NewWorkbook_PreservesRectangularMergedRangeCoordinates() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Merge");
            sheet.CellValue(2, 2, "Merged");
            sheet.MergeRange("B2:D4");

            byte[] package = document.ToBytes(ExcelFileFormat.Xlsb);

            using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(package, writable: false));
            ExcelSheet result = Assert.Single(reloaded.Sheets);
            Assert.Equal("B2:D4", Assert.Single(result.GetMergedRanges()).A1Range);
            Assert.Equal("B2:D4", result.WorksheetPart.Worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetDimension>()?.Reference?.Value);
        }

        [Fact]
        public void Xlsb_NativeRewrite_PreservesUnknownRecordsBetweenTheirSourceCells() {
            using var source = new MemoryStream();
            XlsbRecordWriter.Write(source, 145); // BrtBeginSheetData
            XlsbRecordWriter.Write(source, 0, CreateReviewRowHeaderPayload());
            XlsbRecordWriter.Write(source, 5, CreateReviewNumberCellPayload(zeroBasedColumn: 0, value: 1D));
            XlsbRecordWriter.Write(source, 700, new byte[] { 0xA1 });
            XlsbRecordWriter.Write(source, 5, CreateReviewNumberCellPayload(zeroBasedColumn: 1, value: 2D));
            XlsbRecordWriter.Write(source, 701, new byte[] { 0xB2 });
            XlsbRecordWriter.Write(source, 146); // BrtEndSheetData

            var cells = new[] {
                new XlsbWriteCell(1, 1, 0, XlsbWriteCellKind.Number, 10D),
                new XlsbWriteCell(1, 2, 0, XlsbWriteCellKind.Number, 20D)
            };

            byte[] rewritten = XlsbWorksheetPartWriter.Rewrite(source.ToArray(), cells);
            using var rewrittenStream = new MemoryStream(rewritten, writable: false);
            int[] recordTypes = XlsbRecordReader.ReadAll(rewrittenStream).Select(record => record.Type).ToArray();

            Assert.Equal(new[] { 145, 0, 5, 700, 5, 701, 146 }, recordTypes);
        }

        private static byte[] CreateReviewRowHeaderPayload() {
            var payload = new byte[17];
            WriteXlsbReviewUInt32(payload, 13, 0);
            return payload;
        }

        private static byte[] CreateReviewNumberCellPayload(int zeroBasedColumn, double value) {
            var payload = new byte[16];
            WriteXlsbReviewUInt32(payload, 0, checked((uint)zeroBasedColumn));
            byte[] number = BitConverter.GetBytes(value);
            Buffer.BlockCopy(number, 0, payload, 8, number.Length);
            return payload;
        }

        private static void WriteXlsbReviewUInt32(byte[] data, int offset, uint value) {
            data[offset] = (byte)value;
            data[offset + 1] = (byte)(value >> 8);
            data[offset + 2] = (byte)(value >> 16);
            data[offset + 3] = (byte)(value >> 24);
        }
    }
}
