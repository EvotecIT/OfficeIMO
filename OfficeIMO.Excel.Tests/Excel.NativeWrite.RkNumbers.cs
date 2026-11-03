using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        private static readonly double[] RkBoundaryNumbers = {
            0d, 1d, -1d, 536_870_911d, -536_870_912d, 536_870_912d, -536_870_913d,
            0.01d, 0.5d, 1.5d, 123.45d, -123.456d, 0.1d + 0.2d, 45_292.25d, 12_345_678.91d,
            5_368_709.11d, 1E-300d, 1E300d, double.Epsilon
        };

        [Theory]
        [InlineData(ExcelFileFormat.Xlsb, false)]
        [InlineData(ExcelFileFormat.Xlsb, true)]
        [InlineData(ExcelFileFormat.Xls, false)]
        [InlineData(ExcelFileFormat.Xls, true)]
        public void NativeBinaryWriters_RoundTripNumbersAcrossCompactRkBoundaries(ExcelFileFormat format, bool resave) {
            byte[] workbook;
            using (ExcelDocument document = ExcelDocument.Create()) {
                ExcelSheet sheet = document.AddSheet("Data");
                for (int index = 0; index < RkBoundaryNumbers.Length; index++) {
                    sheet.Cell(index + 1, 1).SetValue(RkBoundaryNumbers[index]);
                }
                workbook = document.ToBytes(new ExcelSaveOptions { Format = format });
            }
            if (resave) {
                using var input = new MemoryStream(workbook);
                using ExcelDocument loaded = ExcelDocument.Load(input);
                workbook = loaded.ToBytes(new ExcelSaveOptions { Format = format });
            }

            using ExcelWorkbookDataReader reader = ExcelDocument.OpenReader(
                workbook, new ExcelReaderOptions { HasHeaderRow = false });
            foreach (double expected in RkBoundaryNumbers) {
                Assert.True(reader.Read());
                Assert.Equal(BitConverter.DoubleToInt64Bits(expected), BitConverter.DoubleToInt64Bits(reader.GetDouble(0)));
            }
            Assert.False(reader.Read());
        }
    }
}
