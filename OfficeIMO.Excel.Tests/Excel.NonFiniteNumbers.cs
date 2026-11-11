using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls.Biff;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        /// <summary>Excel has no NaN or infinite cell value, so every format stores the #NUM! error Excel shows for one.</summary>
        [Theory]
        [InlineData(ExcelFileFormat.Xlsx)]
        [InlineData(ExcelFileFormat.Xlsb)]
        [InlineData(ExcelFileFormat.Xls)]
        public void NonFiniteNumbers_SaveAsNumErrorInEveryFormat(ExcelFileFormat format) {
            double[] values = { double.NaN, double.PositiveInfinity, double.NegativeInfinity, 1.5 };
            byte[] workbook;
            using (ExcelDocument document = ExcelDocument.Create()) {
                document.AddSheet("Data").WriteData(values, ("Value", static value => value));
                workbook = document.ToBytes(new ExcelSaveOptions { Format = format });
            }

            using ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(workbook, writable: false));
            ExcelSheet sheet = loaded.Sheets[0];
            for (int row = 2; row <= 4; row++) {
                ExcelCellData cell = sheet.Cell(row, 1).GetValue();
                Assert.Equal(ExcelCellDataKind.Error, cell.Kind);
                Assert.Equal("#NUM!", cell.Value?.ToString());
            }
            Assert.Equal(1.5, sheet.Cell(5, 1).GetValue<double>());
        }

        /// <summary>
        /// Excel refuses an XLSB RK cell holding negative zero, so negative zero and non-finite values have no RK form
        /// and are written as eight-byte numbers by every XLS and XLSB writer.
        /// </summary>
        [Theory]
        [InlineData(-0.0, false)]
        [InlineData(double.NaN, false)]
        [InlineData(double.PositiveInfinity, false)]
        [InlineData(double.NegativeInfinity, false)]
        [InlineData(0.0, true)]
        [InlineData(-2.5, true)]
        public void RkEncoding_RejectsNegativeZeroAndNonFiniteValues(double value, bool encodable) {
            Assert.Equal(encodable, BiffRkNumberReader.TryEncodeRkNumber(value, out uint rk));
            if (encodable) {
                Assert.Equal(BitConverter.DoubleToInt64Bits(value), BitConverter.DoubleToInt64Bits(BiffRkNumberReader.ReadRkNumber(rk)));
            }
        }
    }
}
