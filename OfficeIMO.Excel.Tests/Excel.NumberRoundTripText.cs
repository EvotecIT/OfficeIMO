using System.Data;
using System.Globalization;
using Xunit;

namespace OfficeIMO.Excel.Tests {
    public partial class Excel {
        // .NET Framework's "R" format writes these with 15 digits, text that parses to a neighbouring double.
        private static readonly double[] RoundTripSensitiveNumbers = {
            double.Parse("0.84551240822557006", CultureInfo.InvariantCulture),
            double.Parse("9.6311049627285008E-06", CultureInfo.InvariantCulture)
        };

        [Fact]
        public void CellValuesKeepEveryBitOfDoublesThroughTheDocumentModel() {
            string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.RoundTripSensitiveNumbers.{Guid.NewGuid():N}.xlsx");
            try {
                WriteAndReadRoundTripSensitiveNumbers(path);
            } finally {
                File.Delete(path);
            }
        }

        private static void WriteAndReadRoundTripSensitiveNumbers(string path) {
            using (var document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddSheet("Data");
                for (int index = 0; index < RoundTripSensitiveNumbers.Length; index++) {
                    sheet.Cell(index + 1, 1).SetValue(RoundTripSensitiveNumbers[index]);
                }
                document.Save();
            }

            using (var document = ExcelDocument.Load(path)) {
                ExcelSheet sheet = document.Sheets[0];
                for (int index = 0; index < RoundTripSensitiveNumbers.Length; index++) {
                    AssertSameDouble(RoundTripSensitiveNumbers[index], sheet.Cell(index + 1, 1).GetValue<double>());
                }
            }
        }

        [Theory]
        [InlineData(ExcelFileFormat.Xlsx)]
        [InlineData(ExcelFileFormat.Xls)]
        [InlineData(ExcelFileFormat.Xlsb)]
        public void WrittenDoublesReadBackWithEveryBitAsValuesAndText(ExcelFileFormat format) {
            using var output = new MemoryStream();
            using (var writer = ExcelDocument.CreateWriter(format, output, leaveOpen: true)) {
                writer.WriteTable("Data", RoundTripSensitiveNumbers, new[] { "Value" }, (row, value) => row.Write(value));
            }

            using (var reader = ExcelDocument.OpenReader(new MemoryStream(output.ToArray()))) {
                foreach (double expected in RoundTripSensitiveNumbers) {
                    Assert.True(reader.Read());
                    AssertSameDouble(expected, reader.GetDouble(0));
                    AssertSameDouble(expected, double.Parse(reader.GetString(0), CultureInfo.InvariantCulture));
                }
            }

            using (var document = ExcelDocument.Load(new MemoryStream(output.ToArray()))) {
                ExcelSheet sheet = document.Sheets[0];
                for (int index = 0; index < RoundTripSensitiveNumbers.Length; index++) {
                    AssertSameDouble(RoundTripSensitiveNumbers[index], sheet.Cell(index + 2, 1).GetValue<double>());
                }
            }
        }

#if NETCOREAPP3_0_OR_GREATER
        // .NET Core formats and parses negative zero with its sign; .NET Framework does neither. Imported values keep
        // the text a round-trip format gives them.
        [Fact]
        public void ImportedNegativeZeroKeepsItsSign() {
            byte[] input = System.Text.Encoding.UTF8.GetBytes(
                "<Workbook xmlns='urn:schemas-microsoft-com:office:spreadsheet' xmlns:ss='urn:schemas-microsoft-com:office:spreadsheet'>"
                + "<Worksheet ss:Name='Data'><Table><Row><Cell><Data ss:Type='Number'>-0</Data></Cell></Row></Table></Worksheet></Workbook>");

            using ExcelDocument document = ExcelDocument.ImportSpreadsheetXml(input).Value;
            AssertSameDouble(-0.0D, document.Sheets[0].Cell(1, 1).GetValue<double>());
        }
#endif

        private static void AssertSameDouble(double expected, double actual) =>
            Assert.Equal(BitConverter.DoubleToInt64Bits(expected), BitConverter.DoubleToInt64Bits(actual));
    }
}
