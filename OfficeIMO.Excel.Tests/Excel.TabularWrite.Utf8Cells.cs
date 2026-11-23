#if NET8_0_OR_GREATER
using System.Text;
using OfficeIMO.Excel;
using OfficeIMO.SharedSource.IO;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        // The row writer formats cells straight into a pooled UTF-8 buffer; any other TextWriter takes the character path.
        [Theory]
        [InlineData(true, ExcelDateSystem.NineteenHundred)]
        [InlineData(false, ExcelDateSystem.NineteenHundred)]
        [InlineData(true, ExcelDateSystem.NineteenFour)]
        public void TabularRowWriter_Utf8CellsMatchTheCharacterPath(bool includeCellReferences, ExcelDateSystem dateSystem) {
            string characterPath;
            using (var text = new StringWriter()) {
                WriteUtf8CellMatrix(text, includeCellReferences, dateSystem);
                characterPath = text.ToString();
            }
            using var bytes = new MemoryStream();
            using (var writer = new PooledUtf8TextWriter(bytes, new UTF8Encoding(false), 65536, leaveOpen: true)) {
                WriteUtf8CellMatrix(writer, includeCellReferences, dateSystem);
            }

            Assert.Equal(characterPath, Encoding.UTF8.GetString(bytes.ToArray()));
        }

        private static void WriteUtf8CellMatrix(TextWriter output, bool includeCellReferences, ExcelDateSystem dateSystem) {
            string[] texts = { "", "plain", " lead", "trail ", "a&b<c>d", "x_x", "_x0041_", "żółć", "tab\tx", "cr\rlf\n",
                "ctl\u0001x", "\"q'", "emoji \U0001F600", new string('a', 5000), new string('&', 3000) };
            double[] doubles = { 0, -0.0, 1.5, -2.25, 0.1, 1.005, 1e-7, 123456789.123, double.NaN, double.NegativeInfinity,
                double.MaxValue, -1e300, 0.30000000000000004, -0.01 };
            long[] longs = { 0, -1, 999_999_999_999_999, 1_000_000_000_000_000, long.MinValue };
            DateTime[] dates = { new(1900, 1, 1), new(1900, 2, 28, 6, 0, 0), new(1900, 3, 1), new(1904, 1, 1), new(1904, 1, 2, 12, 0, 0),
                new(2024, 5, 6, 7, 8, 9, 123), new(2024, 5, 6, 23, 59, 59, 999), new(9999, 12, 31, 23, 59, 59), new(2020, 1, 1, 0, 1, 0) };
            const int columns = 8;
            string[] prefixes = Enumerable.Range(0, columns).Select(column => "<c r=\"" + (char)('A' + column)).ToArray();
            // Column B carries a fixed style; column E applies runtime date styles.
            var styles = new string?[columns];
            styles[1] = " s=\"7\"";
            var valueStyles = new bool[columns];
            valueStyles[4] = true;
            var writer = ExcelDocument.ExcelTabularRowWriter.Create(output, 2, includeCellReferences, prefixes, styles, valueStyles,
                useCellValueNumberFormats: false, value => value.UtcDateTime, dateSystem, sharedStrings: null);
            for (int row = 0; row < 40; row++) {
                writer.BeginRow();
                writer.Write(row % 7 == 6 ? null : texts[row % texts.Length])
                    .Write(doubles[row % doubles.Length])
                    .Write(longs[row % longs.Length])
                    .Write(row * 7919 - 100)
                    .Write(dates[row % dates.Length])
                    .Write(row % 2 == 0)
                    .Write(row * 0.5m);
                writer.WriteBlank();
                writer.EndRow();
                writer.ReleaseRow();
            }
        }
    }
}
#endif
