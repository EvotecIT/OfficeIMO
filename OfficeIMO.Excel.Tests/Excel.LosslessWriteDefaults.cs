using System.Data;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Values Excel cannot hold as numbers are written as their exact text by default, so reading the workbook
    /// returns what was written. Each route is a public write path that tabular exports use.
    /// </summary>
    public sealed class ExcelLosslessWriteDefaultsTests {
        private static readonly DateTimeOffset Timestamp = new(new DateTime(2026, 3, 14, 9, 26, 53).AddTicks(5_897_932), new TimeSpan(5, 45, 0));
        private const decimal WideDecimal = 1234567890123456.7890123456m;
        private static readonly byte[] Bytes = { 0, 1, 2, 250, 255 };

        private static readonly (string Name, Type Type, object Value)[] Columns = {
            ("Timestamp", typeof(DateTimeOffset), Timestamp),
            ("Wide", typeof(decimal), WideDecimal),
            ("DecimalMax", typeof(decimal), decimal.MaxValue),
            ("LongMax", typeof(long), long.MaxValue),
            ("ULongMax", typeof(ulong), ulong.MaxValue),
            ("Price", typeof(decimal), 12.34m),
            ("Count", typeof(int), 42),
            ("Payload", typeof(byte[]), Bytes)
        };

        [Theory]
        [InlineData("cellValue")]
        [InlineData("setValues")]
        [InlineData("addTable")]
        [InlineData("insertDataReader")]
        [InlineData("workbookWriter")]
        public void DefaultWriteReadsBackExactValues(string route) {
            object?[] values = WriteAndRead(route, serialDateTimeOffsets: false);

            Assert.Equal(Timestamp.ToString("o", System.Globalization.CultureInfo.InvariantCulture), values[0]);
            Assert.Equal("1234567890123456.7890123456", values[1]);
            Assert.Equal("79228162514264337593543950335", values[2]);
            Assert.Equal("9223372036854775807", values[3]);
            Assert.Equal("18446744073709551615", values[4]);
            Assert.IsNotType<string>(values[5]);
            Assert.Equal(12.34d, Convert.ToDouble(values[5], System.Globalization.CultureInfo.InvariantCulture));
            Assert.IsNotType<string>(values[6]);
            Assert.Equal(42d, Convert.ToDouble(values[6], System.Globalization.CultureInfo.InvariantCulture));
            Assert.Equal(Convert.ToBase64String(Bytes), values[7]);
        }

        [Theory]
        [InlineData("cellValue")]
        [InlineData("addTable")]
        [InlineData("insertDataReader")]
        [InlineData("workbookWriter")]
        public void SerialDateTimeOffsetModeWritesDateSerials(string route) {
            object?[] values = WriteAndRead(route, serialDateTimeOffsets: true);

            DateTime expected = route == "workbookWriter" ? Timestamp.LocalDateTime : Timestamp.UtcDateTime;
            Assert.IsType<DateTime>(values[0]);
            Assert.Equal(expected, (DateTime)values[0]!, TimeSpan.FromMilliseconds(1));
        }

        private static object?[] WriteAndRead(string route, bool serialDateTimeOffsets) {
            var table = new DataTable("Data");
            foreach (var column in Columns) table.Columns.Add(column.Name, column.Type);
            table.Rows.Add(Columns.Select(column => column.Value).ToArray());

            using var stream = new MemoryStream();
            if (route == "workbookWriter") {
                var options = new ExcelWriterOptions {
                    DateTimeOffsetWriteMode = serialDateTimeOffsets ? ExcelDateTimeOffsetWriteMode.Serial : ExcelDateTimeOffsetWriteMode.Text
                };
                using (var writer = ExcelDocument.CreateWriter(stream, options, leaveOpen: true))
                using (var reader = table.CreateDataReader()) {
                    writer.WriteTable("Data", reader);
                }
            } else {
                using var document = ExcelDocument.Create(stream);
                // Assigning a conversion strategy selects serial dates, as it did before text became the default.
                if (serialDateTimeOffsets) document.DateTimeOffsetWriteStrategy = value => value.UtcDateTime;
                ExcelSheet sheet = document.AddSheet("Data");
                switch (route) {
                    case "cellValue":
                        for (int column = 0; column < Columns.Length; column++) {
                            sheet.CellValue(1, column + 1, Columns[column].Name);
                            sheet.CellValue(2, column + 1, Columns[column].Value);
                        }
                        break;
                    case "setValues":
                        sheet.SetValues(Columns.SelectMany((column, index) => new[] {
                            (1, index + 1, (object)column.Name),
                            (2, index + 1, column.Value)
                        }));
                        break;
                    case "addTable":
                        sheet.AddTable(table);
                        break;
                    case "insertDataReader":
                        using (var reader = table.CreateDataReader()) sheet.InsertDataReader(reader);
                        break;
                    default:
                        throw new ArgumentOutOfRangeException(nameof(route));
                }

                document.Save(stream);
            }

            stream.Position = 0;
            using var result = ExcelDocument.OpenReader(stream);
            Assert.True(result.Read());
            var values = new object?[result.FieldCount];
            for (int i = 0; i < values.Length; i++) values[i] = result.GetValue(i);
            return values;
        }
    }
}
