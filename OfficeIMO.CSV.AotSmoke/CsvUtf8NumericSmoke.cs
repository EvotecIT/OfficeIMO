using System.Globalization;
using System.Text;
using OfficeIMO.CSV;

internal static class CsvUtf8NumericSmoke {
    internal static void Run() {
        const string csv = "Byte;Int16;Int32;Int64;Decimal\n"
            + "255;-32768;2147483647;-9223372036854775808;-0.0000\n"
            + "2.55e2;-3.2768e4;2.147483647e9;-9.223372036854775808e18;1,234.5000\n";
        using var input = new MemoryStream(Encoding.UTF8.GetBytes(csv));
        using var reader = CsvDocument.OpenReader(input, new CsvReaderOptions {
            Delimiter = ';', Culture = CultureInfo.InvariantCulture
        });
        for (int row = 0; row < 2; row++) {
            if (!reader.Read()
                || reader.GetByte(0) != byte.MaxValue
                || reader.GetInt16(1) != short.MinValue
                || reader.GetInt32(2) != int.MaxValue
                || reader.GetInt64(3) != long.MinValue) {
                throw new InvalidOperationException("The UTF-8 numeric getters lost an integer value.");
            }
            decimal expected = row == 0 ? new decimal(0, 0, 0, true, 4) : 1234.5000m;
            if (!decimal.GetBits(expected).SequenceEqual(decimal.GetBits(reader.GetDecimal(4)))) {
                throw new InvalidOperationException("The UTF-8 decimal getter lost precision, sign or scale.");
            }
        }
        if (reader.Read()) throw new InvalidOperationException("The numeric smoke read an extra row.");
        Console.WriteLine("UTF8_NUMERIC_AOT_PASS widths=true decimalBits=true broadFallback=true");
    }
}
