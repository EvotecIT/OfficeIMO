using System.Globalization;
using System.Text;
using OfficeIMO.CSV;

internal static class CsvPackedTextSmoke {
    internal static void Run() {
        foreach (bool quoted in new[] { false, true }) {
            // Exceed the 64 KiB quote probe and cross the 128-row batch boundary
            // so native execution covers packed quote-free reads and quoted fallback.
            string value = new string('x', 512) + "\u012c\u0122\u010a\u800a\u800d\ufffeŻółć😀"
                + (quoted ? " \"quoted\"\r\nnext" : "");
            string encoded = quoted ? "\"" + value.Replace("\"", "\"\"") + "\"" : value;
            var text = new StringBuilder("Id,Value\n");
            const int count = 257;
            for (int row = 0; row < count; row++)
                text.Append(row.ToString(CultureInfo.InvariantCulture)).Append(',').Append(encoded).Append('\n');
            using var reader = CsvDocument.OpenReader(new StringReader(text.ToString()),
                options: new CsvReaderOptions { Culture = CultureInfo.InvariantCulture });
            for (int row = 0; row < count; row++) {
                if (!reader.Read() || reader.GetInt32(0) != row || reader.GetString(1) != value)
                    throw new InvalidOperationException("Native CSV packed text parsing lost a value.");
            }
            if (reader.Read()) throw new InvalidOperationException("Native CSV packed text parsing added a row.");
        }
    }
}
