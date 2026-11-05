using System.Globalization;
using System.Text;
using OfficeIMO.CSV;
using OfficeIMO.Data;

internal static class CsvPackedTextSmoke {
    internal static void Run() {
        VerifyInternedStream();
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
            for (int transport = 0; transport < 3; transport++) {
                using var bytes = new MemoryStream(Encoding.UTF8.GetBytes(text.ToString()));
                using TextReader input = transport == 0 ? new StringReader(text.ToString()) : new StreamReader(bytes);
                using var reader = transport == 1 ? CsvDocument.OpenReader(bytes) : CsvDocument.OpenReader(input);
                for (int row = 0; row < count; row++) {
                    if (!reader.Read() || reader.GetInt32(0) != row || reader.GetString(1) != value)
                        throw new InvalidOperationException("Native CSV packed parsing lost a value.");
                }
                if (reader.Read()) throw new InvalidOperationException("Native CSV packed parsing added a row.");
            }
            int[] ids = CsvDocument.ReadTextRowsAsParallel<int>(text.ToString(), _ => record => {
                if (record.GetString(1) != value) throw new InvalidOperationException("Native CSV partition lost Unicode text.");
                return record.GetInt32(0);
            }, parallelOptions: new ParallelRowMappingOptions { MaxDegreeOfParallelism = 2, BatchSize = 32 }).ToArray();
            if (!ids.SequenceEqual(Enumerable.Range(0, count)))
                throw new InvalidOperationException("Native CSV partition changed row order.");
        }
    }

    private static void VerifyInternedStream() {
        using var input = new MemoryStream(Encoding.UTF8.GetBytes("Value\nrepeat\n\"repeat\"\nrepeat\n"));
        using var reader = CsvDocument.OpenReader(input, new CsvReaderOptions { InternStrings = true, DetectDelimiter = false });
        if (!reader.Read()) throw new InvalidOperationException("Native interning lost the first row.");
        string first = reader.GetString(0);
        for (int row = 0; row < 2; row++) {
            if (!reader.Read() || !ReferenceEquals(first, reader.GetString(0)))
                throw new InvalidOperationException("Native interning lost string identity across quoted fallback.");
        }
        if (first != "repeat" || reader.Read()) throw new InvalidOperationException("Native interning changed the records.");
    }
}
