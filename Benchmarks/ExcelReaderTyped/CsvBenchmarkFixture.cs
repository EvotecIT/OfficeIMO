using OfficeIMO.CSV;
using OfficeIMO.Data;
using System.Data;
using System.Globalization;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks {
    internal static class CsvBenchmarkFixture {
        internal static readonly string[] Names = ["alpha", "beta", "gamma", "delta", "epsilon", "zeta", "eta", "theta"];

        // Exact upstream CsvGenerator.Build/BuildTyped field values and LF separators.
        internal static byte[] Create(int rows, bool headers) {
            StringBuilder text = new StringBuilder(rows * 40);
            if (headers) text.Append("Name,Id,Date,Value\n");
            for (int index = 1; index <= rows; index++) {
                TypedRecord record = TypedWorkbookFixture.ExpectedRecord(index);
                text.Append(record.Name).Append(',').Append(record.Id.ToString(CultureInfo.InvariantCulture)).Append(',')
                    .Append(record.Date.ToString("O", CultureInfo.InvariantCulture)).Append(',')
                    .Append(record.Value.ToString(CultureInfo.InvariantCulture)).Append('\n');
            }
            return Encoding.UTF8.GetBytes(text.ToString());
        }

        internal static byte[] CreateWide(int rows) {
            StringBuilder text = new StringBuilder(rows * 32 * 6);
            for (int row = 0; row < rows; row++) {
                for (int column = 0; column < 32; column++) {
                    if (column > 0) text.Append(',');
                    text.Append(Names[(row + column) % Names.Length]);
                }
                text.Append('\n');
            }
            return Encoding.UTF8.GetBytes(text.ToString());
        }

        internal static TypedRecord ReadRecord(IDataRecord reader) => new() {
            Name = reader.GetString(0), Id = reader.GetInt32(1), Date = reader.GetDateTime(2), Value = reader.GetDouble(3),
        };

        internal static void Validate(IDataReader reader, int rows, bool headers) {
            if (reader.FieldCount != 4) throw new InvalidDataException("CSV must have four fields.");
            if (headers) {
                for (int column = 0; column < 4; column++) {
                    if (reader.GetName(column) != TypedWorkbookFixture.Headers[column])
                        throw new InvalidDataException("CSV headers differ.");
                }
            }
            int count = 0;
            long checksum = 0;
            while (reader.Read()) {
                TypedRecord record = ReadRecord(reader);
                TypedWorkbookFixture.ValidateRecord(record, ++count);
                checksum = unchecked(checksum + TypedWorkbookFixture.Accumulate(record));
            }
            Check(checksum, count, rows);
        }

        internal static long Check(long checksum, int count, int rows) => Check(checksum, count, rows, TypedWorkbookFixture.ExpectedChecksum(rows));

        internal static long Check(long checksum, int count, int rows, long expected) => count == rows && checksum == expected
            ? checksum : throw new InvalidDataException($"CSV count/checksum mismatch: rows={count}, checksum={checksum}.");

        internal static void Describe(byte[] data, int rows, string shape) {
            DescribeEngines();
            Console.WriteLine($"CSV {shape}: rows={rows}, bytes={data.Length}, sha256={Convert.ToHexString(SHA256.HashData(data))}.");
        }

        internal static void DescribeEngines() {
            BenchmarkInput.WriteDescription();
            foreach (Assembly assembly in new[] { typeof(CsvDocument).Assembly, typeof(global::Sylvan.Data.Csv.CsvDataReader).Assembly }) {
                string version = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion
                    ?? assembly.GetName().Version?.ToString() ?? "unknown";
                Console.WriteLine($"CSV engine={assembly.GetName().Name}, version={version}, "
                    + $"sha256={Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(assembly.Location)))}.");
            }
        }
    }
}
