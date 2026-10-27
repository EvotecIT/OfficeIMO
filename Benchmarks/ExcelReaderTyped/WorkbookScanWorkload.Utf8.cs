#if OFFICEIMO_BENCHMARK_NEW_APIS
using ExcelReader.Core.Reader;
using System.Text;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks {
    internal sealed partial class WorkbookScanWorkload {
        /// <summary>Verifies public borrowed text byte-for-byte for both measured peer policies.</summary>
        internal void QualifyBorrowedUtf8() {
            if (_format == ComparisonWorkbookFormat.Xls)
                throw new InvalidOperationException("The XLS comparison has no public borrowed UTF8 lane.");
            foreach (ExcelReadOptions? options in new ExcelReadOptions?[] { null, PrefetchOptions }) {
                (long raw, long materialized, _) = QualifyValues(options, borrowedUtf8: true);
                if (raw != _spanChecksum || materialized != _stringChecksum)
                    throw new InvalidDataException("Borrowed UTF8 qualification changed the native field checksum.");
            }
            OfficeIMOBorrowedUtf8();
            ExcelReader();
            ExcelReader(prefetch: true);
            Console.WriteLine($"Qualified public borrowed UTF8 {_format}: rows={_rows}, columns={_columns}; "
                + "every text byte and native field/order verified for default and prefetched peer reads. "
                + "Every timed scan opens a fresh reader, including normalized-text encoding and reader-owned caching.");
        }

        private static void QualifyBorrowedText(Cell peer, ExcelWorkbookDataReader office,
            int column, object? value, int row) {
            if (value is not string text) return;
            byte[] expected = Encoding.UTF8.GetBytes(text);
            if (!peer.Value.SequenceEqual(expected))
                throw new InvalidDataException($"Peer UTF8 bytes differ at row {row + 1}, column {column + 1}.");
            bool available = office.TryGetUtf8Text(column, out ReadOnlySpan<byte> borrowed);
            // A declined accessor is permitted for an already-qualified empty string;
            // no bytes need conversion or copying. Nonempty unsupported text must fail.
            if (!available && expected.Length == 0) return;
            if (!available || !borrowed.SequenceEqual(expected))
                throw new InvalidDataException($"OfficeIMO public UTF8 bytes unavailable or unequal at row {row + 1}, column {column + 1}.");
        }

        /// <summary>Consumes borrowed text and native values without converting unsupported text.</summary>
        internal long OfficeIMOBorrowedUtf8() {
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenReader(_bytes,
                new ExcelReadOptions { HasHeaderRow = false, SheetIndex = 0 });
            long sum = 0;
            int rows = 0;
            while (reader.Read()) {
                for (int column = 0; column < reader.FieldCount; column++) {
                    if (reader.TryGetUtf8Text(column, out ReadOnlySpan<byte> text))
                        sum = unchecked(sum + text.Length);
                    else switch (reader.GetValue(column)) {
                        case null or DBNull: break;
                        case string { Length: 0 }: break;
                        case double number: sum = unchecked(sum + (long)number); break;
                        case DateTime date: sum = unchecked(sum + date.Ticks); break;
                        default: throw new InvalidDataException($"Unsupported public UTF8 field at row {rows + 1}, column {column + 1}.");
                    }
                }
                rows++;
            }
            return Check(sum, rows, _spanChecksum);
        }
    }
}
#endif
