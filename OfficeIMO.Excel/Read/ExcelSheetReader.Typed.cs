using System.Threading;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Strongly-typed convenience readers for <see cref="ExcelSheetReader"/>.
    /// </summary>
    public sealed partial class ExcelSheetReader {
        /// <summary>
        /// Reads a single-column A1 range as a typed sequence.
        /// </summary>
        /// <typeparam name="T">Target element type.</typeparam>
        /// <param name="a1Range">Single-column A1 range (e.g., "B2:B100").</param>
        /// <param name="convert">Optional custom converter. If null, uses culture-aware conversion.</param>
        /// <param name="ct">Cancellation token.</param>
        public IEnumerable<T> ReadColumnAs<T>(string a1Range, Func<object, T>? convert = null, CancellationToken ct = default) {
            foreach (var obj in ReadColumn(a1Range, ct)) {
                if (ct.IsCancellationRequested) yield break;
                if (obj is null) {
                    yield return default(T)!;
                    continue;
                }
                if (convert != null) {
                    yield return convert(obj);
                } else {
                    var val = TryChangeType(obj, typeof(T), _opt.Culture);
                    yield return val is null ? default(T)! : (T)val;
                }
            }
        }

        /// <summary>
        /// Streams each row within the A1 range as a typed array.
        /// </summary>
        /// <typeparam name="T">Target element type for each cell.</typeparam>
        /// <param name="a1Range">Inclusive A1 range (e.g., "A1:C100").</param>
        /// <param name="convert">Optional custom converter. If null, uses culture-aware conversion.</param>
        /// <param name="ct">Cancellation token.</param>
        public IEnumerable<T[]> ReadRowsAs<T>(string a1Range, Func<object, T>? convert = null, CancellationToken ct = default) {
            var (r1, _, _, _) = A1.ParseRange(a1Range);
            int offset = 0;
            foreach (var row in ReadRows(a1Range, ct)) {
                if (ct.IsCancellationRequested) yield break;
                int rowIndex = r1 + offset;
                offset++;
                if (row is null) {
                    throw new InvalidOperationException($"Row {rowIndex} in range '{a1Range}' on sheet '{Name}' contains no cells.");
                }
                var result = new T[row.Length];
                for (int i = 0; i < row.Length; i++) {
                    var obj = row[i];
                    if (obj is null) {
                        result[i] = default(T)!;
                        continue;
                    }
                    if (convert != null) {
                        result[i] = convert(obj);
                    } else {
                        var val = TryChangeType(obj, typeof(T), _opt.Culture);
                        result[i] = val is null ? default(T)! : (T)val;
                    }
                }
                yield return result;
            }
        }

        /// <summary>
        /// Reads a rectangular range into a dense typed matrix.
        /// </summary>
        /// <typeparam name="T">Target element type for each cell.</typeparam>
        /// <param name="a1Range">Inclusive A1 range (e.g., "A1:C10").</param>
        /// <param name="mode">Execution override (affects conversion only).</param>
        /// <param name="ct">Cancellation token.</param>
        public T[,] ReadRangeAs<T>(string a1Range, OfficeIMO.Excel.ExecutionMode? mode = null, CancellationToken ct = default) {
            var values = ReadRange(a1Range, mode, ct);
            int rows = values.GetLength(0);
            int cols = values.GetLength(1);
            var result = new T[rows, cols];

            for (int r = 0; r < rows; r++) {
                if (ct.IsCancellationRequested) break;
                for (int c = 0; c < cols; c++) {
                    var obj = values[r, c];
                    if (obj is null) {
                        result[r, c] = default(T)!;
                        continue;
                    }
                    var val = TryChangeType(obj, typeof(T), _opt.Culture);
                    result[r, c] = val is null ? default(T)! : (T)val;
                }
            }
            return result;
        }
    }
}

