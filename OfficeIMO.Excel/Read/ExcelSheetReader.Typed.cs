namespace OfficeIMO.Excel {
    /// <summary>
    /// Strongly-typed convenience readers for <see cref="ExcelSheetReader"/>.
    /// </summary>
    internal sealed partial class ExcelSheetReader {
        /// <summary>
        /// Reads a single-column A1 range as a typed sequence.
        /// </summary>
        /// <typeparam name="T">Target element type.</typeparam>
        /// <param name="a1Range">Single-column A1 range (e.g., "B2:B100").</param>
        /// <param name="convert">Optional custom converter. If null, uses culture-aware conversion.</param>
        /// <param name="ct">Cancellation token.</param>
        public IEnumerable<T> ReadColumnAs<T>(string a1Range, Func<object, T>? convert = null, CancellationToken ct = default) {
            bool canCancel = ct.CanBeCanceled;
            using IEnumerator<object?[]>? serials = convert == null ? OpenSerialRows<T>(a1Range, ct) : null;
            foreach (var obj0 in ReadColumn(a1Range, ct)) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                object? obj = obj0;
                double? serial = NextSerial(obj, serials, 0);

                if (obj is null) {
                    yield return default(T)!;
                    continue;
                }

                T result;
                if (convert != null) {
                    result = convert(obj);
                } else {
                    var val = TryChangeType(obj, typeof(T), _opt.Culture, serial);
                    result = val is null ? default(T)! : (T)val;
                }

                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                yield return result;
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
            var (r1, _, _, _) = ExcelA1.ParseRange(a1Range);
            bool canCancel = ct.CanBeCanceled;
            int offset = 0;
            using IEnumerator<object?[]>? serials = convert == null ? OpenSerialRows<T>(a1Range, ct) : null;
            foreach (var row in ReadRows(a1Range, ct)) {
                object?[]? serialRow = serials != null && serials.MoveNext() ? serials.Current : null;
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                int rowIndex = r1 + offset;
                offset++;
                if (row is null) {
                    throw new InvalidOperationException($"Row {rowIndex} in range '{a1Range}' on sheet '{Name}' contains no cells.");
                }
                var result = new T[row.Length];
                for (int i = 0; i < row.Length; i++) {
                    if (canCancel && (i & 1023) == 0) {
                        ct.ThrowIfCancellationRequested();
                    }

                    var obj = row[i];
                    if (obj is null) {
                        result[i] = default(T)!;
                        continue;
                    }
                    if (convert != null) {
                        result[i] = convert(obj);
                    } else {
                        var val = TryChangeType(obj, typeof(T), _opt.Culture, SerialAt(obj, serialRow, i));
                        result[i] = val is null ? default(T)! : (T)val;
                    }
                }

                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
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
        public T[,] ReadRangeAs<T>(string a1Range, OfficeIMO.Excel.ExcelExecutionMode? mode = null, CancellationToken ct = default) {
            var values = ReadRange(a1Range, mode, ct);
            double?[,]? serials = ReadSerialMatrix<T>(a1Range, values.GetLength(0), values.GetLength(1), ct);
            int rows = values.GetLength(0);
            int cols = values.GetLength(1);
            var result = new T[rows, cols];
            bool canCancel = ct.CanBeCanceled;
            int convertedCells = 0;

            if (canCancel) {
                ct.ThrowIfCancellationRequested();
            }

            for (int r = 0; r < rows; r++) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                for (int c = 0; c < cols; c++) {
                    if (canCancel && (++convertedCells & 1023) == 0) {
                        ct.ThrowIfCancellationRequested();
                    }

                    var obj = values[r, c];
                    if (obj is null) {
                        result[r, c] = default(T)!;
                        continue;
                    }

                    var val = TryChangeType(obj, typeof(T), _opt.Culture, serials?[r, c]);
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    result[r, c] = val is null ? default(T)! : (T)val;
                }
            }
            return result;
        }

        /// <summary>
        /// For duration and time-of-day targets, reads each date cell's stored serial alongside the range, as typed
        /// object mapping does. The untyped value of a calendar date is already shifted to the workbook's date system,
        /// so it cannot be read back as a day count. Returns null for other targets.
        /// </summary>
        private IEnumerator<object?[]>? OpenSerialRows<T>(string a1Range, CancellationToken ct) {
            Type target = Nullable.GetUnderlyingType(typeof(T)) ?? typeof(T);
            return target == typeof(TimeSpan) || IsTimeOnly(target) ? ReadSerialRows(a1Range, ct).GetEnumerator() : null;
        }

        private IEnumerable<object?[]> ReadSerialRows(string a1Range, CancellationToken ct) {
            using System.Data.IDataReader reader = ReadRangeAsDataReader(a1Range, headersInFirstRow: false,
                chunkRows: Math.Min(DataReaderChunkRows, _opt.MaxDataReaderChunkRows), schemaSampleRows: 0, ct: ct);
            while (reader.Read()) {
                var serials = new object?[reader.FieldCount];
                for (int ordinal = 0; ordinal < serials.Length; ordinal++) {
                    if (reader.IsDBNull(ordinal) || reader.GetValue(ordinal) is not DateTime) continue;
                    try {
                        serials[ordinal] = reader.GetDouble(ordinal);
                    } catch (InvalidCastException) {
                    } catch (FormatException) {
                    }
                }
                yield return serials;
            }
        }

        private double?[,]? ReadSerialMatrix<T>(string a1Range, int rows, int columns, CancellationToken ct) {
            using IEnumerator<object?[]>? serials = OpenSerialRows<T>(a1Range, ct);
            if (serials == null) return null;
            var matrix = new double?[rows, columns];
            for (int r = 0; r < rows && serials.MoveNext(); r++) {
                object?[] serialRow = serials.Current;
                for (int c = 0; c < columns && c < serialRow.Length; c++) {
                    matrix[r, c] = serialRow[c] as double?;
                }
            }
            return matrix;
        }

        private static double? NextSerial(object? value, IEnumerator<object?[]>? serials, int ordinal) =>
            serials != null && serials.MoveNext() ? SerialAt(value, serials.Current, ordinal) : null;

        private static double? SerialAt(object? value, object?[]? row, int ordinal) =>
            value is DateTime && row != null && ordinal < row.Length ? row[ordinal] as double? : null;
    }
}

