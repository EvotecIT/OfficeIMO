namespace OfficeIMO.Excel.LegacyXls.Read {
    internal sealed partial class LegacyXlsTabularDataReader {
        private List<LegacyXlsBufferedRow> BufferSchemaRows() {
            if (_options.MaxDataReaderSchemaSampleRows < 0) {
                throw new ArgumentOutOfRangeException(
                    nameof(_options.MaxDataReaderSchemaSampleRows),
                    "Schema sample row limit cannot be negative.");
            }
            if (_options.SchemaSampleRows > _options.MaxDataReaderSchemaSampleRows) {
                throw new InvalidOperationException(
                    $"Schema sample row count exceeds {_options.MaxDataReaderSchemaSampleRows}.");
            }
            if (_options.MaxDataReaderBufferedCells <= 0L) {
                throw new ArgumentOutOfRangeException(
                    nameof(_options.MaxDataReaderBufferedCells),
                    "Buffered cell limit must be greater than zero.");
            }

            var rows = new List<LegacyXlsBufferedRow>(Math.Min(_options.SchemaSampleRows, 256));
            var inferred = new Type?[FieldCount];
            var mixed = new bool[FieldCount];
            while (rows.Count < _options.SchemaSampleRows) {
                _cancellationToken.ThrowIfCancellationRequested();
                if (!ReadSourceRow()) {
                    break;
                }

                long bufferedCells = checked((long)(rows.Count + 1) * FieldCount);
                if (bufferedCells > _options.MaxDataReaderBufferedCells) {
                    throw new InvalidOperationException(
                        $"Schema sampling would buffer {bufferedCells} cells, exceeding the configured limit of {_options.MaxDataReaderBufferedCells}.");
                }

                rows.Add(new LegacyXlsBufferedRow(
                    (ValueKind[])_kinds.Clone(),
                    (double[])_numbers.Clone(),
                    (DateTime[])_dates.Clone(),
                    (bool[])_booleans.Clone(),
                    (string?[])_strings.Clone()));
                for (int ordinal = 0; ordinal < FieldCount; ordinal++) {
                    if (mixed[ordinal] || _kinds[ordinal] == ValueKind.Empty) continue;
                    Type next = GetValueType(ordinal);
                    if (inferred[ordinal] == null) {
                        inferred[ordinal] = next;
                    } else if (inferred[ordinal] != next) {
                        inferred[ordinal] = typeof(object);
                        mixed[ordinal] = true;
                    }
                }
            }

            for (int ordinal = 0; ordinal < FieldCount; ordinal++) {
                _columnTypes[ordinal] = inferred[ordinal] ?? typeof(object);
            }

            _hasCurrentRow = false;
            Array.Clear(_kinds, 0, _kinds.Length);
            Array.Clear(_strings, 0, _strings.Length);
            return rows;
        }

        private void LoadBufferedRow(LegacyXlsBufferedRow row) {
            Array.Copy(row.Kinds, _kinds, FieldCount);
            Array.Copy(row.Numbers, _numbers, FieldCount);
            Array.Copy(row.Dates, _dates, FieldCount);
            Array.Copy(row.Booleans, _booleans, FieldCount);
            Array.Copy(row.Strings, _strings, FieldCount);
        }

        private Type GetValueType(int ordinal) => _kinds[ordinal] switch {
            ValueKind.Text or ValueKind.Error => typeof(string),
            ValueKind.Number => GetNumericValue(_numbers[ordinal]).GetType(),
            ValueKind.Boolean => typeof(bool),
            ValueKind.Date => typeof(DateTime),
            _ => typeof(object)
        };

        private static Type[] CreateObjectColumnTypes(int fieldCount) {
            var types = new Type[fieldCount];
            for (int ordinal = 0; ordinal < types.Length; ordinal++) {
                types[ordinal] = typeof(object);
            }
            return types;
        }

        private sealed class LegacyXlsBufferedRow {
            internal LegacyXlsBufferedRow(
                ValueKind[] kinds,
                double[] numbers,
                DateTime[] dates,
                bool[] booleans,
                string?[] strings) {
                Kinds = kinds;
                Numbers = numbers;
                Dates = dates;
                Booleans = booleans;
                Strings = strings;
            }

            internal ValueKind[] Kinds { get; }
            internal double[] Numbers { get; }
            internal DateTime[] Dates { get; }
            internal bool[] Booleans { get; }
            internal string?[] Strings { get; }
        }
    }
}
