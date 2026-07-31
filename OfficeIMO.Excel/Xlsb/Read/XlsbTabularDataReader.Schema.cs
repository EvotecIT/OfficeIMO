namespace OfficeIMO.Excel.Xlsb.Read {
    internal sealed partial class XlsbTabularDataReader {
        private readonly Type[] _columnTypes;
        private readonly List<XlsbBufferedRow>? _schemaRows;
        private int _schemaRowIndex;

        private List<XlsbBufferedRow> BufferSchemaRows() {
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

            var rows = new List<XlsbBufferedRow>(Math.Min(_options.SchemaSampleRows, 256));
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

                rows.Add(new XlsbBufferedRow(
                    (XlsbTabularValueKind[])_kinds.Clone(),
                    (double[])_numbers.Clone(),
                    (bool[])_booleans.Clone(),
                    (string?[])_strings.Clone(),
                    (object?[])_customValues.Clone()));
                for (int ordinal = 0; ordinal < FieldCount; ordinal++) {
                    if (mixed[ordinal]
                        || _kinds[ordinal] == XlsbTabularValueKind.Empty
                        || _kinds[ordinal] == XlsbTabularValueKind.Custom
                        && IsMissingCustomValue(_customValues[ordinal])) {
                        continue;
                    }

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
            Array.Clear(_customValues, 0, _customValues.Length);
            return rows;
        }

        private void LoadBufferedRow(XlsbBufferedRow row) {
            Array.Copy(row.Kinds, _kinds, FieldCount);
            Array.Copy(row.Numbers, _numbers, FieldCount);
            Array.Copy(row.Booleans, _booleans, FieldCount);
            Array.Copy(row.Strings, _strings, FieldCount);
            Array.Copy(row.CustomValues, _customValues, FieldCount);
        }

        private static Type[] CreateObjectColumnTypes(int fieldCount) {
            var types = new Type[fieldCount];
            for (int ordinal = 0; ordinal < types.Length; ordinal++) {
                types[ordinal] = typeof(object);
            }

            return types;
        }

        private sealed class XlsbBufferedRow {
            internal XlsbBufferedRow(
                XlsbTabularValueKind[] kinds,
                double[] numbers,
                bool[] booleans,
                string?[] strings,
                object?[] customValues) {
                Kinds = kinds;
                Numbers = numbers;
                Booleans = booleans;
                Strings = strings;
                CustomValues = customValues;
            }

            internal XlsbTabularValueKind[] Kinds { get; }

            internal double[] Numbers { get; }

            internal bool[] Booleans { get; }

            internal string?[] Strings { get; }

            internal object?[] CustomValues { get; }
        }
    }
}
