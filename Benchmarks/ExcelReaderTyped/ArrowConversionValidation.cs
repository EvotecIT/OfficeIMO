#if OFFICEIMO_BENCHMARK_ARROW && OFFICEIMO_BENCHMARK_NEW_APIS
using Apache.Arrow;
using Apache.Arrow.Arrays;
using Apache.Arrow.Types;
using System.Globalization;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks {
    internal static class ArrowConversionValidation {
        internal static void ValidateExplicit(RecordBatch batch, int offset, ArrowConversionScenario scenario, int totalRows) {
            bool wide = scenario == ArrowConversionScenario.CsvAllString;
            string[] names = wide ? ArrowConversionWorkload.WideHeaders : TypedWorkbookFixture.Headers;
            IArrowType[] types = wide ? Enumerable.Repeat<IArrowType>(StringType.Default, 8).ToArray()
                : [StringType.Default, Int64Type.Default, new TimestampType(TimeUnit.Microsecond, (string)null!), DoubleType.Default];
            ValidateSchema(batch, names, types);
            ValidateRange(batch, offset, totalRows);
            for (int row = 0; row < batch.Length; row++) {
                if (wide) {
                    for (int column = 0; column < 8; column++) {
                        string expected = ArrowConversionWorkload.Pool[(offset + row + column) % ArrowConversionWorkload.Pool.Length];
                        if (((StringArray)batch.Column(column)).GetString(row) != expected) Fail(offset + row, column);
                    }
                } else {
                    TypedRecord expected = TypedWorkbookFixture.ExpectedRecord(offset + row + 1);
                    if (((StringArray)batch.Column(0)).GetString(row) != expected.Name) Fail(offset + row, 0);
                    if (((Int64Array)batch.Column(1)).GetValue(row) != expected.Id) Fail(offset + row, 1);
                    if (((TimestampArray)batch.Column(2)).GetTimestamp(row)?.DateTime != expected.Date) Fail(offset + row, 2);
                    if (((DoubleArray)batch.Column(3)).GetValue(row) != expected.Value) Fail(offset + row, 3);
                }
            }
        }

        internal static void ValidatePeerInference(RecordBatch batch, int offset, int rows) {
            ValidateSchema(batch, TypedWorkbookFixture.Headers, Enumerable.Repeat<IArrowType>(StringType.Default, 4).ToArray());
            ValidateRange(batch, offset, rows);
            for (int row = 0; row < batch.Length; row++) {
                TypedRecord expected = TypedWorkbookFixture.ExpectedRecord(offset + row + 1);
                string[] values = [expected.Name!, expected.Id.ToString(CultureInfo.InvariantCulture),
                    expected.Date.ToString("O", CultureInfo.InvariantCulture), expected.Value.ToString(CultureInfo.InvariantCulture)];
                for (int column = 0; column < values.Length; column++)
                    if (((StringArray)batch.Column(column)).GetString(row) != values[column]) Fail(offset + row, column);
            }
        }

        internal static void ValidateOfficeInference(RecordBatch batch, int offset, int rows) {
            ValidateSchema(batch, TypedWorkbookFixture.Headers,
                [StringType.Default, Int32Type.Default, new TimestampType(TimeUnit.Microsecond, (string)null!), new Decimal128Type(29, 10)]);
            ValidateRange(batch, offset, rows);
            for (int row = 0; row < batch.Length; row++) {
                TypedRecord expected = TypedWorkbookFixture.ExpectedRecord(offset + row + 1);
                if (((StringArray)batch.Column(0)).GetString(row) != expected.Name) Fail(offset + row, 0);
                if (((Int32Array)batch.Column(1)).GetValue(row) != expected.Id) Fail(offset + row, 1);
                if (((TimestampArray)batch.Column(2)).GetTimestamp(row)?.DateTime != expected.Date) Fail(offset + row, 2);
                if (((Decimal128Array)batch.Column(3)).GetValue(row) != (decimal)expected.Value) Fail(offset + row, 3);
            }
        }

        private static void ValidateSchema(RecordBatch batch, string[] names, IArrowType[] types) {
            if (batch.ColumnCount != names.Length) throw new InvalidDataException("Arrow column count differs.");
            for (int ordinal = 0; ordinal < names.Length; ordinal++) {
                Field field = batch.Schema.GetFieldByIndex(ordinal);
                if (field.Name != names[ordinal] || field.IsNullable || field.DataType.TypeId != types[ordinal].TypeId
                    || batch.Column(ordinal).NullCount != 0)
                    throw new InvalidDataException($"Arrow schema or null values differ at column {ordinal}.");
                if (types[ordinal] is TimestampType expectedTimestamp && field.DataType is TimestampType timestamp
                    && (timestamp.Unit != expectedTimestamp.Unit || timestamp.Timezone != expectedTimestamp.Timezone))
                    throw new InvalidDataException("Arrow timestamp unit or timezone differs.");
                if (types[ordinal] is Decimal128Type expectedDecimal && field.DataType is Decimal128Type decimalType
                    && (decimalType.Precision != expectedDecimal.Precision || decimalType.Scale != expectedDecimal.Scale))
                    throw new InvalidDataException("Arrow decimal precision or scale differs.");
            }
        }

        private static void ValidateRange(RecordBatch batch, int offset, int rows) {
            if (offset < 0 || batch.Length < 1 || offset + batch.Length > rows)
                throw new InvalidDataException("Arrow row count or batch bounds differ.");
        }

        private static void Fail(int row, int column) =>
            throw new InvalidDataException($"Arrow value differs at zero-based row {row}, column {column}.");
    }
}
#endif