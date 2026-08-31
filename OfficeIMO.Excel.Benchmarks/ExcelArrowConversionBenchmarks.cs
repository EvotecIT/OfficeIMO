using System.Data.Common;
using Apache.Arrow;
using BenchmarkDotNet.Attributes;
using ExcelReader.Arrow;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Reader;
using OfficeIMO.Benchmarks;
using OfficeIMO.Data.Arrow;
using ExcelReaderApi = ExcelReader.Core.Reader.Excel;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>
/// Equivalent one-batch Arrow conversion over the hash-pinned 65K workbook.
/// The explicit-schema lane makes both adapters produce the same Arrow schema;
/// the inferred lane retains each library's public inference path.
/// </summary>
[MemoryDiagnoser]
[RankColumn]
public class ExcelArrowConversionBenchmarks {
    private const int SchemaSampleRows = 1_024;
    private const int StreamingBatchSize = 8_192;

    private static readonly string[] Headers = [
        "Region",
        "Country",
        "Item Type",
        "Sales Channel",
        "Order Priority",
        "Order Date",
        "Order ID",
        "Ship Date",
        "Units Sold",
        "Unit Price",
        "Unit Cost",
        "Total Revenue",
        "Total Cost",
        "Total Profit"
    ];

    private static readonly Type[] ExplicitClrTypes = [
        typeof(string),
        typeof(string),
        typeof(string),
        typeof(string),
        typeof(string),
        typeof(DateTime),
        typeof(long),
        typeof(DateTime),
        typeof(long),
        typeof(double),
        typeof(double),
        typeof(double),
        typeof(double),
        typeof(double)
    ];

    private ExcelColumnSchema[] _explicitExcelReaderSchema = null!;
    private ArrowConversionObservation _expectedExplicit;
    private ArrowConversionObservation _expectedInferred;
    private ArrowConversionObservation _excelReaderInferred;
    private string? _officeInferredSchemaDescription;
    private string? _excelReaderInferredSchemaDescription;

    [GlobalSetup]
    public void Setup() {
        MarkPflug65KFixture.EnsureAuthentic(MarkPflug65KFixture.XlsxFileName);
        _explicitExcelReaderSchema = CreateExplicitExcelReaderSchema();

        _expectedExplicit = OfficeIMO_ExplicitSchema();
        ValidateExpected(nameof(OfficeIMO_ExplicitSchema), _expectedExplicit);
        ValidateEquivalent(
            "explicit schema",
            _expectedExplicit,
            ExcelReaderNet_ExplicitSchema());

        _expectedInferred = OfficeIMO_InferredSchema();
        ValidateExpected(nameof(OfficeIMO_InferredSchema), _expectedInferred);
        _excelReaderInferred = ExcelReaderNet_InferredSchema();
        ValidateExpected(nameof(ExcelReaderNet_InferredSchema), _excelReaderInferred);
        ValidateEquivalentPayload(
            "inferred schema",
            _expectedInferred,
            _excelReaderInferred);

        ArrowConversionObservation bounded = OfficeIMO_BoundedStreaming();
        ValidateExpected(nameof(OfficeIMO_BoundedStreaming), bounded);
        if (bounded.Batches <= 1) {
            throw new InvalidDataException(
                $"The bounded Arrow lane produced {bounded.Batches} batch instead of exercising streaming.");
        }
        if (bounded with { Batches = 1 } != _expectedExplicit) {
            throw new InvalidDataException(
                $"The bounded Arrow lane produced {bounded}; the one-batch lane produced {_expectedExplicit}.");
        }
    }

    [Benchmark]
    [BenchmarkCategory("Arrow", "ExplicitSchema")]
    public ArrowConversionObservation OfficeIMO_ExplicitSchema() =>
        ConvertOfficeIMO(inferSchema: false, batchSize: MarkPflug65KFixture.ExpectedRows);

    [Benchmark]
    [BenchmarkCategory("Arrow", "ExplicitSchema")]
    public ArrowConversionObservation ExcelReaderNet_ExplicitSchema() =>
        ConvertExcelReader(_explicitExcelReaderSchema);

    [Benchmark]
    [BenchmarkCategory("Arrow", "InferredSchema")]
    public ArrowConversionObservation OfficeIMO_InferredSchema() =>
        ConvertOfficeIMO(inferSchema: true, batchSize: MarkPflug65KFixture.ExpectedRows);

    [Benchmark]
    [BenchmarkCategory("Arrow", "InferredSchema")]
    public ArrowConversionObservation ExcelReaderNet_InferredSchema() {
        using IExcelRowReader reader = ExcelReaderApi.FromFile(MarkPflug65KFixture.XlsxPath);
        ExcelColumnSchema[] schema = ExcelReaderApi.InferSchema(reader, headerRow: 1, SchemaSampleRows);
        using RecordBatch batch = reader.ToArrowRecordBatch(schema, headerRow: 1);
        _excelReaderInferredSchemaDescription ??= DescribeSchema(batch.Schema);
        return Observe(batch, batches: 1);
    }

    [Benchmark]
    [BenchmarkCategory("Arrow", "BoundedStreaming")]
    public ArrowConversionObservation OfficeIMO_BoundedStreaming() =>
        ConvertOfficeIMO(inferSchema: false, batchSize: StreamingBatchSize);

    internal void ValidateExplicit(ArrowConversionObservation observation) =>
        ValidateEquivalent("explicit schema", _expectedExplicit, observation);

    internal void ValidateInferred(ArrowConversionObservation observation) =>
        ValidateEquivalent("inferred schema", _expectedInferred, observation);

    internal void EnsureInferredSchemaIsComparable() {
        if (_expectedInferred.SchemaChecksum != _excelReaderInferred.SchemaChecksum) {
            throw new InvalidDataException(
                "The inferred Arrow payloads are semantically equal but their schemas differ, so the timed ratio is withheld. " +
                $"OfficeIMO schema checksum {_expectedInferred.SchemaChecksum}; " +
                $"ExcelReader.NET schema checksum {_excelReaderInferred.SchemaChecksum}. " +
                $"OfficeIMO schema: {_officeInferredSchemaDescription}. " +
                $"ExcelReader.NET schema: {_excelReaderInferredSchemaDescription}.");
        }
    }

    private ArrowConversionObservation ConvertOfficeIMO(bool inferSchema, int batchSize) {
        using DbDataReader reader = ExcelDocument.OpenDataReader(
            MarkPflug65KFixture.XlsxPath,
            new ExcelReadOptions {
                InferSchema = inferSchema,
                SchemaSampleRows = SchemaSampleRows,
                NumericAsDecimal = false
            });
        var options = new ArrowReadOptions {
            BatchSize = batchSize,
            ColumnTypes = inferSchema ? null : ExplicitClrTypes
        };
        var accumulator = new ExcelObservationAccumulator();
        ulong? schemaChecksum = null;
        int batches = 0;
        foreach (RecordBatch batch in reader.ReadArrowBatches(options)) {
            using (batch) {
                ulong currentSchemaChecksum = ComputeSchemaChecksum(batch.Schema);
                if (schemaChecksum.HasValue && schemaChecksum.Value != currentSchemaChecksum) {
                    throw new InvalidDataException("OfficeIMO changed Arrow schema between record batches.");
                }
                schemaChecksum = currentSchemaChecksum;
                if (inferSchema) {
                    _officeInferredSchemaDescription ??= DescribeSchema(batch.Schema);
                }
                ObserveRows(batch, ref accumulator);
                batches++;
            }
        }

        ExcelReadObservation data = accumulator.Build();
        return new ArrowConversionObservation(
            data.Rows,
            data.Cells,
            data.Checksum,
            schemaChecksum ?? 0UL,
            batches);
    }

    private ArrowConversionObservation ConvertExcelReader(ExcelColumnSchema[] schema) {
        using IExcelRowReader reader = ExcelReaderApi.FromFile(MarkPflug65KFixture.XlsxPath);
        using RecordBatch batch = reader.ToArrowRecordBatch(schema, headerRow: 1);
        return Observe(batch, batches: 1);
    }

    private static ArrowConversionObservation Observe(RecordBatch batch, int batches) {
        var accumulator = new ExcelObservationAccumulator();
        ObserveRows(batch, ref accumulator);
        ExcelReadObservation data = accumulator.Build();
        return new ArrowConversionObservation(
            data.Rows,
            data.Cells,
            data.Checksum,
            ComputeSchemaChecksum(batch.Schema),
            batches);
    }

    private static void ObserveRows(RecordBatch batch, ref ExcelObservationAccumulator observation) {
        if (batch.ColumnCount != MarkPflug65KFixture.ExpectedColumns) {
            throw new InvalidDataException(
                $"Arrow conversion produced {batch.ColumnCount} columns instead of {MarkPflug65KFixture.ExpectedColumns}.");
        }

        for (int row = 0; row < batch.Length; row++) {
            observation.BeginRow();
            for (int ordinal = 0; ordinal <= 4; ordinal++) {
                observation.Add(ReadString(batch.Column(ordinal), row));
            }
            observation.Add(ReadDateTime(batch.Column(5), row));
            observation.Add(ReadInt32(batch.Column(6), row));
            observation.Add(ReadDateTime(batch.Column(7), row));
            observation.Add(ReadInt32(batch.Column(8), row));
            for (int ordinal = 9; ordinal <= 13; ordinal++) {
                observation.Add(ReadDecimal(batch.Column(ordinal), row));
            }
        }
    }

    private static string ReadString(IArrowArray array, int index) =>
        array is StringArray values
            ? values.GetString(index) ?? throw MissingValue(index)
            : throw UnexpectedType(array, index);

    private static DateTime ReadDateTime(IArrowArray array, int index) =>
        array switch {
            TimestampArray values => values.GetTimestamp(index)?.DateTime ?? throw MissingValue(index),
            Date32Array values => values.GetDateTime(index) ?? throw MissingValue(index),
            _ => throw UnexpectedType(array, index)
        };

    private static int ReadInt32(IArrowArray array, int index) =>
        array switch {
            Int32Array values => values.GetValue(index) ?? throw MissingValue(index),
            Int64Array values => checked((int)(values.GetValue(index) ?? throw MissingValue(index))),
            DoubleArray values => checked((int)(values.GetValue(index) ?? throw MissingValue(index))),
            _ => throw UnexpectedType(array, index)
        };

    private static decimal ReadDecimal(IArrowArray array, int index) =>
        array switch {
            Decimal128Array values => values.GetValue(index) ?? throw MissingValue(index),
            DoubleArray values => (decimal)(values.GetValue(index) ?? throw MissingValue(index)),
            FloatArray values => (decimal)(values.GetValue(index) ?? throw MissingValue(index)),
            Int64Array values => values.GetValue(index) ?? throw MissingValue(index),
            Int32Array values => values.GetValue(index) ?? throw MissingValue(index),
            _ => throw UnexpectedType(array, index)
        };

    private static InvalidDataException MissingValue(int index) =>
        new($"Arrow conversion produced an unexpected null at row {index}.");

    private static InvalidDataException UnexpectedType(IArrowArray array, int index) =>
        new($"Arrow conversion produced unsupported array type '{array.GetType().Name}' at row {index}.");

    private static ulong ComputeSchemaChecksum(Schema schema) {
        const ulong offsetBasis = 14695981039346656037UL;
        const ulong prime = 1099511628211UL;
        ulong checksum = offsetBasis;
        foreach (Field field in schema.FieldsList) {
            Add(field.Name);
            Add(field.DataType.Name);
            Add(field.IsNullable ? "nullable" : "required");
        }
        return checksum;

        void Add(string value) {
            foreach (char character in value) {
                checksum ^= character;
                checksum *= prime;
            }
            checksum ^= (ulong)value.Length;
            checksum *= prime;
        }
    }

    private static string DescribeSchema(Schema schema) =>
        string.Join(
            ", ",
            schema.FieldsList.Select(static field =>
                $"{field.Name}:{field.DataType.Name}:{(field.IsNullable ? "nullable" : "required")}"));

    private static ExcelColumnSchema[] CreateExplicitExcelReaderSchema() {
        var types = new[] {
            ExcelColumnType.StringColumn,
            ExcelColumnType.StringColumn,
            ExcelColumnType.StringColumn,
            ExcelColumnType.StringColumn,
            ExcelColumnType.StringColumn,
            ExcelColumnType.TimestampColumn,
            ExcelColumnType.Int64Column,
            ExcelColumnType.TimestampColumn,
            ExcelColumnType.Int64Column,
            ExcelColumnType.Float64Column,
            ExcelColumnType.Float64Column,
            ExcelColumnType.Float64Column,
            ExcelColumnType.Float64Column,
            ExcelColumnType.Float64Column
        };
        var schema = new ExcelColumnSchema[types.Length];
        for (int ordinal = 0; ordinal < schema.Length; ordinal++) {
            schema[ordinal] = new ExcelColumnSchema {
                Index = ordinal,
                Name = Headers[ordinal],
                Type = types[ordinal],
                IsNullable = true
            };
        }
        return schema;
    }

    private static void ValidateExpected(string operation, ArrowConversionObservation observation) {
        ExcelReadObservation expected = MarkPflug65KXlsxBenchmarks.ExpectedObservation();
        if (observation.Rows != expected.Rows
            || observation.Cells != expected.Cells
            || observation.Checksum != expected.Checksum
            || observation.Batches <= 0) {
            throw new InvalidDataException(
                $"{operation} produced {observation}; expected {expected} and at least one record batch.");
        }
    }

    private static void ValidateEquivalent(
        string mode,
        ArrowConversionObservation officeIMO,
        ArrowConversionObservation excelReader) {
        if (officeIMO != excelReader) {
            throw new InvalidDataException(
                $"Arrow {mode} conversion was not equivalent: OfficeIMO={officeIMO}; " +
                $"ExcelReader.NET={excelReader}.");
        }
    }

    private static void ValidateEquivalentPayload(
        string mode,
        ArrowConversionObservation officeIMO,
        ArrowConversionObservation excelReader) {
        if (officeIMO.Rows != excelReader.Rows
            || officeIMO.Cells != excelReader.Cells
            || officeIMO.Checksum != excelReader.Checksum
            || officeIMO.Batches != excelReader.Batches) {
            throw new InvalidDataException(
                $"Arrow {mode} payload was not equivalent: OfficeIMO={officeIMO}; " +
                $"ExcelReader.NET={excelReader}.");
        }
    }
}

public readonly record struct ArrowConversionObservation(
    int Rows,
    int Cells,
    ulong Checksum,
    ulong SchemaChecksum,
    int Batches);
