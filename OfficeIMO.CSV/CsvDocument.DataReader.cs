#nullable enable

using System.Data.Common;

namespace OfficeIMO.CSV;

public sealed partial class CsvDocument
{
    // Large explicit schema samples are faster when streamed twice than kept live through reader traversal.
    private const int StreamingInferredReaderBufferLimit = 1000;
    private const long MemoryBackedCsvFileLimit = 32L * 1024 * 1024;

    /// <summary>
    /// Creates a forward-only data reader over a CSV file.
    /// </summary>
    /// <param name="path">Source CSV path.</param>
    /// <param name="loadOptions">CSV load options.</param>
    /// <param name="readerOptions">Reader projection options. When omitted, all columns are emitted as strings.</param>
    /// <returns>A data reader suitable for DataTable loading and provider bulk-copy APIs.</returns>
    public static CsvDataReader CreateDataReader(string path, CsvLoadOptions? loadOptions = null, CsvDataReaderOptions? readerOptions = null)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("File path cannot be empty.", nameof(path));
        }

        var options = loadOptions?.Clone() ?? new CsvLoadOptions();
        readerOptions ??= new CsvDataReaderOptions();
        if (readerOptions.SchemaSampleSize <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(readerOptions), "Schema sample size must be greater than zero.");
        }

#if NET8_0_OR_GREATER
        if (CanUseMemoryBackedFileDataReader(path, options, readerOptions))
        {
            options.CancellationToken.ThrowIfCancellationRequested();
            using var boundedReader = CsvFile.OpenTextReader(path, options, FileBufferSize);
            var text = boundedReader.ReadToEnd();
            return Parse(text, options).CreateDataReader(readerOptions);
        }
#endif

        if (!CanUseSinglePassFileDataReader(options, readerOptions))
        {
            return Load(path, options).CreateDataReader(readerOptions);
        }

        options = ResolveLoadOptions(() => CsvFile.OpenTextReader(path, options, FileBufferSize), options);
        var reader = CsvFile.OpenTextReader(path, options, FileBufferSize);
        IEnumerator<IReadOnlyList<string>>? records = null;

        try
        {
            records = CsvParser.ParseReusable(reader, options).GetEnumerator();
            if (!records.MoveNext())
            {
                records.Dispose();
                reader.Dispose();
                return CreateEmptyDataReader(readerOptions, options);
            }

            if (ShouldUseGeneralDataReaderForFirstHeaderRecord(records.Current, options))
            {
                records.Dispose();
                reader.Dispose();
                return Load(path, options).CreateDataReader(readerOptions);
            }

            var header = AppendStaticColumnsToHeader(NormalizeParsedHeader(records.Current, options), options);
            var columns = CreateDataReaderColumns(header, readerOptions);
            var rows = EnumerateRemainingStringRows(records);
            var rowOwner = new CsvFileDataReaderRowOwner(reader, records);
            records = null;
            return new CsvDataReader(
                columns,
                rows,
                header.Count - (options.StaticColumns?.Count ?? 0),
                options,
                options.Culture,
                options.DateTimeFormats,
                rowOwner);
        }
        catch
        {
            records?.Dispose();
            reader.Dispose();
            throw;
        }
    }

    /// <summary>
    /// Creates a forward-only data reader over the document rows.
    /// </summary>
    /// <param name="options">Reader projection options. When omitted, all columns are emitted as strings.</param>
    /// <returns>A data reader suitable for DataTable loading and provider bulk-copy APIs.</returns>
    public CsvDataReader CreateDataReader(CsvDataReaderOptions? options = null)
    {
        options ??= new CsvDataReaderOptions();
        if (options.SchemaSampleSize <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(options), "Schema sample size must be greater than zero.");
        }

        if (options.Schema is null &&
            _schema is null &&
            options.InferSchema &&
            _mode == CsvLoadMode.Stream &&
            _streamingSource is not null &&
            (options.SchemaSampleSize <= StreamingInferredReaderBufferLimit ||
                _streamingSource.CanCreateDataReaderTextRowSource))
        {
            return CreateStreamingInferredDataReader(options.SchemaSampleSize);
        }

        var schema = options.Schema ?? _schema ?? (options.InferSchema ? InferSchema(options.SchemaSampleSize) : null);
        var columns = CreateDataReaderColumns(_header, schema);
        if (_mode == CsvLoadMode.Stream && _streamingSource is not null)
        {
            if (_streamingSource.TryCreateDataReaderTextRowSource(out var textRows))
            {
                return new CsvDataReader(
                    columns,
                    textRows!,
                    _streamingSource.SourceColumnCount,
                    _streamingSource.Options,
                    _culture,
                    _dateTimeFormats);
            }

            return new CsvDataReader(
                columns,
                _streamingSource.ReadReusableStringRows(),
                _streamingSource.SourceColumnCount,
                _streamingSource.Options,
                _culture,
                _dateTimeFormats);
        }

        var rows = EnumerateRawRows();
        return new CsvDataReader(columns, rows, _culture, _dateTimeFormats, _rowsAreParsedStringsOnly);
    }

    private static bool CanUseSinglePassFileDataReader(CsvLoadOptions options, CsvDataReaderOptions readerOptions) =>
        options.Mode == CsvLoadMode.Stream &&
        options.HasHeaderRow &&
        options.Header is null &&
        options.SkipInitialRecords == 0 &&
        !options.DetectDelimiter &&
        (!readerOptions.InferSchema || readerOptions.Schema is not null);

#if NET8_0_OR_GREATER
    private static bool CanUseMemoryBackedFileDataReader(
        string path,
        CsvLoadOptions options,
        CsvDataReaderOptions readerOptions)
    {
        return (readerOptions.Schema is not null || readerOptions.InferSchema) &&
            CanUseMemoryBackedFileText(path, options);
    }

    private static bool CanUseMemoryBackedFileText(string path, CsvLoadOptions options)
    {
        if (options.Mode != CsvLoadMode.Stream ||
            CsvFile.ResolveCompression(options.CompressionType, path) != CsvCompressionType.None)
        {
            return false;
        }

        var fileLength = new FileInfo(path).Length;
        return fileLength <= MemoryBackedCsvFileLimit &&
            fileLength <= options.MaxInputBytes &&
            (options.MaxDecompressedBytes is null || fileLength <= options.MaxDecompressedBytes.Value);
    }
#endif

    private static bool ShouldUseGeneralDataReaderForFirstHeaderRecord(IReadOnlyList<string> record, CsvLoadOptions options)
    {
        if (record.Count == 0)
        {
            return false;
        }

        if (options.RecognizeW3CFieldsHeader && TryGetW3CFieldsHeader(record, options, out _))
        {
            return true;
        }

        return options.SkipCommentRowsBeforeHeader &&
            record[0].Length > 0 &&
            record[0][0] == options.CommentCharacter;
    }

    private static CsvDataReader CreateEmptyDataReader(CsvDataReaderOptions readerOptions, CsvLoadOptions options)
    {
        var columns = CreateDataReaderColumns(Array.Empty<string>(), readerOptions);
        return new CsvDataReader(columns, Array.Empty<IReadOnlyList<string>>(), sourceColumnCount: 0, options, options.Culture, options.DateTimeFormats);
    }

    private static CsvDataColumnProjection[] CreateDataReaderColumns(IReadOnlyList<string> header, CsvDataReaderOptions readerOptions)
    {
        return CreateDataReaderColumns(header, readerOptions.Schema);
    }

    private static CsvDataColumnProjection[] CreateDataReaderColumns(IReadOnlyList<string> header, CsvSchema? schema)
    {
        if (schema is null)
        {
            return CsvDataProjectionBuilder.Create(header, schemaColumns: null);
        }

        if (schema.Columns.Count == header.Count)
        {
            var namesMatchByOrdinal = true;
            for (var i = 0; i < header.Count; i++)
            {
                if (!string.Equals(header[i], schema.Columns[i].Name, StringComparison.OrdinalIgnoreCase))
                {
                    namesMatchByOrdinal = false;
                    break;
                }
            }

            if (namesMatchByOrdinal)
            {
                return CsvDataProjectionBuilder.CreateByOrdinal(header, schema.Columns);
            }
        }

        var schemaColumns = schema.Columns.ToDictionary(column => column.Name, StringComparer.OrdinalIgnoreCase);
        return CsvDataProjectionBuilder.Create(header, schemaColumns);
    }

    private static IEnumerable<IReadOnlyList<string>> EnumerateRemainingStringRows(
        IEnumerator<IReadOnlyList<string>> records)
    {
        while (records.MoveNext())
        {
            yield return records.Current;
        }
    }

    private CsvDataReader CreateStreamingInferredDataReader(int schemaSampleSize)
    {
#if NET8_0_OR_GREATER
        if (_streamingSource!.TryCreateDataReaderTextRowSource(out var inferenceRows))
        {
            var rowsForInference = inferenceRows!;
            CsvSchema schema;
            using (rowsForInference)
            {
                schema = InferSchema(rowsForInference, schemaSampleSize, _streamingSource.Options.NullValue);
            }

            var columns = CreateDataReaderColumns(_header, schema);
            if (_streamingSource.TryCreateDataReaderTextRowSource(out var typedRows))
            {
                return new CsvDataReader(
                    columns,
                    typedRows!,
                    _streamingSource.SourceColumnCount,
                    _streamingSource.Options,
                    _culture,
                    _dateTimeFormats);
            }
        }
#endif

        var rows = _streamingSource!.ReadReusableRows().GetEnumerator();
        try
        {
            var sampledRows = new List<object?[]>(Math.Min(schemaSampleSize, 4096));
            var schema = InferSchema(rows, schemaSampleSize, sampledRows, cloneSampledRows: true);
            var columns = CreateDataReaderColumns(_header, schema);
            var rowOwner = new CsvStreamingDataReaderRowOwner(rows);
            return new CsvDataReader(
                columns,
                EnumerateSampledThenRemainingRows(sampledRows, rowOwner),
                _culture,
                _dateTimeFormats,
                rowOwner: rowOwner);
        }
        catch
        {
            rows.Dispose();
            throw;
        }
    }

    private static IEnumerable<object?[]> EnumerateSampledThenRemainingRows(
        IReadOnlyList<object?[]> sampledRows,
        CsvStreamingDataReaderRowOwner remainingRows)
    {
        try
        {
            for (var i = 0; i < sampledRows.Count; i++)
            {
                yield return sampledRows[i];
            }

            while (remainingRows.MoveNext())
            {
                yield return remainingRows.Current;
            }
        }
        finally
        {
            remainingRows.Dispose();
        }
    }

    private sealed class CsvStreamingDataReaderRowOwner : IDisposable
    {
        private IEnumerator<object?[]>? _rows;

        internal CsvStreamingDataReaderRowOwner(IEnumerator<object?[]> rows)
        {
            _rows = rows;
        }

        internal object?[] Current => _rows!.Current;

        internal bool MoveNext() => _rows?.MoveNext() == true;

        public void Dispose()
        {
            _rows?.Dispose();
            _rows = null;
        }
    }

    private sealed class CsvFileDataReaderRowOwner : IDisposable
    {
        private TextReader? _reader;
        private IEnumerator<IReadOnlyList<string>>? _records;

        internal CsvFileDataReaderRowOwner(TextReader reader, IEnumerator<IReadOnlyList<string>> records)
        {
            _reader = reader;
            _records = records;
        }

        public void Dispose()
        {
            _records?.Dispose();
            _records = null;
            _reader?.Dispose();
            _reader = null;
        }
    }
}
