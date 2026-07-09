#nullable enable

using System.Data.Common;

namespace OfficeIMO.CSV;

public sealed partial class CsvDocument
{
    // Large explicit schema samples are faster when streamed twice than kept live through reader traversal.
    private const int StreamingInferredReaderBufferLimit = 1000;

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
            options.SchemaSampleSize <= StreamingInferredReaderBufferLimit &&
            _mode == CsvLoadMode.Stream &&
            _streamingSource is not null)
        {
            return CreateStreamingInferredDataReader(options.SchemaSampleSize);
        }

        var schema = options.Schema ?? _schema ?? (options.InferSchema ? InferSchema(options.SchemaSampleSize) : null);
        var schemaColumns = schema?.Columns.ToDictionary(column => column.Name, StringComparer.OrdinalIgnoreCase);
        var columns = CsvDataProjectionBuilder.Create(_header, schemaColumns);
        if (_mode == CsvLoadMode.Stream && _streamingSource is not null)
        {
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
        var schemaColumns = readerOptions.Schema?.Columns.ToDictionary(column => column.Name, StringComparer.OrdinalIgnoreCase);
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
        var rows = _streamingSource!.ReadReusableRows().GetEnumerator();
        try
        {
            var sampledRows = new List<object?[]>(Math.Min(schemaSampleSize, 4096));
            var schema = InferSchema(rows, schemaSampleSize, sampledRows, cloneSampledRows: true);
            var schemaColumns = schema.Columns.ToDictionary(column => column.Name, StringComparer.OrdinalIgnoreCase);
            var columns = CsvDataProjectionBuilder.Create(_header, schemaColumns);
            return new CsvDataReader(columns, EnumerateSampledThenRemainingRows(sampledRows, rows), _culture, _dateTimeFormats);
        }
        catch
        {
            rows.Dispose();
            throw;
        }
    }

    private static IEnumerable<object?[]> EnumerateSampledThenRemainingRows(IReadOnlyList<object?[]> sampledRows, IEnumerator<object?[]> remainingRows)
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
