#nullable enable

using System.Data.Common;
using System.Threading;

namespace OfficeIMO.CSV;

public sealed partial class CsvDocument
{
    // Large explicit schema samples are faster when streamed twice than kept live through reader traversal.
    private const int StreamingInferredReaderBufferLimit = 1000;
    private const long MemoryBackedCsvFileLimit = 32L * 1024 * 1024;
    private const int StreamingDataReaderFileBufferSize = 128 * 1024;
    private const int Utf8StreamingDataReaderFileBufferSize = 1;

    /// <summary>
    /// Creates a forward-only data reader over already-decoded CSV text.
    /// </summary>
    /// <param name="text">Decoded CSV text.</param>
    /// <param name="loadOptions">CSV load options. <see cref="CsvLoadOptions.MaxInputBytes"/> is measured using UTF-8, matching the prior text-import stream contract.</param>
    /// <param name="readerOptions">Reader projection options. When omitted, all columns are emitted as strings.</param>
    /// <returns>A data reader suitable for DataTable loading and provider bulk-copy APIs.</returns>
    public static DbDataReader OpenTextDataReader(
        string text,
        CsvLoadOptions? loadOptions = null,
        CsvDataReaderOptions? readerOptions = null)
    {
        if (text == null)
        {
            throw new ArgumentNullException(nameof(text));
        }

        var options = loadOptions?.Clone() ?? new CsvLoadOptions();
        if (options.MaxInputBytes <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(loadOptions),
                "MaxInputBytes must be greater than zero.");
        }

        var utf8 = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
        // A UTF-16 code unit expands to at most three UTF-8 bytes (a surrogate pair
        // expands to four bytes total). Avoid scanning ordinary bounded inputs merely
        // to prove that they fit comfortably below the configured byte limit.
        if (text.Length > options.MaxInputBytes / 3 &&
            utf8.GetByteCount(text) > options.MaxInputBytes)
        {
            throw new InvalidDataException(
                $"CSV data exceeds the configured maximum size ({options.MaxInputBytes} bytes).");
        }

        options.Mode = CsvLoadMode.Stream;
        readerOptions ??= new CsvDataReaderOptions();
        if (readerOptions.SchemaSampleSize <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(readerOptions),
                "Schema sample size must be greater than zero.");
        }

#if NET8_0_OR_GREATER
        if (TryCreateHeaderlessTextDataReader(text, options, readerOptions, out CsvDataReader? textDataReader))
        {
            return CsvParallelDataReader.Apply(textDataReader!, readerOptions);
        }
#endif

        CsvDocument document = LoadInternal(
            () => new StringReader(text), options, utf8, text);
        return document.CreateDataReader(readerOptions);
    }

    /// <summary>
    /// Creates a forward-only data reader over a CSV file.
    /// </summary>
    /// <param name="path">Source CSV path.</param>
    /// <param name="loadOptions">CSV load options.</param>
    /// <param name="readerOptions">Reader projection options. When omitted, all columns are emitted as strings.</param>
    /// <returns>A data reader suitable for DataTable loading and provider bulk-copy APIs.</returns>
    public static DbDataReader OpenDataReader(string path, CsvLoadOptions? loadOptions = null, CsvDataReaderOptions? readerOptions = null)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("File path cannot be empty.", nameof(path));
        }

        var options = loadOptions?.Clone() ?? new CsvLoadOptions();
        // OpenDataReader is the forward-only entry point. Reader routing remains
        // internal so callers cannot accidentally request streaming through Load.
        options.Mode = CsvLoadMode.Stream;
        readerOptions ??= new CsvDataReaderOptions();
        if (readerOptions.SchemaSampleSize <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(readerOptions), "Schema sample size must be greater than zero.");
        }

#if NET8_0_OR_GREATER
        if (CanUseStreamSpanDataReader(options, readerOptions)
            && CsvFile.ResolveCompression(options.CompressionType, path) == CsvCompressionType.None)
        {
            if (CanUseUtf8StreamDataReader(options))
            {
                Stream utf8Stream = CsvFile.OpenReadStream(
                    path,
                    options,
                    Utf8StreamingDataReaderFileBufferSize,
                    useAsync: false);
                if (CsvParser.CsvUtf8StreamDataReaderRowSource.TryCreate(
                        utf8Stream,
                        options,
                        out CsvParser.CsvUtf8StreamDataReaderRowSource? utf8Rows))
                {
                    if (TryCreateStreamSpanDataReader(utf8Rows!, options, readerOptions, out CsvDataReader? utf8DataReader))
                    {
                        return CsvParallelDataReader.Apply(utf8DataReader!, readerOptions);
                    }
                }
                else
                {
                    utf8Stream.Dispose();
                }
            }

            var spanReader = CsvFile.OpenTextReader(path, options, StreamingDataReaderFileBufferSize);
            if (TryCreateStreamSpanDataReader(spanReader, options, readerOptions, out CsvDataReader? dataReader))
            {
                return CsvParallelDataReader.Apply(dataReader!, readerOptions);
            }
        }

        if (CanUseMemoryBackedFileDataReader(path, options, readerOptions))
        {
            using var boundedReader = CsvFile.OpenTextReader(path, options, FileBufferSize);
            var text = ReadAllTextWithCancellation(
                boundedReader,
                options.CancellationToken);
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
                return CsvParallelDataReader.Apply(CreateEmptyDataReader(readerOptions, options), readerOptions);
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
            return CsvParallelDataReader.Apply(new CsvDataReader(
                columns,
                rows,
                header.Count - (options.StaticColumns?.Count ?? 0),
                options,
                options.Culture,
                options.DateTimeFormats,
                rowOwner), readerOptions);
        }
        catch
        {
            records?.Dispose();
            reader.Dispose();
            throw;
        }
    }

    /// <summary>
    /// Opens a forward-only data reader over a CSV stream.
    /// </summary>
    /// <param name="stream">Readable CSV stream. The stream remains open after the reader is disposed.</param>
    /// <param name="loadOptions">CSV load options.</param>
    /// <param name="readerOptions">Reader projection options. When omitted, all columns are emitted as strings.</param>
    /// <returns>A data reader suitable for DataTable loading and provider bulk-copy APIs.</returns>
    public static DbDataReader OpenDataReader(
        Stream stream,
        CsvLoadOptions? loadOptions = null,
        CsvDataReaderOptions? readerOptions = null)
    {
        if (stream == null)
        {
            throw new ArgumentNullException(nameof(stream));
        }

        if (!stream.CanRead)
        {
            throw new ArgumentException("Stream must be readable.", nameof(stream));
        }

        var options = loadOptions?.Clone() ?? new CsvLoadOptions();
        options.Mode = CsvLoadMode.Stream;
        readerOptions ??= new CsvDataReaderOptions();
        if (readerOptions.SchemaSampleSize <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(readerOptions), "Schema sample size must be greater than zero.");
        }

        if (!stream.CanSeek)
        {
            return CanUseSinglePassFileDataReader(options, readerOptions)
                ? CreateSinglePassNonSeekableDataReader(stream, options, readerOptions)
                : Load(stream, options).CreateDataReader(readerOptions);
        }

        long startPosition = stream.Position;
        if (!CanUseSinglePassFileDataReader(options, readerOptions))
        {
            return CreateBufferedDataReaderFromCurrentPosition(stream, startPosition, options, readerOptions);
        }

#if NET8_0_OR_GREATER
        if (CanUseStreamSpanDataReader(options, readerOptions))
        {
            var spanReader = CsvFile.OpenTextReader(stream, options, leaveOpen: true, StreamingDataReaderFileBufferSize);
            if (TryCreateStreamSpanDataReader(spanReader, options, readerOptions, out CsvDataReader? dataReader))
            {
                return CsvParallelDataReader.Apply(dataReader!, readerOptions);
            }

            stream.Position = startPosition;
        }
#endif

        var reader = CsvFile.OpenTextReader(stream, options, leaveOpen: true, FileBufferSize);
        IEnumerator<IReadOnlyList<string>>? records = null;
        try
        {
            records = CsvParser.ParseReusable(reader, options).GetEnumerator();
            if (!records.MoveNext())
            {
                records.Dispose();
                reader.Dispose();
                return CsvParallelDataReader.Apply(CreateEmptyDataReader(readerOptions, options), readerOptions);
            }

            if (ShouldUseGeneralDataReaderForFirstHeaderRecord(records.Current, options))
            {
                records.Dispose();
                reader.Dispose();
                stream.Position = startPosition;
                return CreateBufferedDataReaderFromCurrentPosition(stream, startPosition, options, readerOptions);
            }

            var header = AppendStaticColumnsToHeader(NormalizeParsedHeader(records.Current, options), options);
            var columns = CreateDataReaderColumns(header, readerOptions);
            var rows = EnumerateRemainingStringRows(records);
            var rowOwner = new CsvFileDataReaderRowOwner(reader, records);
            records = null;
            return CsvParallelDataReader.Apply(new CsvDataReader(
                columns,
                rows,
                header.Count - (options.StaticColumns?.Count ?? 0),
                options,
                options.Culture,
                options.DateTimeFormats,
                rowOwner), readerOptions);
        }
        catch
        {
            records?.Dispose();
            reader.Dispose();
            throw;
        }
    }

    private static DbDataReader CreateSinglePassNonSeekableDataReader(
        Stream stream,
        CsvLoadOptions options,
        CsvDataReaderOptions readerOptions)
    {
        var reader = CsvFile.OpenTextReader(stream, options, leaveOpen: true, FileBufferSize);
        IEnumerator<CsvParser.CsvParsedRecord>? records = null;
        try
        {
            records = CsvParser.ParseWithMetadata(reader, options).GetEnumerator();
            if (!TryReadHeader(records, options, out IReadOnlyList<string> header, out _))
            {
                records.Dispose();
                reader.Dispose();
                return CsvParallelDataReader.Apply(CreateEmptyDataReader(readerOptions, options), readerOptions);
            }

            var columns = CreateDataReaderColumns(header, readerOptions);
            var rows = EnumerateRemainingParsedRows(records);
            var rowOwner = new CsvFileDataReaderRowOwner(reader, records);
            records = null;
            return CsvParallelDataReader.Apply(new CsvDataReader(
                columns,
                rows,
                header.Count - (options.StaticColumns?.Count ?? 0),
                options,
                options.Culture,
                options.DateTimeFormats,
                rowOwner), readerOptions);
        }
        catch
        {
            records?.Dispose();
            reader.Dispose();
            throw;
        }
    }

    private static DbDataReader CreateBufferedDataReaderFromCurrentPosition(
        Stream stream,
        long startPosition,
        CsvLoadOptions options,
        CsvDataReaderOptions readerOptions)
    {
        byte[] snapshot;
        try
        {
            snapshot = OfficeIMO.Core.Internal.OfficeStreamReader.ReadRemainingBytes(
                stream,
                options.CancellationToken,
                options.MaxInputBytes);
        }
        finally
        {
            stream.Position = startPosition;
        }

        using var snapshotStream = new MemoryStream(snapshot, writable: false);
        return Load(snapshotStream, options).CreateDataReader(readerOptions);
    }

    /// <summary>
    /// Creates a forward-only data reader over the document rows.
    /// </summary>
    /// <param name="options">Reader projection options. When omitted, all columns are emitted as strings.</param>
    /// <returns>A data reader suitable for DataTable loading and provider bulk-copy APIs.</returns>
    public DbDataReader CreateDataReader(CsvDataReaderOptions? options = null)
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
            return CsvParallelDataReader.Apply(
                CreateStreamingInferredDataReader(options.SchemaSampleSize),
                options);
        }

        var schema = options.Schema ?? _schema ?? (options.InferSchema ? InferSchema(options.SchemaSampleSize) : null);
        var columns = CreateDataReaderColumns(_header, schema);
        if (_mode == CsvLoadMode.Stream && _streamingSource is not null)
        {
            if (_streamingSource.TryCreateDataReaderTextRowSource(out var textRows))
            {
                return CsvParallelDataReader.Apply(new CsvDataReader(
                    columns,
                    textRows!,
                    _streamingSource.SourceColumnCount,
                    _streamingSource.Options,
                    _culture,
                    _dateTimeFormats), options);
            }

            return CsvParallelDataReader.Apply(new CsvDataReader(
                columns,
                _streamingSource.ReadReusableStringRows(),
                _streamingSource.SourceColumnCount,
                _streamingSource.Options,
                _culture,
                _dateTimeFormats), options);
        }

        var rows = EnumerateRawRows();
        return CsvParallelDataReader.Apply(new CsvDataReader(
            columns,
            rows,
            _culture,
            _dateTimeFormats,
            _delimiter,
            _mappingErrorValuePolicy,
            _rowsAreParsedStringsOnly), options);
    }

    private static bool CanUseSinglePassFileDataReader(CsvLoadOptions options, CsvDataReaderOptions readerOptions) =>
        options.Mode == CsvLoadMode.Stream &&
        options.HasHeaderRow &&
        options.Header is null &&
        options.SkipInitialRecords == 0 &&
        !options.DetectDelimiter &&
        (!readerOptions.InferSchema || readerOptions.Schema is not null);

#if NET8_0_OR_GREATER
    private static bool TryCreateHeaderlessTextDataReader(
        string text,
        CsvLoadOptions options,
        CsvDataReaderOptions readerOptions,
        out CsvDataReader? dataReader)
    {
        if (options.HasHeaderRow ||
            options.Header is not null ||
            options.DetectDelimiter ||
            readerOptions.Schema is not null ||
            readerOptions.InferSchema ||
            !CsvParser.CanReadDataReaderRowsFromText(text, options))
        {
            dataReader = null;
            return false;
        }

        var recordsToSkip = GetInitialRecordsToSkip(options);
        if (!CsvParser.TryGetFirstTextDataReaderRecordFieldCount(
            text,
            options,
            recordsToSkip,
            out var sourceColumnCount))
        {
            dataReader = CreateEmptyDataReader(readerOptions, options);
            return true;
        }

        var header = GenerateDefaultHeader(sourceColumnCount);
        var columns = CreateDataReaderColumns(header, readerOptions);
        var rows = new CsvParser.CsvTextDataReaderRowSource(
            text,
            options,
            recordsToSkip,
            sourceColumnCount);
        dataReader = new CsvDataReader(
            columns,
            rows,
            sourceColumnCount,
            options,
            options.Culture,
            options.DateTimeFormats);
        return true;
    }

    private static bool CanUseStreamSpanDataReader(
        CsvLoadOptions options,
        CsvDataReaderOptions readerOptions) =>
        CanUseSinglePassFileDataReader(options, readerOptions)
        && readerOptions.Schema is null
        && !readerOptions.InferSchema
        && options.StaticColumns is null
        && string.IsNullOrEmpty(options.DelimiterText)
        && options.MaxFieldLength is null
        && options.MaxQuotedFieldLength is null
        && !options.NormalizeQuotes
        && !options.InternStrings;

    private static bool CanUseUtf8StreamDataReader(CsvLoadOptions options)
    {
        char delimiter = CsvParser.GetDelimiterChar(options);
        return !options.TrimWhitespace &&
            (options.Encoding is null || options.Encoding.CodePage == Encoding.UTF8.CodePage) &&
            delimiter is > (char)0 and <= (char)127 &&
            delimiter is not '"' and not '\r' and not '\n' &&
            options.CommentCharacter is > (char)0 and <= (char)127;
    }

    private static bool TryCreateStreamSpanDataReader(
        TextReader reader,
        CsvLoadOptions options,
        CsvDataReaderOptions readerOptions,
        out CsvDataReader? dataReader)
    {
        var rows = new CsvParser.CsvStreamDataReaderRowSource(reader, options);
        return TryCreateStreamSpanDataReader(rows, options, readerOptions, out dataReader);
    }

    private static bool TryCreateStreamSpanDataReader(
        ICsvDataReaderHeaderRowSource rows,
        CsvLoadOptions options,
        CsvDataReaderOptions readerOptions,
        out CsvDataReader? dataReader)
    {
        try
        {
            if (!rows.Read())
            {
                rows.Dispose();
                dataReader = CreateEmptyDataReader(readerOptions, options);
                return true;
            }

            var firstRecord = new string[rows.FieldCount];
            for (int index = 0; index < firstRecord.Length; index++)
            {
                firstRecord[index] = rows.GetString(index);
            }

            if (ShouldUseGeneralDataReaderForFirstHeaderRecord(firstRecord, options))
            {
                rows.Dispose();
                dataReader = null;
                return false;
            }

            IReadOnlyList<string> header = NormalizeParsedHeader(firstRecord, options);
            rows.SetSourceColumnCount(header.Count);
            CsvDataColumnProjection[] columns = CreateDataReaderColumns(header, readerOptions);
            dataReader = new CsvDataReader(
                columns,
                rows,
                header.Count,
                options,
                options.Culture,
                options.DateTimeFormats);
            return true;
        }
        catch
        {
            rows.Dispose();
            throw;
        }
    }

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

    internal static string ReadAllTextWithCancellation(
        TextReader reader,
        CancellationToken cancellationToken)
    {
        if (reader == null)
        {
            throw new ArgumentNullException(nameof(reader));
        }

        var text = new StringBuilder();
        var buffer = new char[FileBufferSize];
        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();
            int count = reader.Read(buffer, 0, buffer.Length);
            if (count == 0)
            {
                break;
            }

            text.Append(buffer, 0, count);
        }

        cancellationToken.ThrowIfCancellationRequested();
        return text.ToString();
    }

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

    private static IEnumerable<IReadOnlyList<string>> EnumerateRemainingParsedRows(
        IEnumerator<CsvParser.CsvParsedRecord> records)
    {
        while (records.MoveNext())
        {
            yield return records.Current.Values;
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
                _streamingSource.Options.Delimiter,
                _streamingSource.Options.MappingErrorValuePolicy,
                rowOwner: rowOwner,
                processingCancellationToken: _streamingSource.Options.CancellationToken);
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
        private IDisposable? _records;

        internal CsvFileDataReaderRowOwner(TextReader reader, IDisposable records)
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
