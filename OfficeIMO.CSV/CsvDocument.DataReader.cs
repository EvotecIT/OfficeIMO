using System.Data.Common;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.CSV;

public sealed partial class CsvDocument
{
    // Large explicit schema samples are faster when streamed twice than kept live through reader traversal.
    private const int StreamingInferredReaderBufferLimit = 1000;
    private const long MemoryBackedCsvFileLimit = 32L * 1024 * 1024;
    private const int StreamingDataReaderFileBufferSize = 128 * 1024;
    private const int Utf8StreamingDataReaderFileBufferSize = 1;

    /// <summary>
    /// Creates a forward-only data reader over a CSV file.
    /// </summary>
    /// <param name="path">Source CSV path.</param>
    /// <param name="options">Parsing and projection settings. Columns are strings unless a schema is supplied or inferred.</param>
    /// <returns>A data reader suitable for DataTable loading and provider bulk-copy APIs.</returns>
    public static CsvDataReader OpenReader(string path, CsvReaderOptions? options = null)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("File path cannot be empty.", nameof(path));
        }

        options = options?.Clone() ?? new CsvReaderOptions();
        // OpenReader is the forward-only entry point. Reader routing remains
        // internal so callers cannot accidentally request streaming through Load.
        options.Mode = CsvLoadMode.Stream;
        if (options.SchemaSampleRows <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(options), "Schema sample size must be greater than zero.");
        }

#if NET8_0_OR_GREATER
        if (CanUseStreamSpanDataReader(options)
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
                    if (TryCreateStreamSpanDataReader(utf8Rows!, options, out CsvDataReader? utf8DataReader))
                    {
                        return CsvParallelDataReader.Apply(utf8DataReader!, options);
                    }
                }
                else
                {
                    utf8Stream.Dispose();
                }
            }

            var spanReader = CsvFile.OpenTextReader(path, options, StreamingDataReaderFileBufferSize);
            if (TryCreateStreamSpanDataReader(spanReader, options, out CsvDataReader? dataReader))
            {
                return CsvParallelDataReader.Apply(dataReader!, options);
            }
        }

        if (CanUseMemoryBackedFileDataReader(path, options))
        {
            using var boundedReader = CsvFile.OpenTextReader(path, options, FileBufferSize);
            var text = ReadAllTextWithCancellation(
                boundedReader,
                options.CancellationToken);
            return ParseText(text, options.Parsing).CreateDataReader(options);
        }
#endif

        if (!CanUseSinglePassFileDataReader(options))
        {
            return LoadFromPath(path, options.Parsing).CreateDataReader(options);
        }

        options = ResolveReaderOptions(() => CsvFile.OpenTextReader(path, options, FileBufferSize), options);
        var reader = CsvFile.OpenTextReader(path, options, FileBufferSize);
        IEnumerator<IReadOnlyList<string>>? records = null;

        try
        {
            records = CsvParser.ParseReusable(reader, options).GetEnumerator();
            if (!records.MoveNext())
            {
                records.Dispose();
                reader.Dispose();
                return CsvParallelDataReader.Apply(CreateEmptyDataReader(options), options);
            }

            if (ShouldUseGeneralDataReaderForFirstHeaderRecord(records.Current, options))
            {
                records.Dispose();
                reader.Dispose();
                return LoadFromPath(path, options.Parsing).CreateDataReader(options);
            }

            var header = AppendStaticColumnsToHeader(NormalizeParsedHeader(records.Current, options), options);
            var columns = CreateDataReaderColumns(header, options);
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
                rowOwner), options);
        }
        catch
        {
            records?.Dispose();
            reader.Dispose();
            throw;
        }
    }

#if NET8_0_OR_GREATER
    private static void ValidateAsyncReaderOptions(CsvReaderOptions? options)
    {
        if (options is { SchemaSampleRows: <= 0 })
        {
            throw new ArgumentOutOfRangeException(nameof(options), "Schema sample size must be greater than zero.");
        }
    }
#endif

    /// <summary>
    /// Opens a forward-only data reader over a CSV stream.
    /// </summary>
    /// <param name="stream">Readable CSV stream. The stream remains open after the reader is disposed.</param>
    /// <param name="options">Parsing and projection settings. Columns are strings unless a schema is supplied or inferred.</param>
    /// <returns>A data reader suitable for DataTable loading and provider bulk-copy APIs.</returns>
    public static CsvDataReader OpenReader(
        Stream stream,
        CsvReaderOptions? options = null)
    {
        if (stream == null)
        {
            throw new ArgumentNullException(nameof(stream));
        }

        if (!stream.CanRead)
        {
            throw new ArgumentException("Stream must be readable.", nameof(stream));
        }

        options = options?.Clone() ?? new CsvReaderOptions();
        options.Mode = CsvLoadMode.Stream;
        if (options.SchemaSampleRows <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(options), "Schema sample size must be greater than zero.");
        }

        if (!stream.CanSeek)
        {
            return CanUseSinglePassFileDataReader(options)
                ? CreateSinglePassNonSeekableDataReader(stream, options)
                : LoadFromStream(stream, options.Parsing).CreateDataReader(options);
        }

        long startPosition = stream.Position;
        if (!CanUseSinglePassFileDataReader(options)
#if NET8_0_OR_GREATER
            && !CanUseStreamSpanDataReader(options)
#endif
            )
        {
            return CreateBufferedDataReaderFromCurrentPosition(stream, startPosition, options);
        }

#if NET8_0_OR_GREATER
        if (CanUseStreamSpanDataReader(options))
        {
            if (CanUseUtf8StreamDataReader(options))
            {
                Stream utf8Stream = CsvFile.OpenReadStream(stream, options, leaveOpen: true);
                if (CsvParser.CsvUtf8StreamDataReaderRowSource.TryCreate(utf8Stream, options, out var utf8Rows))
                {
                    if (TryCreateStreamSpanDataReader(utf8Rows!, options, out CsvDataReader? utf8DataReader))
                    {
                        return CsvParallelDataReader.Apply(utf8DataReader!, options);
                    }
                }
                else
                {
                    utf8Stream.Dispose();
                }

                stream.Position = startPosition;
            }

            var spanReader = CsvFile.OpenTextReader(stream, options, leaveOpen: true, StreamingDataReaderFileBufferSize);
            if (TryCreateStreamSpanDataReader(spanReader, options, out CsvDataReader? dataReader))
            {
                return CsvParallelDataReader.Apply(dataReader!, options);
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
                return CsvParallelDataReader.Apply(CreateEmptyDataReader(options), options);
            }

            if (ShouldUseGeneralDataReaderForFirstHeaderRecord(records.Current, options))
            {
                records.Dispose();
                reader.Dispose();
                stream.Position = startPosition;
                return CreateBufferedDataReaderFromCurrentPosition(stream, startPosition, options);
            }

            var header = AppendStaticColumnsToHeader(NormalizeParsedHeader(records.Current, options), options);
            var columns = CreateDataReaderColumns(header, options);
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
                rowOwner), options);
        }
        catch
        {
            records?.Dispose();
            reader.Dispose();
            throw;
        }
    }

    private static CsvDataReader CreateSinglePassNonSeekableDataReader(
        Stream stream,
        CsvReaderOptions options)
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
                return CsvParallelDataReader.Apply(CreateEmptyDataReader(options), options);
            }

            var columns = CreateDataReaderColumns(header, options);
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
                rowOwner), options);
        }
        catch
        {
            records?.Dispose();
            reader.Dispose();
            throw;
        }
    }

    private static CsvDataReader CreateBufferedDataReaderFromCurrentPosition(
        Stream stream,
        long startPosition,
        CsvReaderOptions options)
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
        return LoadFromStream(snapshotStream, options.Parsing).CreateDataReader(options);
    }

    /// <summary>
    /// Creates a forward-only data reader over the document rows.
    /// </summary>
    /// <param name="options">Reader projection options. When omitted, all columns are emitted as strings.</param>
    /// <returns>A data reader suitable for DataTable loading and provider bulk-copy APIs.</returns>
    public CsvDataReader CreateDataReader(CsvReaderOptions? options = null) =>
        CreateDataReader(options, options?.CancellationToken ?? CancellationToken.None);

    internal CsvDataReader CreateDataReader(
        CsvReaderOptions? options,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        options ??= new CsvReaderOptions();
        if (options.SchemaSampleRows <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(options), "Schema sample size must be greater than zero.");
        }

        if (options.Schema is null &&
            _schema is null &&
            options.InferSchema &&
            _mode == CsvLoadMode.Stream &&
            _streamingSource is not null &&
            (options.SchemaSampleRows <= StreamingInferredReaderBufferLimit ||
                _streamingSource.CanCreateDataReaderTextRowSource))
        {
            return CsvParallelDataReader.Apply(
                CreateStreamingInferredDataReader(options.SchemaSampleRows, cancellationToken),
                options);
        }

        var schema = options.Schema ?? _schema ?? (options.InferSchema
            ? InferSchema(options.SchemaSampleRows, cancellationToken)
            : null);
        cancellationToken.ThrowIfCancellationRequested();
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

            CsvParsingSettings readerLoadOptions = _streamingSource.Options.Clone();
            return CsvParallelDataReader.Apply(new CsvDataReader(
                columns,
                _streamingSource.ReadReusableStringRows(readerLoadOptions),
                _streamingSource.SourceColumnCount,
                readerLoadOptions,
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
            _rowsAreParsedStringsOnly,
            processingCancellationToken: cancellationToken,
            delimiterText: DelimiterText), options);
    }

    private static bool CanUseSinglePassFileDataReader(CsvReaderOptions options) =>
        options.Mode == CsvLoadMode.Stream &&
        options.HasHeaderRow &&
        options.Header is null &&
        options.SkipInitialRecords == 0 &&
        !options.DetectDelimiter &&
        (!options.InferSchema || options.Schema is not null);

#if NET8_0_OR_GREATER
    private static bool TryCreateHeaderlessTextDataReader(
        string text,
        CsvReaderOptions options,
        out CsvDataReader? dataReader)
    {
        if (options.HasHeaderRow ||
            options.Header is not null ||
            options.DetectDelimiter ||
            options.Schema is not null ||
            options.InferSchema ||
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
            dataReader = CreateEmptyDataReader(options);
            return true;
        }

        var header = GenerateDefaultHeader(sourceColumnCount);
        var columns = CreateDataReaderColumns(header, options);
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
        CsvReaderOptions options) =>
        (CanUseSinglePassFileDataReader(options)
            || (options.Mode == CsvLoadMode.Stream
                && (options.Header is not null || !options.HasHeaderRow)
                && options.SkipInitialRecords == 0
                && !options.DetectDelimiter))
        && options.Schema is null
        && !options.InferSchema
        && options.StaticColumns is null
        && string.IsNullOrEmpty(options.DelimiterText)
        && options.MaxFieldLength is null
        && options.MaxQuotedFieldLength is null
        && !options.NormalizeQuotes
        && !options.InternStrings;

    private static bool CanUseUtf8StreamDataReader(CsvReaderOptions options)
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
        CsvReaderOptions options,
        out CsvDataReader? dataReader)
    {
        var rows = new CsvParser.CsvStreamDataReaderRowSource(reader, options);
        return TryCreateStreamSpanDataReader(rows, options, out dataReader);
    }

    private static bool TryCreateStreamSpanDataReader(
        ICsvDataReaderHeaderRowSource rows,
        CsvReaderOptions options,
        out CsvDataReader? dataReader)
    {
        try
        {
            IReadOnlyList<string>? explicitHeader = NormalizeExplicitHeader(options);
            if (explicitHeader is not null)
            {
                rows.SetSourceColumnCount(explicitHeader.Count);
                dataReader = new CsvDataReader(
                    CreateDataReaderColumns(explicitHeader, options),
                    rows,
                    explicitHeader.Count,
                    options,
                    options.Culture,
                    options.DateTimeFormats);
                return true;
            }

            if (!rows.Read())
            {
                rows.Dispose();
                dataReader = CreateEmptyDataReader(options);
                return true;
            }

            if (!options.HasHeaderRow)
            {
                IReadOnlyList<string> generatedHeader = GenerateDefaultHeader(rows.FieldCount);
                rows.SetSourceColumnCount(generatedHeader.Count);
                dataReader = new CsvDataReader(
                    CreateDataReaderColumns(generatedHeader, options),
                    rows,
                    generatedHeader.Count,
                    options,
                    options.Culture,
                    options.DateTimeFormats,
                    hasBufferedTextRow: true);
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
            CsvDataColumnProjection[] columns = CreateDataReaderColumns(header, options);
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
        CsvReaderOptions options)
    {
        return (options.Schema is not null || options.InferSchema) &&
            CanUseMemoryBackedFileText(path, options);
    }

    private static bool CanUseMemoryBackedFileText(string path, CsvParsingSettings options)
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

    private static bool ShouldUseGeneralDataReaderForFirstHeaderRecord(IReadOnlyList<string> record, CsvReaderOptions options)
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

    private static CsvDataReader CreateEmptyDataReader(CsvReaderOptions options)
    {
        var columns = CreateDataReaderColumns(Array.Empty<string>(), options);
        return new CsvDataReader(columns, Array.Empty<IReadOnlyList<string>>(), sourceColumnCount: 0, options, options.Culture, options.DateTimeFormats);
    }

    private static CsvDataColumnProjection[] CreateDataReaderColumns(IReadOnlyList<string> header, CsvReaderOptions options)
    {
        return CreateDataReaderColumns(header, options.Schema);
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

    private CsvDataReader CreateStreamingInferredDataReader(
        int schemaSampleSize,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
#if NET8_0_OR_GREATER
        if (_streamingSource!.TryCreateDataReaderTextRowSource(out var inferenceRows))
        {
            var rowsForInference = inferenceRows!;
            CsvSchema schema;
            using (rowsForInference)
            {
                schema = InferSchema(
                    rowsForInference,
                    schemaSampleSize,
                    _streamingSource.Options.NullValue,
                    cancellationToken);
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

        CsvParsingSettings inferenceOptions = _streamingSource!.Options.Clone();
        return CreateStreamingStringInferredDataReader(schemaSampleSize, inferenceOptions, cancellationToken);
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
