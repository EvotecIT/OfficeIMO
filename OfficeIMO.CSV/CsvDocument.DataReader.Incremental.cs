#nullable enable
#if NET8_0_OR_GREATER
using System.Data.Common;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.CSV;

public sealed partial class CsvDocument
{
    /// <summary>
    /// Opens a forward-only CSV reader whose ReadAsync calls perform incremental asynchronous I/O.
    /// Initialization reads the header and, when requested, at most SchemaSampleSize data records.
    /// Sequential projection and an explicit delimiter are required. Dispose the reader to close its file.
    /// </summary>
    public static Task<DbDataReader> OpenStreamingDataReaderAsync(string path,
        CsvLoadOptions? loadOptions = null, CsvDataReaderOptions? readerOptions = null,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        return CreateIncrementalReaderAsync(options => CsvFile.OpenTextReaderForAsyncRead(path, options, 4096),
            loadOptions, readerOptions, cancellationToken);
    }

    /// <summary>
    /// Opens an incremental CSV reader at the stream's current position. The stream remains open
    /// on failure and after disposal; its position advances and is not restored. Read and HasRows
    /// may perform synchronous I/O; use ReadAsync for asynchronous consumption. Cancellation or
    /// a parsing failure ends the reader. The opening token remains active for its lifetime.
    /// Sequential projection and an explicit delimiter are required.
    /// </summary>
    public static Task<DbDataReader> OpenStreamingDataReaderAsync(Stream stream,
        CsvLoadOptions? loadOptions = null, CsvDataReaderOptions? readerOptions = null,
        CancellationToken cancellationToken = default)
    {
        if (stream is null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));
        return CreateIncrementalReaderAsync(options => CsvFile.OpenTextReader(stream, options, leaveOpen: true, 4096),
            loadOptions, readerOptions, cancellationToken);
    }

    private static async Task<DbDataReader> CreateIncrementalReaderAsync(Func<CsvLoadOptions, TextReader> readerFactory,
        CsvLoadOptions? loadOptions, CsvDataReaderOptions? readerOptions, CancellationToken token)
    {
        readerOptions ??= new CsvDataReaderOptions();
        ValidateAsyncReaderOptions(readerOptions);
        var options = loadOptions?.Clone() ?? new CsvLoadOptions();
        if (options.DetectDelimiter)
            throw new NotSupportedException("Incremental reading currently requires an explicit delimiter.");
        if (readerOptions.ParallelProcessing is not null)
            throw new NotSupportedException("Incremental reading currently requires sequential projection.");
        int skip = GetInitialRecordsToSkip(options);
        var explicitHeader = NormalizeExplicitHeader(options);
        var lifetime = CancellationTokenSource.CreateLinkedTokenSource(token, options.CancellationToken);
        options.CancellationToken = lifetime.Token;
        options.Mode = CsvLoadMode.Stream;
        CsvParser.IncrementalRecords? records = null;
        CsvIncrementalRowSource? source = null;
        try
        {
            lifetime.Token.ThrowIfCancellationRequested();
            records = new CsvParser.IncrementalRecords(readerFactory(options), options);
            IReadOnlyList<string>? header = explicitHeader is null ? null : AppendStaticColumnsToHeader(explicitHeader, options);
            var buffered = new Queue<CsvIncrementalRowSource.BufferedRecord>();
            while (await records.ReadAsync(true, lifetime.Token).ConfigureAwait(false))
            {
                var record = records.Current;
                bool w3c = TryGetW3CFieldsHeader(record.Values, options, out var w3cHeader);
                if (explicitHeader is null && options.HasHeaderRow && options.SkipCommentRowsBeforeHeader && IsCommentRecord(record, options) && !w3c) continue;
                if (skip > 0) { skip--; continue; }
                if (header is null)
                {
                    header = AppendStaticColumnsToHeader(options.HasHeaderRow
                        ? NormalizeParsedHeader(w3c ? w3cHeader : record.Values, options)
                        : GenerateDefaultHeader(record.Values.Count), options);
                    if (options.HasHeaderRow) break;
                }
                buffered.Enqueue(new(record.Values, records.StartLine, records.EndLine));
                break;
            }
            header ??= Array.Empty<string>();
            CsvSchema? schema = readerOptions.Schema;
            if (schema is null && readerOptions.InferSchema)
            {
                while (buffered.Count < readerOptions.SchemaSampleSize &&
                    await records.ReadAsync(true, lifetime.Token).ConfigureAwait(false))
                    buffered.Enqueue(new(records.Current.Values, records.StartLine, records.EndLine));
                var columns = header.Select(name => new InferredColumn(name)).ToArray();
                foreach (var sample in buffered)
                {
                    var values = BuildParsedObjectValues(sample.Values, header.Count, options);
                    for (int i = 0; i < columns.Length; i++) columns[i].Observe(values[i], options.Culture, options.DateTimeFormats);
                }
                schema = new CsvSchema(columns.Select(column => column.ToSchemaColumn(buffered.Count)).ToArray());
            }
            source = new CsvIncrementalRowSource(records, options, lifetime, header.Count, buffered);
            var result = new CsvDataReader(CreateDataReaderColumns(header, schema), source, header.Count,
                options, options.Culture, options.DateTimeFormats);
            lifetime.Token.ThrowIfCancellationRequested();
            return result;
        }
        catch
        {
            if (source is not null) source.Dispose();
            else { try { records?.Dispose(); } finally { lifetime.Dispose(); } }
            throw;
        }
    }
}
#endif
