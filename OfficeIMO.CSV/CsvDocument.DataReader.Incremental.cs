#if NET8_0_OR_GREATER
using System.Data.Common;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.CSV;

public sealed partial class CsvDocument
{
    /// <summary>
    /// Opens a forward-only CSV reader whose ReadAsync calls perform incremental asynchronous I/O.
    /// Initialization reads the header and, when requested, at most SchemaSampleRows data records.
    /// Delimiter detection replays a bounded prefix. Dispose the reader to close its file.
    /// </summary>
    public static Task<CsvDataReader> OpenReaderAsync(string path,
        CsvReaderOptions? options = null,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        return CreateIncrementalReaderAsync(
            parsing => CsvFile.OpenTextReaderForAsyncRead(path, parsing,
                CsvLineReader.DefaultBufferSize, largeFileBufferSize: 128 * 1024),
            options, cancellationToken);
    }

    /// <summary>
    /// Opens an incremental CSV reader at the stream's current position. The stream remains open
    /// on failure and after disposal; its position advances and is not restored. Read and HasRows
    /// may perform synchronous I/O; use ReadAsync for asynchronous consumption. Cancellation or
    /// a parsing failure ends the reader. The opening token remains active for its lifetime.
    /// Delimiter detection samples at most 64 Ki characters and replays the prefix.
    /// </summary>
    public static Task<CsvDataReader> OpenReaderAsync(Stream stream,
        CsvReaderOptions? options = null,
        CancellationToken cancellationToken = default)
    {
        if (stream is null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));
        return CreateIncrementalReaderAsync(
            parsing => CsvFile.OpenTextReader(stream, parsing, leaveOpen: true, CsvLineReader.DefaultBufferSize),
            options, cancellationToken);
    }

    private static async Task<CsvDataReader> CreateIncrementalReaderAsync(Func<CsvReaderOptions, TextReader> readerFactory,
        CsvReaderOptions? readerOptions, CancellationToken token)
    {
        var options = readerOptions?.Clone() ?? new CsvReaderOptions();
        ValidateAsyncReaderOptions(options);
        options.ParallelProcessing?.GetDegreeOfParallelism();
        int skip = GetInitialRecordsToSkip(options);
        var explicitHeader = NormalizeExplicitHeader(options);
        CancellationToken loadToken = options.CancellationToken;
        var lifetime = CancellationTokenSource.CreateLinkedTokenSource(token, loadToken);
        options.CancellationToken = lifetime.Token;
        options.Mode = CsvLoadMode.Stream;
        CsvParser.IncrementalRecords? records = null;
        CsvIncrementalRowSource? source = null;
        TextReader? textReader = null;
        try
        {
            lifetime.Token.ThrowIfCancellationRequested();
            textReader = readerFactory(options);
            if (options.DetectDelimiter)
                textReader = await DetectIncrementalDelimiterAsync(textReader, options, lifetime.Token).ConfigureAwait(false);
            records = new CsvParser.IncrementalRecords(textReader, options);
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
            CsvSchema? schema = options.Schema;
            if (schema is null && options.InferSchema)
            {
                while (buffered.Count < options.SchemaSampleRows &&
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
            source = new CsvIncrementalRowSource(records, options, lifetime, token, loadToken, header.Count, buffered);
            var result = new CsvDataReader(CreateDataReaderColumns(header, schema), source, header.Count,
                options, options.Culture, options.DateTimeFormats);
            lifetime.Token.ThrowIfCancellationRequested();
            return CsvParallelDataReader.Apply(result, options);
        }
        catch
        {
            if (source is not null) source.Dispose();
            else { try { if (records is not null) records.Dispose(); else textReader?.Dispose(); } finally { lifetime.Dispose(); } }
            throw;
        }
    }
}
#endif
