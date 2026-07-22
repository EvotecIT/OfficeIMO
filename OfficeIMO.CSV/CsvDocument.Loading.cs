#nullable enable

using OfficeIMO.Drawing.Internal;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.CSV;

public sealed partial class CsvDocument
{
    /// <summary>
    /// Loads a CSV document from disk.
    /// </summary>
    public static CsvDocument Load(string path, CsvLoadOptions? options = null)
    {
        var fileOptions = options?.Clone() ?? new CsvLoadOptions();
        var encoding = fileOptions.Encoding ?? new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
        return LoadInternal(() => CsvFile.OpenTextReader(path, fileOptions, FileBufferSize), fileOptions, encoding);
    }

    /// <summary>
    /// Reads a CSV file in a single pass and invokes an action for each data row.
    /// </summary>
    /// <param name="path">Source CSV path.</param>
    /// <param name="rowAction">Action receiving the header and current row values.</param>
    /// <param name="options">Optional load settings.</param>
    public static void ReadRows(string path, Action<IReadOnlyList<string>, IReadOnlyList<string>> rowAction, CsvLoadOptions? options = null)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("File path cannot be empty.", nameof(path));
        }

        if (rowAction == null)
        {
            throw new ArgumentNullException(nameof(rowAction));
        }

        var fileOptions = options?.Clone() ?? new CsvLoadOptions();
        var readerFactory = () => CsvFile.OpenTextReader(path, fileOptions, FileBufferSize);
        var resolvedOptions = ResolveLoadOptions(readerFactory, fileOptions);
        using var reader = readerFactory();
        ReadRows(reader, rowAction, resolvedOptions);
    }

    /// <summary>
    /// Reads a CSV file in a single pass while reusing the row value buffer for unquoted rows.
    /// </summary>
    /// <param name="path">Source CSV path.</param>
    /// <param name="rowAction">Action receiving the header and current row values. Row values must not be captured after the callback returns.</param>
    /// <param name="options">Optional load settings.</param>
    public static void ReadRowsReusable(string path, Action<IReadOnlyList<string>, IReadOnlyList<string>> rowAction, CsvLoadOptions? options = null)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("File path cannot be empty.", nameof(path));
        }

        if (rowAction == null)
        {
            throw new ArgumentNullException(nameof(rowAction));
        }

        var fileOptions = options?.Clone() ?? new CsvLoadOptions();
        var readerFactory = () => CsvFile.OpenTextReader(path, fileOptions, FileBufferSize);
        var resolvedOptions = ResolveLoadOptions(readerFactory, fileOptions);
        using var reader = readerFactory();
        ReadRowsReusable(reader, rowAction, resolvedOptions);
    }

    /// <summary>
    /// Reads CSV data in a single pass and invokes an action for each data row.
    /// </summary>
    /// <param name="reader">Source text reader.</param>
    /// <param name="rowAction">Action receiving the header and current row values.</param>
    /// <param name="options">Optional load settings.</param>
    public static void ReadRows(TextReader reader, Action<IReadOnlyList<string>, IReadOnlyList<string>> rowAction, CsvLoadOptions? options = null)
    {
        if (reader == null)
        {
            throw new ArgumentNullException(nameof(reader));
        }

        if (rowAction == null)
        {
            throw new ArgumentNullException(nameof(rowAction));
        }

        options ??= new CsvLoadOptions();
        if (options.DetectDelimiter)
        {
            var text = reader.ReadToEnd();
            var resolvedOptions = ResolveLoadOptions(() => new StringReader(text), options);
            using var bufferedReader = new StringReader(text);
            ReadRows(bufferedReader, rowAction, resolvedOptions);
            return;
        }

        var recordsToSkip = GetInitialRecordsToSkip(options);
        var explicitHeader = NormalizeExplicitHeader(options);
        if (explicitHeader is not null)
        {
            ReadRecordsSkippingInitial(reader, options, recordsToSkip, record =>
            {
                InvokeRowAction(rowAction, AppendStaticColumnsToHeader(explicitHeader, options), record, options);
            });

            return;
        }

        IReadOnlyList<string>? header = null;
        if (options.HasHeaderRow)
        {
            CsvParser.ReadRecordsWithMetadata(reader, options, record =>
            {
                if (header is null)
                {
                    var isW3CFieldsHeader = TryGetW3CFieldsHeader(record.Values, options, out var w3cHeader);
                    if (options.SkipCommentRowsBeforeHeader && IsCommentRecord(record, options) && !isW3CFieldsHeader)
                    {
                        return;
                    }

                    if (recordsToSkip > 0)
                    {
                        recordsToSkip--;
                        return;
                    }

                    if (isW3CFieldsHeader)
                    {
                        header = AppendStaticColumnsToHeader(NormalizeParsedHeader(w3cHeader, options), options);
                        return;
                    }

                    header = AppendStaticColumnsToHeader(NormalizeParsedHeader(record.Values, options), options);
                    return;
                }

                InvokeRowAction(rowAction, header, record.Values, options);
            });

            return;
        }

        ReadRecordsSkippingInitial(reader, options, recordsToSkip, record =>
        {
            header ??= AppendStaticColumnsToHeader(GenerateDefaultHeader(record.Length), options);
            InvokeRowAction(rowAction, header, record, options);
        });
    }

    /// <summary>
    /// Reads CSV data in a single pass while reusing the row value buffer for unquoted rows.
    /// </summary>
    /// <param name="reader">Source text reader.</param>
    /// <param name="rowAction">Action receiving the header and current row values. Row values must not be captured after the callback returns.</param>
    /// <param name="options">Optional load settings.</param>
    public static void ReadRowsReusable(TextReader reader, Action<IReadOnlyList<string>, IReadOnlyList<string>> rowAction, CsvLoadOptions? options = null)
    {
        if (reader == null)
        {
            throw new ArgumentNullException(nameof(reader));
        }

        if (rowAction == null)
        {
            throw new ArgumentNullException(nameof(rowAction));
        }

        options ??= new CsvLoadOptions();
        if (options.DetectDelimiter)
        {
            var text = reader.ReadToEnd();
            var resolvedOptions = ResolveLoadOptions(() => new StringReader(text), options);
            using var bufferedReader = new StringReader(text);
            ReadRowsReusable(bufferedReader, rowAction, resolvedOptions);
            return;
        }

        var recordsToSkip = GetInitialRecordsToSkip(options);
        var explicitHeader = NormalizeExplicitHeader(options);
        if (explicitHeader is not null)
        {
            ReadRecordsReusableSkippingInitial(reader, options, recordsToSkip, record =>
            {
                InvokeRowAction(rowAction, AppendStaticColumnsToHeader(explicitHeader, options), record, options);
            });

            return;
        }

        IReadOnlyList<string>? header = null;
        if (options.HasHeaderRow)
        {
            if (!options.SkipCommentRows)
            {
                CsvParser.ReadRecordsReusableWithMetadataUntilAccepted(
                    reader,
                    options,
                    record =>
                    {
                        var isW3CFieldsHeader = TryGetW3CFieldsHeader(record.Values, options, out var w3cHeader);
                        if (options.SkipCommentRowsBeforeHeader && IsCommentRecord(record, options) && !isW3CFieldsHeader)
                        {
                            return false;
                        }

                        if (recordsToSkip > 0)
                        {
                            recordsToSkip--;
                            return false;
                        }

                        header = isW3CFieldsHeader
                            ? AppendStaticColumnsToHeader(NormalizeParsedHeader(w3cHeader, options), options)
                            : AppendStaticColumnsToHeader(NormalizeParsedHeader(record.Values, options), options);
                        return true;
                    },
                    values => InvokeRowAction(rowAction, header!, values, options));

                return;
            }

            CsvParser.ReadRecordsReusableWithMetadata(reader, options, record =>
            {
                if (header is null)
                {
                    var isW3CFieldsHeader = TryGetW3CFieldsHeader(record.Values, options, out var w3cHeader);
                    if (options.SkipCommentRowsBeforeHeader && IsCommentRecord(record, options) && !isW3CFieldsHeader)
                    {
                        return;
                    }

                    if (recordsToSkip > 0)
                    {
                        recordsToSkip--;
                        return;
                    }

                    if (isW3CFieldsHeader)
                    {
                        header = AppendStaticColumnsToHeader(NormalizeParsedHeader(w3cHeader, options), options);
                        return;
                    }

                    header = AppendStaticColumnsToHeader(NormalizeParsedHeader(record.Values, options), options);
                    return;
                }

                InvokeRowAction(rowAction, header, record.Values, options);
            });

            return;
        }

        ReadRecordsReusableSkippingInitial(reader, options, recordsToSkip, record =>
        {
            header ??= AppendStaticColumnsToHeader(GenerateDefaultHeader(record.Count), options);
            InvokeRowAction(rowAction, header, record, options);
        });
    }

    /// <summary>
    /// Loads a CSV document from a stream.
    /// </summary>
    /// <param name="stream">Source stream.</param>
    /// <param name="options">Load options.</param>
    public static CsvDocument Load(Stream stream, CsvLoadOptions? options = null)
    {
        if (stream == null)
        {
            throw new ArgumentNullException(nameof(stream));
        }

        if (!stream.CanRead)
        {
            throw new ArgumentException("Stream must be readable.", nameof(stream));
        }

        options = options?.Clone() ?? new CsvLoadOptions();
        var encoding = options.Encoding ?? new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
        byte[] snapshot = OfficeStreamReader.ReadAllBytes(stream, options.MaxInputBytes);
        return LoadInternal(
            () => CsvFile.OpenTextReader(
                new MemoryStream(snapshot, writable: false),
                options,
                leaveOpen: false,
                FileBufferSize),
            options,
            encoding);
    }

    /// <summary>Asynchronously loads a CSV document from disk.</summary>
    public static async Task<CsvDocument> LoadAsync(string path, CsvLoadOptions? options = null, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        CsvLoadOptions resolved = options?.Clone() ?? new CsvLoadOptions();
        Encoding encoding = resolved.Encoding ?? new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
        using TextReader reader = CsvFile.OpenTextReaderForAsyncRead(path, resolved, FileBufferSize);
        string text = await ReadAllTextAsync(reader, cancellationToken).ConfigureAwait(false);
        return LoadInternal(() => new StringReader(text), resolved, encoding, text);
    }

    /// <summary>Asynchronously loads a CSV document from a caller-owned stream.</summary>
    public static async Task<CsvDocument> LoadAsync(Stream stream, CsvLoadOptions? options = null, CancellationToken cancellationToken = default)
    {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));
        CsvLoadOptions resolved = options?.Clone() ?? new CsvLoadOptions();
        Encoding encoding = resolved.Encoding ?? new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
        byte[] snapshot = await OfficeStreamReader.ReadAllBytesAsync(
            stream,
            cancellationToken,
            resolved.MaxInputBytes).ConfigureAwait(false);
        return LoadInternal(
            () => CsvFile.OpenTextReader(
                new MemoryStream(snapshot, writable: false),
                resolved,
                leaveOpen: false,
                FileBufferSize),
            resolved,
            encoding);
    }

    private static async Task<string> ReadAllTextAsync(TextReader reader, CancellationToken cancellationToken)
    {
        var text = new StringBuilder();
        var buffer = new char[FileBufferSize];
        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();
            int count = await reader.ReadAsync(buffer, 0, buffer.Length).ConfigureAwait(false);
            if (count == 0)
            {
                break;
            }

            text.Append(buffer, 0, count);
        }

        cancellationToken.ThrowIfCancellationRequested();
        return text.ToString();
    }

    /// <summary>
    /// Parses a CSV document from text.
    /// </summary>
    public static CsvDocument Parse(string text, CsvLoadOptions? options = null)
    {
        options = options?.Clone() ?? new CsvLoadOptions();
        var encoding = options.Encoding ?? new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
        return LoadInternal(() => new StringReader(text), options, encoding, text);
    }

    private static CsvDocument LoadInternal(Func<TextReader> readerFactory, CsvLoadOptions options, Encoding encoding, string? sourceText = null)
    {
        options = ResolveLoadOptions(readerFactory, options);
        var initialRecordsToSkip = GetInitialRecordsToSkip(options);
        var document = new CsvDocument(options.Mode, options.Delimiter, options.Culture, encoding, options.ColumnCountMismatchPolicy, options.DateTimeFormats);

        var explicitHeader = NormalizeExplicitHeader(options);
        if (explicitHeader is not null)
        {
            document.SetHeader(AppendStaticColumnsToHeader(explicitHeader, options));
            if (options.Mode == CsvLoadMode.InMemory)
            {
                using var explicitHeaderReader = readerFactory();
                var skipped = 0;
                foreach (var record in CsvParser.Parse(explicitHeaderReader, options))
                {
                    if (skipped < initialRecordsToSkip)
                    {
                        skipped++;
                        continue;
                    }

                    document.AddParsedRowInternal(record, options);
                }
            }
            else
            {
                document._streamingSource = new CsvStreamingSource(readerFactory, options, skipRecordCount: initialRecordsToSkip, document._header.Count, sourceText);
            }

            return document;
        }

        using var reader = readerFactory();
        using var enumerator = CsvParser.ParseWithMetadata(reader, options).GetEnumerator();

        if (options.HasHeaderRow)
        {
            if (!TryReadHeader(enumerator, options, out var header, out var consumedRecordCount))
            {
                return document;
            }

            document.SetHeader(header);

            if (options.Mode == CsvLoadMode.InMemory)
            {
                while (enumerator.MoveNext())
                {
                    document.AddParsedRowInternal(enumerator.Current.Values, options);
                }
            }
            else
            {
                document._streamingSource = new CsvStreamingSource(readerFactory, options, consumedRecordCount, document._header.Count, sourceText);
            }

            return document;
        }

        var skippedInitialRecords = 0;
        while (skippedInitialRecords < initialRecordsToSkip && enumerator.MoveNext())
        {
            skippedInitialRecords++;
        }

        if (!enumerator.MoveNext())
        {
            return document;
        }

        var firstRecord = enumerator.Current;
        document.SetHeader(AppendStaticColumnsToHeader(GenerateDefaultHeader(firstRecord.Values.Count), options));

        if (options.Mode == CsvLoadMode.InMemory)
        {
            document.AddParsedRowInternal(firstRecord.Values, options);
            while (enumerator.MoveNext())
            {
                document.AddParsedRowInternal(enumerator.Current.Values, options);
            }
        }
        else
        {
            document._streamingSource = new CsvStreamingSource(readerFactory, options, skipRecordCount: initialRecordsToSkip, document._header.Count, sourceText);
        }

        return document;
    }

    private static void InvokeRowAction(
        Action<IReadOnlyList<string>, IReadOnlyList<string>> rowAction,
        IReadOnlyList<string> header,
        IReadOnlyList<string> values,
        CsvLoadOptions options)
    {
        rowAction(header, BuildParsedStringValues(values, header.Count, options));
    }

    private static string[]? NormalizeExplicitHeader(CsvLoadOptions options)
    {
        if (options.Header is null)
        {
            return null;
        }

        if (options.Header.Length == 0)
        {
            throw new ArgumentException("Explicit header must contain at least one column.", nameof(options));
        }

        var header = new string[options.Header.Length];
        for (var i = 0; i < options.Header.Length; i++)
        {
            header[i] = options.Header[i] ?? string.Empty;
        }

        return NormalizeParsedHeader(header, options).ToArray();
    }

    private static bool TryReadHeader(
        IEnumerator<CsvParser.CsvParsedRecord> enumerator,
        CsvLoadOptions options,
        out IReadOnlyList<string> header,
        out int consumedRecordCount)
    {
        consumedRecordCount = 0;
        var initialRecordsToSkip = GetInitialRecordsToSkip(options);
        while (enumerator.MoveNext())
        {
            consumedRecordCount++;
            var record = enumerator.Current;
            var isW3CFieldsHeader = TryGetW3CFieldsHeader(record.Values, options, out var w3cHeader);

            if (options.SkipCommentRowsBeforeHeader && IsCommentRecord(record, options) && !isW3CFieldsHeader)
            {
                continue;
            }

            if (initialRecordsToSkip > 0)
            {
                initialRecordsToSkip--;
                continue;
            }

            if (isW3CFieldsHeader)
            {
                header = AppendStaticColumnsToHeader(NormalizeParsedHeader(w3cHeader, options), options);
                return true;
            }

            header = AppendStaticColumnsToHeader(NormalizeParsedHeader(record.Values, options), options);
            return true;
        }

        header = Array.Empty<string>();
        return false;
    }

    private static bool TryGetW3CFieldsHeader(IReadOnlyList<string> record, CsvLoadOptions options, out IReadOnlyList<string> header)
    {
        header = Array.Empty<string>();
        if (!options.RecognizeW3CFieldsHeader || record.Count == 0)
        {
            return false;
        }

        const string prefix = "#Fields:";
        if (record.Count == 1 && record[0].StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
        {
            var fields = record[0].Substring(prefix.Length)
                .Trim()
                .Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
            if (fields.Length > 0)
            {
                header = fields;
                return true;
            }
        }

        if (string.Equals(record[0], prefix, StringComparison.OrdinalIgnoreCase) && record.Count > 1)
        {
            var fields = record.Skip(1).Where(field => field.Length > 0).ToArray();
            if (fields.Length > 0)
            {
                header = fields;
                return true;
            }
        }

        return false;
    }

    private static int GetInitialRecordsToSkip(CsvLoadOptions options)
    {
        if (options.SkipInitialRecords < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(options), "SkipInitialRecords cannot be negative.");
        }

        return options.SkipInitialRecords;
    }

    private static void ReadRecordsSkippingInitial(TextReader reader, CsvLoadOptions options, int recordsToSkip, Action<string[]> recordAction)
    {
        if (recordsToSkip == 0)
        {
            CsvParser.ReadRecords(reader, options, recordAction);
            return;
        }

        CsvParser.ReadRecords(reader, options, record =>
        {
            if (recordsToSkip > 0)
            {
                recordsToSkip--;
                return;
            }

            recordAction(record);
        });
    }

    private static void ReadRecordsReusableSkippingInitial(TextReader reader, CsvLoadOptions options, int recordsToSkip, Action<IReadOnlyList<string>> recordAction)
    {
        if (recordsToSkip == 0)
        {
            CsvParser.ReadRecordsReusable(reader, options, recordAction);
            return;
        }

        CsvParser.ReadRecordsReusable(reader, options, record =>
        {
            if (recordsToSkip > 0)
            {
                recordsToSkip--;
                return;
            }

            recordAction(record);
        });
    }

    private static void ReadRecordsWithMetadataSkippingInitial(TextReader reader, CsvLoadOptions options, int recordsToSkip, Action<CsvParser.CsvParsedRecord> recordAction)
    {
        if (recordsToSkip == 0)
        {
            CsvParser.ReadRecordsWithMetadata(reader, options, recordAction);
            return;
        }

        CsvParser.ReadRecordsWithMetadata(reader, options, record =>
        {
            if (recordsToSkip > 0)
            {
                recordsToSkip--;
                return;
            }

            recordAction(record);
        });
    }

    private static void ReadRecordsReusableWithMetadataSkippingInitial(TextReader reader, CsvLoadOptions options, int recordsToSkip, Action<CsvParser.CsvParsedRecord> recordAction)
    {
        if (recordsToSkip == 0)
        {
            CsvParser.ReadRecordsReusableWithMetadata(reader, options, recordAction);
            return;
        }

        CsvParser.ReadRecordsReusableWithMetadata(reader, options, record =>
        {
            if (recordsToSkip > 0)
            {
                recordsToSkip--;
                return;
            }

            recordAction(record);
        });
    }

    private static bool IsCommentRecord(CsvParser.CsvParsedRecord record, CsvLoadOptions options) =>
        record.StartsWithCommentCharacter;
}
