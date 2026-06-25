#nullable enable

using System.Text;

namespace OfficeIMO.CSV;

public sealed partial class CsvDocument
{
    /// <summary>
    /// Loads a CSV document from disk.
    /// </summary>
    public static CsvDocument Load(string path, CsvLoadOptions? options = null)
    {
        options ??= new CsvLoadOptions();
        var encoding = options.Encoding ?? new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
        return LoadInternal(() => new StreamReader(path, encoding, detectEncodingFromByteOrderMarks: true, bufferSize: FileBufferSize), options, encoding);
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

        options ??= new CsvLoadOptions();
        var encoding = options.Encoding ?? new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
        var readerFactory = () => new StreamReader(path, encoding, detectEncodingFromByteOrderMarks: true, bufferSize: FileBufferSize);
        var resolvedOptions = ResolveLoadOptions(readerFactory, options);
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

        options ??= new CsvLoadOptions();
        var encoding = options.Encoding ?? new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
        var readerFactory = () => new StreamReader(path, encoding, detectEncodingFromByteOrderMarks: true, bufferSize: FileBufferSize);
        var resolvedOptions = ResolveLoadOptions(readerFactory, options);
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

        var explicitHeader = NormalizeExplicitHeader(options);
        if (explicitHeader is not null)
        {
            CsvParser.ReadRecords(reader, options, record =>
            {
                InvokeRowAction(rowAction, explicitHeader, record, options.ColumnCountMismatchPolicy);
            });

            return;
        }

        IReadOnlyList<string>? header = null;
        if (options.HasHeaderRow)
        {
            CsvParser.ReadRecords(reader, options, record =>
            {
                if (header is null)
                {
                    if (TryGetW3CFieldsHeader(record, options, out var w3cHeader))
                    {
                        header = w3cHeader;
                        return;
                    }

                    if (options.SkipCommentRowsBeforeHeader && IsCommentRecord(record, options))
                    {
                        return;
                    }

                    header = NormalizeParsedHeader(record, options);
                    return;
                }

                InvokeRowAction(rowAction, header, record, options.ColumnCountMismatchPolicy);
            });

            return;
        }

        CsvParser.ReadRecords(reader, options, record =>
        {
            header ??= GenerateDefaultHeader(record.Length);
            InvokeRowAction(rowAction, header, record, options.ColumnCountMismatchPolicy);
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

        var explicitHeader = NormalizeExplicitHeader(options);
        if (explicitHeader is not null)
        {
            CsvParser.ReadRecordsReusable(reader, options, record =>
            {
                InvokeRowAction(rowAction, explicitHeader, record, options.ColumnCountMismatchPolicy);
            });

            return;
        }

        IReadOnlyList<string>? header = null;
        if (options.HasHeaderRow)
        {
            CsvParser.ReadRecordsReusable(reader, options, record =>
            {
                if (header is null)
                {
                    if (TryGetW3CFieldsHeader(record, options, out var w3cHeader))
                    {
                        header = w3cHeader;
                        return;
                    }

                    if (options.SkipCommentRowsBeforeHeader && IsCommentRecord(record, options))
                    {
                        return;
                    }

                    header = NormalizeParsedHeader(record, options);
                    return;
                }

                InvokeRowAction(rowAction, header, record, options.ColumnCountMismatchPolicy);
            });

            return;
        }

        CsvParser.ReadRecordsReusable(reader, options, record =>
        {
            header ??= GenerateDefaultHeader(record.Count);
            InvokeRowAction(rowAction, header, record, options.ColumnCountMismatchPolicy);
        });
    }

    /// <summary>
    /// Loads a CSV document from a stream.
    /// </summary>
    /// <param name="stream">Source stream.</param>
    /// <param name="options">Load options.</param>
    /// <param name="leaveOpen">Whether to leave the source stream open after loading.</param>
    public static CsvDocument Load(Stream stream, CsvLoadOptions? options = null, bool leaveOpen = true)
    {
        if (stream == null)
        {
            throw new ArgumentNullException(nameof(stream));
        }

        if (!stream.CanRead)
        {
            throw new ArgumentException("Stream must be readable.", nameof(stream));
        }

        options ??= new CsvLoadOptions();
        var encoding = options.Encoding ?? new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);

        if (options.Mode == CsvLoadMode.Stream || options.DetectDelimiter)
        {
            // Streaming mode and delimiter detection require a re-openable source for subsequent enumerations.
            // For arbitrary streams (including non-seekable), snapshot once into memory.
            var snapshot = ReadAllBytes(stream, leaveOpen);
            return LoadInternal(
                () => new StreamReader(
                    new MemoryStream(snapshot, writable: false),
                    encoding,
                    detectEncodingFromByteOrderMarks: true,
                    bufferSize: FileBufferSize,
                    leaveOpen: false),
                options,
                encoding);
        }

        return LoadInternal(
            () => new StreamReader(
                stream,
                encoding,
                detectEncodingFromByteOrderMarks: true,
                bufferSize: FileBufferSize,
                leaveOpen: leaveOpen),
            options,
            encoding);
    }

    /// <summary>
    /// Parses a CSV document from text.
    /// </summary>
    public static CsvDocument Parse(string text, CsvLoadOptions? options = null)
    {
        options ??= new CsvLoadOptions();
        var encoding = options.Encoding ?? new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
        return LoadInternal(() => new StringReader(text), options, encoding);
    }

    private static CsvDocument LoadInternal(Func<TextReader> readerFactory, CsvLoadOptions options, Encoding encoding)
    {
        options = ResolveLoadOptions(readerFactory, options);
        var document = new CsvDocument(options.Mode, options.Delimiter, options.Culture, encoding, options.ColumnCountMismatchPolicy);

        var explicitHeader = NormalizeExplicitHeader(options);
        if (explicitHeader is not null)
        {
            document.SetHeader(explicitHeader);
            if (options.Mode == CsvLoadMode.InMemory)
            {
                using var explicitHeaderReader = readerFactory();
                foreach (var record in CsvParser.Parse(explicitHeaderReader, options))
                {
                    document.AddParsedRowInternal(record, options.ColumnCountMismatchPolicy);
                }
            }
            else
            {
                document._streamingSource = new CsvStreamingSource(readerFactory, options, skipRecordCount: 0);
            }

            return document;
        }

        using var reader = readerFactory();
        using var enumerator = CsvParser.Parse(reader, options).GetEnumerator();

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
                    document.AddParsedRowInternal(enumerator.Current, options.ColumnCountMismatchPolicy);
                }
            }
            else
            {
                document._streamingSource = new CsvStreamingSource(readerFactory, options, consumedRecordCount);
            }

            return document;
        }

        if (!enumerator.MoveNext())
        {
            return document;
        }

        var firstRecord = enumerator.Current;
        document.SetHeader(GenerateDefaultHeader(firstRecord.Length));

        if (options.Mode == CsvLoadMode.InMemory)
        {
            document.AddParsedRowInternal(firstRecord, options.ColumnCountMismatchPolicy);
            while (enumerator.MoveNext())
            {
                document.AddParsedRowInternal(enumerator.Current, options.ColumnCountMismatchPolicy);
            }
        }
        else
        {
            document._streamingSource = new CsvStreamingSource(readerFactory, options, skipRecordCount: 0);
        }

        return document;
    }

    private static byte[] ReadAllBytes(Stream stream, bool leaveOpen)
    {
        try
        {
            using var ms = new MemoryStream();
            stream.CopyTo(ms);
            return ms.ToArray();
        }
        finally
        {
            if (!leaveOpen)
            {
                stream.Dispose();
            }
        }
    }

    private static void InvokeRowAction(
        Action<IReadOnlyList<string>, IReadOnlyList<string>> rowAction,
        IReadOnlyList<string> header,
        IReadOnlyList<string> values,
        CsvColumnCountMismatchPolicy policy)
    {
        rowAction(header, AlignParsedStringValues(values, header.Count, policy));
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
        IEnumerator<string[]> enumerator,
        CsvLoadOptions options,
        out IReadOnlyList<string> header,
        out int consumedRecordCount)
    {
        consumedRecordCount = 0;
        while (enumerator.MoveNext())
        {
            consumedRecordCount++;
            var record = enumerator.Current;
            if (TryGetW3CFieldsHeader(record, options, out var w3cHeader))
            {
                header = w3cHeader;
                return true;
            }

            if (options.SkipCommentRowsBeforeHeader && IsCommentRecord(record, options))
            {
                continue;
            }

            header = NormalizeParsedHeader(record, options);
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
            header = record.Skip(1).ToArray();
            return true;
        }

        return false;
    }

    private static bool IsCommentRecord(IReadOnlyList<string> record, CsvLoadOptions options) =>
        record.Count > 0 &&
        record[0].Length > 0 &&
        record[0][0] == options.CommentCharacter;
}
