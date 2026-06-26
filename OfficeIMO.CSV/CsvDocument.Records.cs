#nullable enable

using System.Text;

namespace OfficeIMO.CSV;

public sealed partial class CsvDocument
{
    /// <summary>
    /// Reads raw CSV records from a file as an enumerable sequence.
    /// </summary>
    /// <param name="path">Source CSV path.</param>
    /// <param name="options">Optional load settings. Header handling is not applied; records are emitted as parsed.</param>
    public static IEnumerable<string[]> ReadRecords(string path, CsvLoadOptions? options = null)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("File path cannot be empty.", nameof(path));
        }

        options = CreateRawRecordOptions(options);
        var encoding = options.Encoding ?? new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
        var readerFactory = () => new StreamReader(path, encoding, detectEncodingFromByteOrderMarks: true, bufferSize: FileBufferSize);
        var resolvedOptions = ResolveLoadOptions(readerFactory, options, useHeaderDiscoveryForDelimiterDetection: false);

        return ReadRecordsIterator(readerFactory, resolvedOptions, disposeReader: true);
    }

    /// <summary>
    /// Reads raw CSV records from a reader as an enumerable sequence.
    /// </summary>
    /// <param name="reader">Source text reader.</param>
    /// <param name="options">Optional load settings. Header handling is not applied; records are emitted as parsed.</param>
    public static IEnumerable<string[]> ReadRecords(TextReader reader, CsvLoadOptions? options = null)
    {
        if (reader == null)
        {
            throw new ArgumentNullException(nameof(reader));
        }

        options = CreateRawRecordOptions(options);
        if (options.DetectDelimiter)
        {
            var text = reader.ReadToEnd();
            var resolvedOptions = ResolveLoadOptions(() => new StringReader(text), options, useHeaderDiscoveryForDelimiterDetection: false);
            return ReadRecordsIterator(() => new StringReader(text), resolvedOptions, disposeReader: true);
        }

        return ReadRecordsIterator(() => reader, options, disposeReader: false);
    }

    /// <summary>
    /// Reads raw CSV records from a file in a single pass while reusing the record value buffer for unquoted records.
    /// </summary>
    /// <param name="path">Source CSV path.</param>
    /// <param name="recordAction">Action receiving the current record values. Record values must not be captured after the callback returns.</param>
    /// <param name="options">Optional load settings. Header handling is not applied; records are emitted as parsed.</param>
    public static void ReadRecordsReusable(string path, Action<IReadOnlyList<string>> recordAction, CsvLoadOptions? options = null)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("File path cannot be empty.", nameof(path));
        }

        if (recordAction == null)
        {
            throw new ArgumentNullException(nameof(recordAction));
        }

        options = CreateRawRecordOptions(options);
        var encoding = options.Encoding ?? new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
        var readerFactory = () => new StreamReader(path, encoding, detectEncodingFromByteOrderMarks: true, bufferSize: FileBufferSize);
        var resolvedOptions = ResolveLoadOptions(readerFactory, options, useHeaderDiscoveryForDelimiterDetection: false);
        using var reader = readerFactory();
        ReadRecordsReusable(reader, recordAction, resolvedOptions);
    }

    /// <summary>
    /// Reads raw CSV records from a reader in a single pass while reusing the record value buffer for unquoted records.
    /// </summary>
    /// <param name="reader">Source text reader.</param>
    /// <param name="recordAction">Action receiving the current record values. Record values must not be captured after the callback returns.</param>
    /// <param name="options">Optional load settings. Header handling is not applied; records are emitted as parsed.</param>
    public static void ReadRecordsReusable(TextReader reader, Action<IReadOnlyList<string>> recordAction, CsvLoadOptions? options = null)
    {
        if (reader == null)
        {
            throw new ArgumentNullException(nameof(reader));
        }

        if (recordAction == null)
        {
            throw new ArgumentNullException(nameof(recordAction));
        }

        options = CreateRawRecordOptions(options);
        if (options.DetectDelimiter)
        {
            var text = reader.ReadToEnd();
            var resolvedOptions = ResolveLoadOptions(() => new StringReader(text), options, useHeaderDiscoveryForDelimiterDetection: false);
            using var bufferedReader = new StringReader(text);
            ReadRecordsReusable(bufferedReader, recordAction, resolvedOptions);
            return;
        }

        ReadRecordsReusableSkippingInitial(reader, options, GetInitialRecordsToSkip(options), recordAction);
    }

    private static IEnumerable<string[]> ReadRecordsIterator(Func<TextReader> readerFactory, CsvLoadOptions options, bool disposeReader)
    {
        var reader = readerFactory();
        try
        {
            var recordsToSkip = GetInitialRecordsToSkip(options);
            foreach (var record in CsvParser.Parse(reader, options))
            {
                if (recordsToSkip > 0)
                {
                    recordsToSkip--;
                    continue;
                }

                yield return record;
            }
        }
        finally
        {
            if (disposeReader)
            {
                reader.Dispose();
            }
        }
    }

    private static CsvLoadOptions CreateRawRecordOptions(CsvLoadOptions? options)
    {
        var rawOptions = options?.Clone() ?? new CsvLoadOptions();
        rawOptions.HasHeaderRow = false;
        rawOptions.Header = null;
        rawOptions.SkipCommentRowsBeforeHeader = false;
        return rawOptions;
    }
}
