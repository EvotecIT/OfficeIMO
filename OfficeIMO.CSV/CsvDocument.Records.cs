namespace OfficeIMO.CSV;

public sealed partial class CsvDocument
{
    /// <summary>
    /// Reads raw CSV records from a file as an enumerable sequence.
    /// </summary>
    /// <param name="path">Source CSV path.</param>
    /// <param name="options">Optional load settings. Header handling is not applied; records are emitted as parsed.</param>
    internal static IEnumerable<string[]> ReadRecords(string path, CsvParsingSettings? options = null)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("File path cannot be empty.", nameof(path));
        }

        options = CreateRawRecordOptions(options);
        var readerFactory = () => CsvFile.OpenTextReader(path, options, FileBufferSize);
        var resolvedOptions = ResolveParsingSettings(readerFactory, options, useHeaderDiscoveryForDelimiterDetection: false);

        return ReadRecordsIterator(readerFactory, resolvedOptions, disposeReader: true);
    }

    /// <summary>
    /// Reads raw CSV records from a reader as an enumerable sequence.
    /// </summary>
    /// <param name="reader">Source text reader.</param>
    /// <param name="options">Optional load settings. Header handling is not applied; records are emitted as parsed.</param>
    internal static IEnumerable<string[]> ReadRecords(TextReader reader, CsvParsingSettings? options = null)
    {
        if (reader == null)
        {
            throw new ArgumentNullException(nameof(reader));
        }

        options = CreateRawRecordOptions(options);
        if (options.DetectDelimiter)
        {
            var text = reader.ReadToEnd();
            var resolvedOptions = ResolveParsingSettings(() => new StringReader(text), options, useHeaderDiscoveryForDelimiterDetection: false);
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
    internal static void ReadRecordsReusable(string path, Action<IReadOnlyList<string>> recordAction, CsvParsingSettings? options = null)
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
        var readerFactory = () => CsvFile.OpenTextReader(path, options, FileBufferSize);
        var resolvedOptions = ResolveParsingSettings(readerFactory, options, useHeaderDiscoveryForDelimiterDetection: false);
        using var reader = readerFactory();
        ReadRecordsReusable(reader, recordAction, resolvedOptions);
    }

    /// <summary>
    /// Reads raw CSV records from a reader in a single pass while reusing the record value buffer for unquoted records.
    /// </summary>
    /// <param name="reader">Source text reader.</param>
    /// <param name="recordAction">Action receiving the current record values. Record values must not be captured after the callback returns.</param>
    /// <param name="options">Optional load settings. Header handling is not applied; records are emitted as parsed.</param>
    internal static void ReadRecordsReusable(TextReader reader, Action<IReadOnlyList<string>> recordAction, CsvParsingSettings? options = null)
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
            var resolvedOptions = ResolveParsingSettings(() => new StringReader(text), options, useHeaderDiscoveryForDelimiterDetection: false);
            using var bufferedReader = new StringReader(text);
            ReadRecordsReusable(bufferedReader, recordAction, resolvedOptions);
            return;
        }

        ReadRecordsReusableSkippingInitial(reader, options, GetInitialRecordsToSkip(options), recordAction);
    }

    private static IEnumerable<string[]> ReadRecordsIterator(Func<TextReader> readerFactory, CsvParsingSettings options, bool disposeReader)
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

    private static CsvParsingSettings CreateRawRecordOptions(CsvParsingSettings? options)
    {
        var rawOptions = options?.Clone() ?? new CsvParsingSettings();
        rawOptions.HasHeaderRow = false;
        rawOptions.Header = null;
        rawOptions.SkipCommentRowsBeforeHeader = false;
        return rawOptions;
    }
}
