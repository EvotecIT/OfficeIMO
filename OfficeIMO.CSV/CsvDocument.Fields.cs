#nullable enable

namespace OfficeIMO.CSV;

public sealed partial class CsvDocument
{
#if NET8_0_OR_GREATER
    /// <summary>
    /// Reads CSV fields from a file in a single pass without materializing unquoted fields as strings.
    /// </summary>
    /// <param name="path">Source CSV path.</param>
    /// <param name="fieldAction">Action receiving each field as a transient span.</param>
    /// <param name="options">Optional load settings. Header handling is not applied; records are emitted as parsed.</param>
    public static void ReadFieldSpans(string path, CsvFieldSpanAction fieldAction, CsvLoadOptions? options = null)
    {
        if (fieldAction == null)
        {
            throw new ArgumentNullException(nameof(fieldAction));
        }

        var visitor = new CsvFieldSpanActionVisitor(fieldAction);
        ReadFieldSpans(path, ref visitor, options);
    }

    /// <summary>
    /// Reads CSV fields from a file in a single pass without materializing unquoted fields as strings.
    /// </summary>
    /// <typeparam name="TVisitor">Struct visitor type receiving each field.</typeparam>
    /// <param name="path">Source CSV path.</param>
    /// <param name="fieldVisitor">Visitor receiving each field as a transient span.</param>
    /// <param name="options">Optional load settings. Header handling is not applied; records are emitted as parsed.</param>
    public static void ReadFieldSpans<TVisitor>(string path, ref TVisitor fieldVisitor, CsvLoadOptions? options = null)
        where TVisitor : struct, ICsvFieldSpanVisitor
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("File path cannot be empty.", nameof(path));
        }

        options = CreateRawRecordOptions(options);
        var readerFactory = () => CsvFile.OpenTextReader(path, options, FileBufferSize);
        var resolvedOptions = ResolveLoadOptions(readerFactory, options, useHeaderDiscoveryForDelimiterDetection: false);
        using var reader = readerFactory();
        ReadFieldSpans(reader, ref fieldVisitor, resolvedOptions);
    }

    /// <summary>
    /// Reads CSV fields from a reader in a single pass without materializing unquoted fields as strings.
    /// </summary>
    /// <param name="reader">Source text reader.</param>
    /// <param name="fieldAction">Action receiving each field as a transient span.</param>
    /// <param name="options">Optional load settings. Header handling is not applied; records are emitted as parsed.</param>
    public static void ReadFieldSpans(TextReader reader, CsvFieldSpanAction fieldAction, CsvLoadOptions? options = null)
    {
        if (fieldAction == null)
        {
            throw new ArgumentNullException(nameof(fieldAction));
        }

        var visitor = new CsvFieldSpanActionVisitor(fieldAction);
        ReadFieldSpans(reader, ref visitor, options);
    }

    /// <summary>
    /// Reads CSV fields from a reader in a single pass without materializing unquoted fields as strings.
    /// </summary>
    /// <typeparam name="TVisitor">Struct visitor type receiving each field.</typeparam>
    /// <param name="reader">Source text reader.</param>
    /// <param name="fieldVisitor">Visitor receiving each field as a transient span.</param>
    /// <param name="options">Optional load settings. Header handling is not applied; records are emitted as parsed.</param>
    public static void ReadFieldSpans<TVisitor>(TextReader reader, ref TVisitor fieldVisitor, CsvLoadOptions? options = null)
        where TVisitor : struct, ICsvFieldSpanVisitor
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
            using var bufferedReader = new StringReader(text);
            ReadFieldSpans(bufferedReader, ref fieldVisitor, resolvedOptions);
            return;
        }

        CsvParser.ReadFieldSpans(reader, options, GetInitialRecordsToSkip(options), ref fieldVisitor);
    }
#endif
}
