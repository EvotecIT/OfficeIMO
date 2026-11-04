#nullable enable

namespace OfficeIMO.CSV;

public sealed partial class CsvDocument
{
    /// <summary>Opens a forward-only reader over decoded CSV text.</summary>
    /// <param name="reader">Text source, left open after opening and disposal.</param>
    /// <param name="options">Parsing and projection settings. Text is charged as UTF-8 against MaxInputBytes.</param>
    /// <returns>A CSV reader positioned before its first row.</returns>
    /// <remarks>The text is buffered during opening; use the Stream overload for byte-based input.</remarks>
    public static CsvDataReader OpenReader(TextReader reader, CsvReaderOptions? options = null)
    {
        if (reader is null) throw new ArgumentNullException(nameof(reader));
        options = options?.Clone() ?? new CsvReaderOptions();
        if (options.MaxInputBytes <= 0) throw new ArgumentOutOfRangeException(nameof(options));
        options.CancellationToken.ThrowIfCancellationRequested();
        if (reader is StringReader)
        {
            return CreateTextReader(reader.ReadToEnd(), options);
        }

        var text = new StringBuilder();
        var buffer = new char[4096];
        int count;
        while ((count = reader.Read(buffer, 0, buffer.Length)) != 0)
        {
            options.CancellationToken.ThrowIfCancellationRequested();
            if (text.Length + (long)count > options.MaxInputBytes)
                throw new InvalidDataException($"CSV data exceeds the configured maximum size ({options.MaxInputBytes} bytes).");
            text.Append(buffer, 0, count);
        }
        return CreateTextReader(text.ToString(), options);
    }

    /// <summary>
    /// Creates a forward-only data reader over already-decoded CSV text.
    /// </summary>
    /// <param name="text">Decoded CSV text.</param>
    /// <param name="options">Parsing and projection settings; decoded text is charged as UTF-8 against MaxInputBytes.</param>
    /// <returns>A data reader suitable for DataTable loading and provider bulk-copy APIs.</returns>
    internal static CsvDataReader CreateTextReader(
        string text,
        CsvReaderOptions? options = null)
    {
        if (text == null)
        {
            throw new ArgumentNullException(nameof(text));
        }

        options = options?.Clone() ?? new CsvReaderOptions();
        if (options.MaxInputBytes <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(options),
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
        if (options.SchemaSampleRows <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(options),
                "Schema sample size must be greater than zero.");
        }

#if NET8_0_OR_GREATER
        if (TryCreateHeaderlessTextDataReader(text, options, out CsvDataReader? textDataReader))
        {
            return CsvParallelDataReader.Apply(textDataReader!, options);
        }
#endif

        CsvDocument document = LoadInternal(
            () => new StringReader(text), options, utf8, text);
        return document.CreateDataReader(options);
    }

}
