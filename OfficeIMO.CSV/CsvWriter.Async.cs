#nullable enable
using System.Globalization;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.CSV;

internal static partial class CsvWriter
{
    // Reuses the canonical quoting/formatting code. Only one encoded record is
    // staged at a time, so asynchronous saving does not duplicate the whole file.
    internal static async Task WriteAsync(TextWriter writer, CsvDocument document,
        CsvSaveOptions options, CancellationToken cancellationToken, string? initialSeparator = null,
        Func<Task>? drainOutput = null)
    {
        string delimiter = GetDelimiterText(options);
        var quoteFields = CreateQuoteFieldSet(options.QuoteFields);
        bool defaultFormatting = delimiter.Length == 1 && options.NullValue == null
            && options.DateTimeFormat == null && !options.UseUtc
            && options.FormulaInjectionPolicy == CsvFormulaInjectionPolicy.Preserve
            && options.QuoteMode == CsvQuoteMode.AsNeeded && quoteFields == null;
        var buffer = new StringBuilder();
        using var recordWriter = new StringWriter(buffer, CultureInfo.InvariantCulture);

        async Task EmitAsync()
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (initialSeparator != null)
            {
#if NET8_0_OR_GREATER
                await writer.WriteAsync(initialSeparator.AsMemory(), cancellationToken).ConfigureAwait(false);
#else
                await writer.WriteAsync(initialSeparator).ConfigureAwait(false);
#endif
                initialSeparator = null;
            }
#if NET8_0_OR_GREATER
            foreach (ReadOnlyMemory<char> chunk in buffer.GetChunks())
                await writer.WriteAsync(chunk, cancellationToken).ConfigureAwait(false);
#else
            await writer.WriteAsync(buffer.ToString()).ConfigureAwait(false);
#endif
            buffer.Clear();
            if (drainOutput != null) await drainOutput().ConfigureAwait(false);
        }

        if (options.IncludeHeader && document.Header.Count > 0)
        {
            if (delimiter.Length == 1)
                WriteRecord(recordWriter, document.Header, delimiter[0], options.NewLine,
                    CultureInfo.InvariantCulture, options.FormulaInjectionPolicy, options.QuoteMode, quoteFields, document.Header);
            else
                WriteRecord(recordWriter, document.Header, delimiter, options.NewLine,
                    CultureInfo.InvariantCulture, options.FormulaInjectionPolicy, options.QuoteMode, quoteFields, document.Header);
            await EmitAsync().ConfigureAwait(false);
        }
        foreach (var row in document.AsEnumerable())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (defaultFormatting)
                AppendRecordDefault(buffer, row.Values, delimiter[0], options.NewLine, options.Culture);
            else if (delimiter.Length == 1)
                WriteRecord(recordWriter, row.Values, delimiter[0], options.NewLine, options.Culture,
                    options.FormulaInjectionPolicy, options.QuoteMode, quoteFields, document.Header, options.DateTimeFormat, options.UseUtc, options.NullValue);
            else
                WriteRecord(recordWriter, row.Values, delimiter, options.NewLine, options.Culture,
                    options.FormulaInjectionPolicy, options.QuoteMode, quoteFields, document.Header, options.DateTimeFormat, options.UseUtc, options.NullValue);
            await EmitAsync().ConfigureAwait(false);
        }
#if NET8_0_OR_GREATER
        await writer.FlushAsync(cancellationToken).ConfigureAwait(false);
#else
        cancellationToken.ThrowIfCancellationRequested();
        await writer.FlushAsync().ConfigureAwait(false);
#endif
    }
}
