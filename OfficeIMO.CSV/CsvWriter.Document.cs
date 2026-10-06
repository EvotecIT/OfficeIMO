#nullable enable

using System.Globalization;
using System.Text;

namespace OfficeIMO.CSV;

internal static partial class CsvWriter
{
    public static void Write(TextWriter writer, CsvDocument document, CsvSaveOptions options)
    {
        var delimiter = GetDelimiterChar(options);
        var delimiterText = GetDelimiterText(options);
        var culture = options.Culture;
        var includeHeader = options.IncludeHeader;
        var newLine = options.NewLine;
        var formulaInjectionPolicy = options.FormulaInjectionPolicy;
        var quoteMode = options.QuoteMode;
        var quoteFields = CreateQuoteFieldSet(options.QuoteFields);

        if (includeHeader && document.Header.Count > 0)
        {
            if (delimiterText.Length == 1)
            {
                WriteRecord(writer, document.Header, delimiter, newLine, CultureInfo.InvariantCulture, formulaInjectionPolicy, quoteMode, quoteFields, document.Header);
            }
            else
            {
                WriteRecord(writer, document.Header, delimiterText, newLine, CultureInfo.InvariantCulture, formulaInjectionPolicy, quoteMode, quoteFields, document.Header);
            }
        }

        if (writer is StringWriter textWriter && delimiterText.Length == 1
            && options.NullValue == null && options.DateTimeFormat == null && !options.UseUtc
            && formulaInjectionPolicy == CsvFormulaInjectionPolicy.Preserve
            && quoteMode == CsvQuoteMode.AsNeeded && quoteFields == null)
        {
            StringBuilder output = textWriter.GetStringBuilder();
            foreach (var row in document.AsEnumerable())
                AppendRecordDefault(output, row.Values, delimiter, newLine, culture);
            return;
        }
        foreach (var row in document.AsEnumerable())
        {
            if (delimiterText.Length == 1)
            {
                WriteRecord(writer, row.Values, delimiter, newLine, culture, formulaInjectionPolicy, quoteMode, quoteFields, document.Header, options.DateTimeFormat, options.UseUtc, options.NullValue);
            }
            else
            {
                WriteRecord(writer, row.Values, delimiterText, newLine, culture, formulaInjectionPolicy, quoteMode, quoteFields, document.Header, options.DateTimeFormat, options.UseUtc, options.NullValue);
            }
        }
    }

}
