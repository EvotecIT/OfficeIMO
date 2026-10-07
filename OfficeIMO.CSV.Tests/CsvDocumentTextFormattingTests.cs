using System;
using System.Globalization;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public class CsvDocumentTextFormattingTests
{
    [Theory]
    [InlineData("Default")]
    [InlineData("Semicolon")]
    [InlineData("Culture")]
    [InlineData("EmptyQuoteFields")]
    [InlineData("SelectedQuotes")]
    [InlineData("CustomNull")]
    public void ToString_Preserves_Typed_Values_And_Quoting(string profile)
    {
        var options = new CsvSaveOptions { NewLine = "\n" };
        string delimiter = ",";
        string firstNumber = "12.50";
        string secondNumber = "-2";
        string missing = string.Empty;
        switch (profile)
        {
            case "Semicolon": options.DelimiterText = delimiter = ";"; break;
            case "Culture": options.Culture = CultureInfo.GetCultureInfo("fr-FR"); firstNumber = "\"12,50\""; break;
            case "EmptyQuoteFields": options.QuoteFields = new[] { "", "  " }; break;
            case "SelectedQuotes":
                options.QuoteFields = new[] { "NUMBER" };
                firstNumber = "\"12.50\"";
                secondNumber = "\"-2\"";
                break;
            case "CustomNull": options.NullValue = "<missing, value>"; missing = "\"<missing, value>\""; break;
        }
        var document = new CsvDocument().WithHeader("Number", "Flag", "Text", "Missing")
            .AddRow(12.50m, true, "Łódź, \"quoted\"\nnext;end", null)
            .AddRow(-2, false, "plain", DBNull.Value);
        string firstHeader = profile == "SelectedQuotes" ? "\"Number\"" : "Number";
        string expected = firstHeader + delimiter + "Flag" + delimiter + "Text" + delimiter + "Missing\n"
            + firstNumber + delimiter + "True" + delimiter + "\"Łódź, \"\"quoted\"\"\nnext;end\"" + delimiter + missing + "\n"
            + secondNumber + delimiter + "False" + delimiter + "plain" + delimiter + missing + "\n";

        Assert.Equal(expected, document.ToString(options));
    }

    [Theory]
    [InlineData("", "12.50")]
    [InlineData("fr-FR", "12,50")]
    public void ToString_Quotes_Custom_Formatted_Values(string cultureName, string number)
    {
        var options = new CsvSaveOptions {
            IncludeHeader = false,
            NewLine = "\n",
            Culture = CultureInfo.GetCultureInfo(cultureName)
        };
        var document = new CsvDocument().AddRow(new ShortFormattedValue());

        Assert.Equal("\"" + number + ",\"\"quoted\"\"\"\n", document.ToString(options));
    }

    [Fact]
    public void ToString_Propagates_Custom_Formatting_Failure()
    {
        var document = new CsvDocument().AddRow(new ShortFormattedValue(throwOnFormat: true));

        var error = Assert.Throws<FormatException>(() => document.ToString());
        Assert.Equal("Cannot format this value.", error.Message);
    }

    private sealed class ShortFormattedValue :
#if NET6_0_OR_GREATER
        ISpanFormattable
#else
        IFormattable
#endif
    {
        private readonly bool _throwOnFormat;

        internal ShortFormattedValue(bool throwOnFormat = false) => _throwOnFormat = throwOnFormat;

        public string ToString(string? format, IFormatProvider? formatProvider)
        {
            if (_throwOnFormat) throw new FormatException("Cannot format this value.");
            return 12.50m.ToString(format, formatProvider) + ",\"quoted\"";
        }

        public override string ToString() => ToString(null, CultureInfo.InvariantCulture);

#if NET6_0_OR_GREATER
        public bool TryFormat(Span<char> destination, out int charsWritten,
            ReadOnlySpan<char> format, IFormatProvider? provider)
        {
            string text = ToString(format.IsEmpty ? null : format.ToString(), provider);
            bool copied = text.AsSpan().TryCopyTo(destination);
            charsWritten = copied ? text.Length : 0;
            return copied;
        }
#endif
    }
}
