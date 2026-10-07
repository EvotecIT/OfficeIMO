using System;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public class CsvAsyncFormattingContractTests
{
    [Theory]
    [InlineData("Default")]
    [InlineData("Semicolon")]
    [InlineData("TextDelimiter")]
    [InlineData("Culture")]
    [InlineData("AlwaysQuote")]
    [InlineData("NeverQuote")]
    [InlineData("QuoteFields")]
    [InlineData("EmptyQuoteFields")]
    [InlineData("CustomValues")]
    [InlineData("EscapeFormulas")]
    public async Task SaveAsync_Matches_Synchronous_Typed_Value_Formatting(string profile)
    {
        var options = new CsvSaveOptions { NewLine = "\r\n", Encoding = new UTF8Encoding(false, true) };
        switch (profile)
        {
            case "Semicolon": options.DelimiterText = ";"; break;
            case "TextDelimiter": options.DelimiterText = "||"; break;
            case "Culture": options.Culture = CultureInfo.GetCultureInfo("fr-FR"); break;
            case "AlwaysQuote": options.QuoteMode = CsvQuoteMode.Always; break;
            case "NeverQuote": options.QuoteMode = CsvQuoteMode.Never; break;
            case "QuoteFields": options.QuoteFields = new[] { "number", "NULL" }; break;
            case "EmptyQuoteFields": options.QuoteFields = new[] { "", "  " }; break;
            case "CustomValues":
                options.DateTimeFormat = "O";
                options.UseUtc = true;
                options.NullValue = "<missing, value>";
                break;
            case "EscapeFormulas": options.FormulaInjectionPolicy = CsvFormulaInjectionPolicy.Escape; break;
        }

        var document = new CsvDocument().WithHeader("Number", "Null", "Text");
        document.AddRow(12.50m, null, "Łódź 🚀 漢字, \"quoted\"\r\nnext;part||end");
        document.AddRow(-123, DBNull.Value, "=SUM(A1:A2)");
        document.AddRow(long.MinValue, true, Guid.Parse("d0ce49a7-fb04-489e-9478-82d07f355a24"));
        document.AddRow(ulong.MaxValue, false, new DateTime(2026, 10, 4, 12, 34, 56, DateTimeKind.Utc));
        document.AddRow(1.25e100, double.NaN, new DateTimeOffset(2026, 10, 4, 14, 34, 56, TimeSpan.FromHours(2)));
        document.AddRow(float.NegativeInfinity, (short)-12, TimeSpan.FromSeconds(123.45));
        document.AddRow((uint)12, (ushort)34, (byte)56);
        document.AddRow((sbyte)-78, '"', new Version(1, 2, 3, 4));
        document.AddRow(new LongFormattedValue(), string.Empty, "@formula");
#if NET6_0_OR_GREATER
        document.AddRow(new DateOnly(2026, 10, 4), new TimeOnly(12, 34, 56), "date-only");
#endif

        using var synchronous = new MemoryStream();
        document.Save(synchronous, options);
        using var asynchronous = new MemoryStream();
        await document.SaveAsync(asynchronous, options);

        Assert.Equal(Encoding.UTF8.GetBytes(document.ToString(options)), synchronous.ToArray());
        Assert.Equal(synchronous.ToArray(), asynchronous.ToArray());
    }

    private sealed class LongFormattedValue :
#if NET6_0_OR_GREATER
        ISpanFormattable
#else
        IFormattable
#endif
    {
        public string ToString(string? format, IFormatProvider? formatProvider) =>
            12.50m.ToString(format, formatProvider) + ",\"" + new string('x', 300) + "\"";

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
