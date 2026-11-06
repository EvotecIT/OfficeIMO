using System.Globalization;
using OfficeIMO.Spreadsheet;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public class SpreadsheetNumberFormatDisplayContractsTests {
    [Theory]
    [InlineData("nl-NL", "1,23E+03")]
    [InlineData("en-US", "1.23E+03")]
    public void ScientificPayloadUsesProviderWithoutReplacingLiteralPunctuation(string culture, string expected) {
        Assert.Equal(".," + expected + ",.", SpreadsheetNumberFormatDisplay.FormatNumericValue(
            1234d, "\".,\"0.00E+00\",.\"", CultureInfo.GetCultureInfo(culture)));
    }

    [Theory]
    [InlineData("0.0#", "nl-NL", "12,5")]
    [InlineData("000.0#", "nl-NL", "012,5")]
    [InlineData("#,##0.00", "en-US", "12.50")]
    public void OptionalPrecisionAndIntegerPaddingUseProviderSeparators(string format, string culture, string expected) {
        Assert.Equal(expected, SpreadsheetNumberFormatDisplay.FormatNumericValue(12.5, format, CultureInfo.GetCultureInfo(culture)));
    }

    [Theory]
    [InlineData("0;0;0;[Unknown]@", "[Unknown]abc")]
    [InlineData("0;0;0;[ColorUnknown]@", "[ColorUnknown]abc")]
    [InlineData("0;0;0;\"@;\"@", "@;abc")]
    public void TextSectionRetainsUnknownSyntaxAndQuotedPlaceholders(string format, string expected) {
        Assert.Equal(expected, SpreadsheetNumberFormatDisplay.FormatTextValue("abc", format));
    }
}
